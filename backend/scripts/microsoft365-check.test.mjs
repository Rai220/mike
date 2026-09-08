import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdtemp, writeFile, chmod, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { checkCredentials, checkMail, createOAuthHandler, loadConfig, safeFailure, SCOPES, REDIRECT_URI } from './microsoft365-check.mjs';

const config = { clientId: '11111111-1111-1111-1111-111111111111', tenantId: '22222222-2222-2222-2222-222222222222', clientSecret: 'test-secret+&= %~' };
const json = (body, status = 200) => new Response(JSON.stringify(body), { status });
const token = () => json({ access_token: 'PRIVATE_TOKEN', scope: SCOPES.join(' ') });
const mailFetch = (calls) => async (url, options) => {
  calls.push({ url, options });
  return calls.length === 1 ? token() : calls.length === 2 ? json({ id: 'PRIVATE_ACCOUNT' }) : json({ value: [{ id: 'PRIVATE_MAIL_ID', body: { content: 'PRIVATE_MAIL_BODY' } }] });
};
async function invoke(handler, url, overrides = {}) {
  const response = { once(event, callback) { if (event === 'finish') this.finished = callback; }, writeHead(status, headers) { this.status = status; this.headers = headers; }, end(body) { this.body = body; this.finished?.(); } };
  await handler({ method: 'GET', headers: { host: 'localhost:8400' }, url, ...overrides }, response);
  return response;
}

test('credentials encodes form fields and makes no Graph calls or token disclosure', async () => {
  const calls = [];
  const result = await checkCredentials(config, async (url, options) => { calls.push({ url, options }); return token(); });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].url, `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`);
  const form = new URLSearchParams(calls[0].options.body);
  assert.equal(form.get('client_secret'), config.clientSecret);
  assert.equal(form.get('grant_type'), 'client_credentials');
  assert.equal(form.get('scope'), 'https://graph.microsoft.com/.default');
  assert.equal(calls[0].options.redirect, 'error');
  assert.ok(calls[0].options.signal instanceof AbortSignal);
  assert.deepEqual(result, { ok: true, token_http_status: 200, credential_valid: true, mail_access_verified: false });
});

test('provider/network failures never expose arbitrary descriptions, tokens or secrets', async () => {
  for (const fetchImpl of [async () => json({ error_description: config.clientSecret, error_codes: [7000215, config.clientSecret] }, 401), async () => { throw new Error(config.clientSecret); }]) {
    await assert.rejects(checkCredentials(config, fetchImpl), error => {
      const safe = JSON.stringify(safeFailure(error));
      assert.ok(!safe.includes(config.clientSecret));
      assert.ok(!safe.includes('error_description'));
      return true;
    });
  }
  assert.deepEqual(safeFailure(new Error(config.clientSecret)), { ok: false, error: 'Microsoft 365 check failed' });
});

test('delegated callback validates state, exact scopes, PKCE; ignores invalid attempts and replay', async () => {
  const calls = [], reports = [], completed = [];
  let timestamp = 0;
  const handler = createOAuthHandler(config, { fetchImpl: mailFetch(calls), now: () => timestamp, report: r => reports.push(r), done: ok => completed.push(ok) });
  const start = await invoke(handler, '/start');
  const auth = new URL(start.headers.Location);
  const state = auth.searchParams.get('state');
  assert.equal(auth.origin, 'https://login.microsoftonline.com');
  assert.equal(auth.pathname, `/${config.tenantId}/oauth2/v2.0/authorize`);
  assert.equal(auth.searchParams.get('scope'), SCOPES.join(' '));
  assert.equal(auth.searchParams.get('redirect_uri'), REDIRECT_URI);
  assert.equal(auth.searchParams.get('code_challenge_method'), 'S256');
  assert.ok(!auth.href.includes('client_secret'));
  assert.ok(!auth.href.includes('code_verifier'));
  assert.ok(!auth.href.includes('offline_access'));
  const valid = `/microsoft365/callback?state=${state}&code=PRIVATE_CODE`;
  for (const url of ['/microsoft365/callback?code=PRIVATE_CODE', valid.replace(state, 'invalid'), `${valid}&state=${state}`, `${valid}&code=duplicate`, valid.replace('PRIVATE_CODE', ''), `${valid}&error=denied`]) {
    assert.equal((await invoke(handler, url)).status, 400);
  }
  assert.equal((await invoke(handler, valid, { method: 'POST' })).status, 400);
  assert.equal((await invoke(handler, valid, { headers: { host: 'evil.test' } })).status, 400);
  assert.equal(calls.length, 0);
  assert.equal(completed.length, 0);
  const response = await invoke(handler, valid);
  assert.equal(response.status, 200);
  assert.equal(response.headers['Cache-Control'], 'no-store');
  assert.equal(response.headers['Referrer-Policy'], 'no-referrer');
  const form = new URLSearchParams(calls[0].options.body);
  assert.equal(createHash('sha256').update(form.get('code_verifier')).digest('base64url'), auth.searchParams.get('code_challenge'));
  assert.equal(form.get('client_secret'), config.clientSecret);
  assert.equal(form.get('code'), 'PRIVATE_CODE');
  assert.equal(calls[1].url, 'https://graph.microsoft.com/v1.0/me?$select=id');
  assert.equal(calls[2].url, 'https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=id,body');
  assert.equal(calls[2].options.headers.Prefer, 'outlook.body-content-type="text"');
  assert.deepEqual(reports, [{ ok: true, granted_scopes: ['User.Read', 'Mail.Read'], token_http_status: 200, profile_http_status: 200, mail_http_status: 200, message_count: 1, body_read: true, mailbox_empty: false }]);
  assert.ok(!JSON.stringify({ reports, response }).includes('PRIVATE_'));
  assert.equal((await invoke(handler, valid)).status, 400);
  assert.equal(calls.length, 3);
  assert.deepEqual(completed, [true]);
  const expired = createOAuthHandler(config, { now: () => timestamp, fetchImpl: mailFetch(calls) });
  const oldState = new URL((await invoke(expired, '/start')).headers.Location).searchParams.get('state');
  timestamp = 600_000;
  assert.equal((await invoke(expired, `/microsoft365/callback?state=${oldState}&code=code`)).status, 400);
  assert.equal(calls.length, 3);
});

test('callbacks are consumed before awaiting token endpoint', async () => {
  let unblock, calls = 0;
  const handler = createOAuthHandler(config, { fetchImpl: async () => { calls++; return await new Promise(resolve => { unblock = resolve; }); } });
  const state = new URL((await invoke(handler, '/start')).headers.Location).searchParams.get('state');
  const path = `/microsoft365/callback?state=${state}&code=code`;
  const pending = invoke(handler, path);
  assert.equal((await invoke(handler, path)).status, 400);
  assert.equal(calls, 1);
  unblock(json({ error_description: 'PRIVATE' }, 400));
  await pending;
});

test('missing delegated permission prevents Graph requests; empty mailbox is distinct from body read', async () => {
  let calls = 0;
  await assert.rejects(checkMail(config, 'code', 'verifier', async () => { calls++; return json({ access_token: 'PRIVATE', scope: 'User.Read' }); }), error => safeFailure(error).reason === 'required_scopes_missing');
  assert.equal(calls, 1);
  calls = 0;
  const result = await checkMail(config, 'code', 'verifier', async () => ++calls === 1 ? json({ access_token: 'PRIVATE', scope: 'Mail.Read User.Read' }) : calls === 2 ? json({ id: 'PRIVATE' }) : json({ value: [] }));
  assert.equal(result.body_read, false);
  assert.equal(result.mailbox_empty, true);
  assert.equal(result.message_count, 0);
});

test('Graph failures only report allowlisted codes and HTTP status', async () => {
  for (const code of ['ErrorAccessDenied', 'PRIVATE_CODE']) {
    let calls = 0;
    await assert.rejects(checkMail(config, 'code', 'verifier', async () => ++calls === 1 ? token() : json({ error: { code, message: 'PRIVATE_MAIL_BODY' } }, 403)), error => {
      assert.deepEqual(safeFailure(error), { ok: false, error: 'Microsoft 365 check failed', stage: 'profile', http_status: 403, ...(code === 'ErrorAccessDenied' ? { graph_code: code } : {}) });
      return true;
    });
  }
});

test('local config requires private permissions and parses environment without logging it', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'm365-test-'));
  const file = join(dir, 'config');
  try {
    await writeFile(file, `MICROSOFT365_CLIENT_ID=${config.clientId}\nMICROSOFT365_TENANT_ID=${config.tenantId}\nMICROSOFT365_CLIENT_SECRET="${config.clientSecret}"\n`, { mode: 0o600 });
    assert.deepEqual(await loadConfig(file), config);
    await chmod(file, 0o644);
    await assert.rejects(loadConfig(file), error => safeFailure(error).reason === 'config_requires_mode_600');
  } finally { await rm(dir, { recursive: true, force: true }); }
});
