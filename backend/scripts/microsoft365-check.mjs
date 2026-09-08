// Temporary, read-only Microsoft 365 diagnostic. Never persists OAuth tokens or mail.
import { createHash, randomBytes, timingSafeEqual } from 'node:crypto';
import { readFile, stat } from 'node:fs/promises';
import { createServer } from 'node:http';
import { parseEnv } from 'node:util';
import { pathToFileURL } from 'node:url';

export const REDIRECT_URI = 'http://localhost:8400/microsoft365/callback';
export const SCOPES = ['https://graph.microsoft.com/User.Read', 'https://graph.microsoft.com/Mail.Read'];
const GRAPH = 'https://graph.microsoft.com/v1.0';
const LIFETIME = 600_000;
const GRAPH_ERRORS = new Set(['InvalidAuthenticationToken', 'Authorization_RequestDenied', 'ErrorAccessDenied', 'MailboxNotEnabledForRESTAPI', 'ErrorInvalidUser', 'TooManyRequests']);

class CheckError extends Error {
  constructor(details) { super('Microsoft 365 check failed'); this.details = details; }
}

export function safeFailure(error) {
  return { ok: false, error: 'Microsoft 365 check failed', ...(error instanceof CheckError ? error.details : {}) };
}

export async function loadConfig(file = new URL('../.env.microsoft365.local', import.meta.url)) {
  if (((await stat(file)).mode & 0o777) !== 0o600) throw new CheckError({ reason: 'config_requires_mode_600' });
  const env = parseEnv(await readFile(file, 'utf8'));
  const config = { clientId: env.MICROSOFT365_CLIENT_ID, tenantId: env.MICROSOFT365_TENANT_ID, clientSecret: env.MICROSOFT365_CLIENT_SECRET };
  const uuid = /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i;
  if (!uuid.test(config.clientId ?? '') || !uuid.test(config.tenantId ?? '') || !config.clientSecret) throw new CheckError({ reason: 'invalid_configuration' });
  return config;
}

async function request(fetchImpl, url, options, stage) {
  let response;
  try { response = await fetchImpl(url, { ...options, redirect: 'error', signal: AbortSignal.timeout(20_000) }); }
  catch { throw new CheckError({ stage, reason: 'network_failure' }); }
  let data;
  try { data = await response.json(); }
  catch { throw new CheckError({ stage, http_status: response.status, reason: 'invalid_response' }); }
  if (!response.ok) {
    const details = { stage, http_status: response.status };
    if (stage === 'token' && Array.isArray(data?.error_codes)) details.entra_codes = data.error_codes.filter(Number.isSafeInteger).slice(0, 10);
    if (GRAPH_ERRORS.has(data?.error?.code)) details.graph_code = data.error.code;
    throw new CheckError(details);
  }
  return { data, status: response.status };
}

function tokenRequest(config, fields, fetchImpl) {
  return request(fetchImpl, `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: config.clientId, client_secret: config.clientSecret, ...fields }).toString(),
  }, 'token');
}

export async function checkCredentials(config, fetchImpl = fetch) {
  const { data, status } = await tokenRequest(config, { grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' }, fetchImpl);
  if (typeof data?.access_token !== 'string' || !data.access_token) throw new CheckError({ reason: 'missing_access_token' });
  // An app token proves the credential works, not permission to access a mailbox.
  delete data.access_token;
  return { ok: true, token_http_status: status, credential_valid: true, mail_access_verified: false };
}

export async function checkMail(config, code, verifier, fetchImpl = fetch) {
  const token = await tokenRequest(config, { grant_type: 'authorization_code', code, redirect_uri: REDIRECT_URI, code_verifier: verifier, scope: SCOPES.join(' ') }, fetchImpl);
  let accessToken = token.data?.access_token;
  if (token.data) delete token.data.access_token;
  try {
    const granted = new Set(String(token.data?.scope ?? '').split(/\s+/).map(s => s.replace(/^https:\/\/graph\.microsoft\.com\//, '')));
    if (!['User.Read', 'Mail.Read'].every(s => granted.has(s))) throw new CheckError({ reason: 'required_scopes_missing' });
    if (typeof accessToken !== 'string' || !accessToken) throw new CheckError({ reason: 'missing_access_token' });
    const options = { headers: { Authorization: `Bearer ${accessToken}`, Prefer: 'outlook.body-content-type="text"' } };
    const me = await request(fetchImpl, `${GRAPH}/me?$select=id`, options, 'profile');
    if (typeof me.data?.id !== 'string' || !me.data.id) throw new CheckError({ reason: 'invalid_profile_response' });
    const mail = await request(fetchImpl, `${GRAPH}/me/messages?$top=1&$select=id,body`, options, 'mail');
    if (!Array.isArray(mail.data?.value)) throw new CheckError({ reason: 'invalid_mail_response' });
    const messages = mail.data.value;
    const bodyRead = messages.length > 0 && messages.every(m => typeof m?.body?.content === 'string');
    if (messages.length > 0 && !bodyRead) throw new CheckError({ reason: 'mail_body_missing' });
    return { ok: true, granted_scopes: ['User.Read', 'Mail.Read'], token_http_status: token.status,
      profile_http_status: me.status, mail_http_status: mail.status, message_count: messages.length,
      body_read: bodyRead, mailbox_empty: messages.length === 0 };
  } finally { accessToken = undefined; }
}

export function createOAuthHandler(config, { fetchImpl = fetch, now = Date.now, report = () => {}, done = () => {} } = {}) {
  let state = randomBytes(32).toString('base64url');
  let verifier = randomBytes(32).toString('base64url');
  const expires = now() + LIFETIME;
  let consumed = false;
  const authorize = new URL(`https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize`);
  authorize.search = new URLSearchParams({ client_id: config.clientId, response_type: 'code', response_mode: 'query',
    redirect_uri: REDIRECT_URI, scope: SCOPES.join(' '), state,
    code_challenge: createHash('sha256').update(verifier).digest('base64url'), code_challenge_method: 'S256' }).toString();
  return async (req, res) => {
    const reply = (status, text) => { res.writeHead(status, { 'Content-Type': 'text/plain; charset=utf-8', 'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer' }); res.end(text); };
    if (req.method !== 'GET' || req.headers.host !== 'localhost:8400') return reply(400, 'Invalid request.');
    let url;
    try { url = new URL(req.url, 'http://localhost:8400'); }
    catch { return reply(400, 'Invalid request.'); }
    if (url.origin !== 'http://localhost:8400' || !req.url.startsWith('/')) return reply(400, 'Invalid request.');
    if (url.pathname === '/start' && !url.search && !consumed && now() < expires) {
      res.writeHead(302, { Location: authorize.href, 'Cache-Control': 'no-store', 'Referrer-Policy': 'no-referrer' });
      return res.end();
    }
    if (url.pathname !== '/microsoft365/callback') return reply(404, 'Not found.');
    const states = url.searchParams.getAll('state');
    const codes = url.searchParams.getAll('code');
    const errors = url.searchParams.getAll('error');
    const supplied = Buffer.from(states[0] ?? '');
    const expected = Buffer.from(state);
    if (consumed || now() >= expires || states.length !== 1 || supplied.length !== expected.length || !timingSafeEqual(supplied, expected)
      || (errors.length === 0 ? codes.length !== 1 || !codes[0] || codes[0].length > 8192 : errors.length !== 1 || codes.length !== 0)) return reply(400, 'Invalid or expired callback.');
    consumed = true; // Consume synchronously, before exchanging the code.
    const codeVerifier = verifier;
    state = ''; verifier = '';
    let result;
    try {
      if (errors.length) throw new CheckError({ reason: 'authorization_denied' });
      result = await checkMail(config, codes[0], codeVerifier, fetchImpl);
    } catch (error) { result = safeFailure(error); }
    report(result);
    res.once('finish', () => done(result.ok));
    reply(result.ok ? 200 : 400, result.ok ? 'Check complete. You may close this window.' : 'Check failed. See the terminal for safe diagnostic details.');
  };
}

async function main() {
  if (!['credentials', 'mail'].includes(process.argv[2]) || process.argv.length !== 3) throw new CheckError({ reason: 'usage_credentials_or_mail' });
  const config = await loadConfig();
  if (process.argv[2] === 'credentials') return console.log(JSON.stringify(await checkCredentials(config)));
  let timer;
  const stop = (ok) => {
    clearTimeout(timer);
    process.removeListener('SIGINT', interrupted);
    process.removeListener('SIGTERM', interrupted);
    process.exitCode = ok ? 0 : 1;
    server.close();
    server.closeAllConnections();
  };
  const interrupted = () => stop(false);
  const server = createServer(createOAuthHandler(config, { report: result => console.log(JSON.stringify(result)), done: stop }));
  server.requestTimeout = 30_000;
  server.headersTimeout = 10_000;
  server.on('error', () => { console.log(JSON.stringify(safeFailure(new CheckError({ reason: 'local_server_failed' })))); stop(false); });
  process.once('SIGINT', interrupted);
  process.once('SIGTERM', interrupted);
  timer = setTimeout(() => { console.log(JSON.stringify(safeFailure(new CheckError({ reason: 'authorization_timeout' })))); stop(false); }, LIFETIME);
  server.listen(8400, '127.0.0.1', () => console.log('Open http://localhost:8400/start in your browser.'));
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch(error => { console.log(JSON.stringify(safeFailure(error))); process.exitCode = 1; });
}
