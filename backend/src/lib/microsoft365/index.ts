import crypto from "node:crypto";
import type { createServerSupabase } from "../supabase";
import { getOrgRole } from "../access";

type Db = ReturnType<typeof createServerSupabase>;
const TABLE = "user_microsoft365_connections";
const STATES = "microsoft365_oauth_states";
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
export const MICROSOFT365_SCOPES = [
  "offline_access",
  "User.Read",
  "Mail.Read",
] as const;
export type ErrorCode =
  | "invalid_source"
  | "source_too_large"
  | "unsupported_source"
  | "source_changed"
  | "invalid_query"
  | "unavailable"
  | "access_denied"
  | "not_found"
  | "invalid_state"
  | "account_mismatch"
  | "reconnect_required"
  | "busy"
  | "rate_limited"
  | "provider_unavailable"
  | "storage_unavailable";
const ERROR_DETAILS: Record<ErrorCode, string> = {
  invalid_source: "This source is not available in this chat.",
  source_too_large: "This source exceeds the supported size limit.",
  unsupported_source: "This source format is not supported or is protected.",
  source_changed:
    "A source changed. Start a new chat with its current version.",
  invalid_query: "Please provide a valid search or message within the limits.",
  unavailable: "Microsoft 365 connection is not available.",
  access_denied:
    "You do not have permission to access this Microsoft 365 connection.",
  not_found: "Microsoft 365 connection was not found.",
  invalid_state:
    "This Microsoft sign-in request expired or was already used. Please start again.",
  account_mismatch:
    "Disconnect the existing Microsoft account before connecting a different account.",
  reconnect_required: "Please reconnect your Microsoft 365 account.",
  busy: "Another Microsoft 365 connection operation is in progress. Please try again shortly.",
  rate_limited:
    "Microsoft is temporarily limiting requests. Please try again shortly.",
  provider_unavailable:
    "Microsoft 365 could not be reached. Please try again later.",
  storage_unavailable:
    "Microsoft 365 connection could not be saved or loaded. Please try again later.",
};
export class Microsoft365Error extends Error {
  readonly detail: string;
  constructor(
    readonly code: ErrorCode,
    readonly status: number,
    readonly retryAfterSeconds?: number,
  ) {
    super(code);
    this.name = "Microsoft365Error";
    this.detail = ERROR_DETAILS[code];
  }
}
function fail(code: ErrorCode, status = 400): never {
  throw new Microsoft365Error(code, status);
}

export interface Microsoft365Config {
  clientId: string;
  tenantId: string;
  clientSecret: string;
  orgId: string;
  encryptionSecret: string;
  callbackUrl: string;
  frontendOrigin: string;
}
/** No host/header-derived redirects, shared encryption fallback, or implicit enablement. */
export function getMicrosoft365Config(): Microsoft365Config | null {
  if (process.env.MICROSOFT365_ENABLED !== "true") return null;
  const clientId = process.env.MICROSOFT365_CLIENT_ID?.trim() ?? "";
  const tenantId = process.env.MICROSOFT365_TENANT_ID?.trim() ?? "";
  const orgId = process.env.MICROSOFT365_ORG_ID?.trim() ?? "";
  const clientSecret = process.env.MICROSOFT365_CLIENT_SECRET ?? "";
  const encryptionSecret = process.env.MICROSOFT365_ENCRYPTION_SECRET ?? "";
  if (
    ![clientId, tenantId, orgId].every((v) => UUID.test(v)) ||
    !clientSecret ||
    encryptionSecret.trim().length < 32
  )
    return null;
  try {
    const api = new URL(process.env.API_PUBLIC_URL ?? "");
    const frontend = new URL(process.env.FRONTEND_URL ?? "");
    for (const url of [api, frontend]) {
      const local =
        process.env.NODE_ENV !== "production" &&
        ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname);
      if (
        (url.protocol !== "https:" && !(local && url.protocol === "http:")) ||
        url.username ||
        url.password ||
        url.search ||
        url.hash
      )
        return null;
    }
    if (frontend.pathname !== "/") return null;
    return {
      clientId: clientId.toLowerCase(),
      tenantId: tenantId.toLowerCase(),
      orgId: orgId.toLowerCase(),
      clientSecret,
      encryptionSecret,
      callbackUrl: `${api.href.replace(/\/+$/, "")}/integrations/microsoft365/callback`,
      frontendOrigin: frontend.origin,
    };
  } catch {
    return null;
  }
}
function microsoft365Scopes(): string[] {
  return [
    ...MICROSOFT365_SCOPES,
    ...(process.env.MICROSOFT365_FILES_ENABLED === "true"
      ? ["Files.Read.All"]
      : []),
  ];
}
function config(): Microsoft365Config {
  return getMicrosoft365Config() ?? fail("unavailable", 503);
}
async function member(userId: string, c: Microsoft365Config, db: Db) {
  if (!UUID.test(userId) || !(await getOrgRole(userId, c.orgId, db)))
    fail("access_denied", 403);
}
interface Row {
  id: string;
  user_id: string;
  org_id: string;
  tenant_id: string;
  client_id: string;
  status: "pending" | "connected" | "reconnect_required";
  account_id: string | null;
  account_label: string | null;
  connected_at: string | null;
  last_checked_at: string | null;
  token_ciphertext: string | null;
  token_expires_at: string | null;
  version: string;
  oauth_generation: string;
  enabled: boolean;
}
interface Tokens {
  accessToken: string;
  refreshToken: string;
  scopes: string[];
}
const hash = (value: string) =>
  crypto.createHash("sha256").update(value).digest("hex");
/** Called only with the JWT already verified by requireAuth. Session ID survives token refresh. */
function sessionHash(userId: string, token: string): string {
  try {
    const payload = JSON.parse(
      Buffer.from(token.split(".")[1], "base64url").toString("utf8"),
    ) as { session_id?: unknown };
    if (typeof payload.session_id === "string" && UUID.test(payload.session_id))
      return hash(`${userId}:${payload.session_id}`);
  } catch {
    /* Legacy verified tokens bind to their exact bytes and fail closed on refresh. */
  }
  return hash(`${userId}:${token}`);
}
function context(
  row: Pick<Row, "id" | "user_id" | "org_id" | "tenant_id" | "client_id">,
  purpose: string,
) {
  return JSON.stringify([
    "mike-microsoft365-v1",
    purpose,
    row.id,
    row.user_id,
    row.org_id,
    row.tenant_id,
    row.client_id,
  ]);
}
function encryptionKeys(c: Microsoft365Config) {
  const active = process.env.MICROSOFT365_ENCRYPTION_KEY_ID || "legacy";
  if (!/^[a-zA-Z0-9_-]{1,32}$/.test(active)) fail("unavailable", 503);
  let previous: Record<string, string> = {};
  try {
    previous = JSON.parse(
      process.env.MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS || "{}",
    );
    if (
      !previous ||
      Array.isArray(previous) ||
      typeof previous !== "object" ||
      Object.keys(previous).length > 8 ||
      Object.entries(previous).some(
        ([id, key]) =>
          !/^[a-zA-Z0-9_-]{1,32}$/.test(id) ||
          typeof key !== "string" ||
          key.length < 32,
      )
    )
      throw new Error();
  } catch {
    fail("unavailable", 503);
  }
  return { active, keys: { ...previous, [active]: c.encryptionSecret } };
}
function encrypt(value: string, aad: string, c: Microsoft365Config): string {
  const { active, keys } = encryptionKeys(c);
  const key = crypto.scryptSync(keys[active], "mike-microsoft365-v1", 32);
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
  cipher.setAAD(Buffer.from(aad));
  const encrypted = Buffer.concat([
    cipher.update(value, "utf8"),
    cipher.final(),
  ]);
  return [
    "v2",
    active,
    ...[iv, cipher.getAuthTag(), encrypted].map((v) => v.toString("base64url")),
  ].join(".");
}
function decrypt(value: string, aad: string, c: Microsoft365Config): string {
  try {
    const parts = value.split(".");
    const id =
      parts.length === 3
        ? "legacy"
        : parts[0] === "v2" && parts.length === 5
          ? parts[1]
          : "";
    const secret = encryptionKeys(c).keys[id];
    if (!secret) fail("reconnect_required", 409);
    const [iv, tag, bytes] = (parts.length === 3 ? parts : parts.slice(2)).map(
      (v) => Buffer.from(v, "base64url"),
    );
    const decipher = crypto.createDecipheriv(
      "aes-256-gcm",
      crypto.scryptSync(secret, "mike-microsoft365-v1", 32),
      iv,
    );
    decipher.setAAD(Buffer.from(aad));
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(bytes), decipher.final()]).toString(
      "utf8",
    );
  } catch {
    return fail("reconnect_required", 409);
  }
}
function checked<T extends { error: unknown }>(result: T): T {
  if (result.error) fail("storage_unavailable", 503);
  return result;
}
function matches(row: Row, c: Microsoft365Config) {
  return (
    row.enabled &&
    row.org_id === c.orgId &&
    row.tenant_id === c.tenantId &&
    row.client_id === c.clientId
  );
}
async function owned(userId: string, id: string, db: Db): Promise<Row> {
  if (!UUID.test(id)) fail("not_found", 404);
  const { data } = checked(
    await db
      .from(TABLE)
      .select("*")
      .eq("id", id)
      .eq("user_id", userId)
      .maybeSingle(),
  );
  return (data as Row) ?? fail("not_found", 404);
}
async function authorized(
  userId: string,
  id: string,
  c: Microsoft365Config,
  db: Db,
) {
  await member(userId, c, db);
  const row = await owned(userId, id, db);
  if (!matches(row, c)) fail("access_denied", 403);
  return row;
}
function publicRow(row: Row, showLabel: boolean) {
  return {
    id: row.id,
    status: row.status,
    accountLabel: showLabel ? row.account_label : null,
    connectedAt: row.connected_at,
    lastCheckedAt: row.last_checked_at,
  };
}
export async function getMicrosoft365Status(userId: string, db: Db) {
  const c = getMicrosoft365Config();
  const allowed = !!c && !!(await getOrgRole(userId, c.orgId, db));
  const query = () =>
    db
      .from(TABLE)
      .select(
        "id,user_id,org_id,tenant_id,client_id,enabled,status,account_label,connected_at,last_checked_at",
      )
      .eq("user_id", userId);
  let result = c
    ? await query()
        .eq("org_id", c.orgId)
        .eq("tenant_id", c.tenantId)
        .eq("client_id", c.clientId)
        .maybeSingle()
    : await query()
        .order("created_at", { ascending: false })
        .limit(1)
        .maybeSingle();
  // Existing deployments can keep this opt-in feature disabled before migrating.
  if (
    process.env.MICROSOFT365_ENABLED !== "true" &&
    result.error?.code === "42P01"
  )
    return { available: false, connection: null };
  checked(result);
  if (c && !result.data)
    result = checked(
      await query()
        .order("created_at", { ascending: false })
        .limit(1)
        .maybeSingle(),
    );
  const row = result.data as Row | null;
  return {
    available: allowed,
    connection: row ? publicRow(row, allowed && !!c && matches(row, c)) : null,
  };
}
export async function startMicrosoft365Connection(
  userId: string,
  sessionToken: string,
  db: Db,
) {
  const c = config();
  await member(userId, c, db);
  if (!sessionToken) fail("invalid_state");
  const { data: existing } = checked(
    await db
      .from(TABLE)
      .select("*")
      .eq("user_id", userId)
      .eq("org_id", c.orgId)
      .maybeSingle(),
  );
  if (existing && !matches(existing as Row, c)) fail("reconnect_required", 409);
  const generation = crypto.randomUUID();
  const row = existing as Row | null;
  const values = {
    oauth_generation: generation,
    version: crypto.randomUUID(),
    refresh_lease_id: null,
    refresh_lease_expires_at: null,
  };
  const result = row
    ? await db
        .from(TABLE)
        .update(values)
        .eq("id", row.id)
        .eq("user_id", userId)
        .eq("version", row.version)
        .or(
          `refresh_lease_expires_at.is.null,refresh_lease_expires_at.lt.${new Date().toISOString()}`,
        )
        .select("*")
        .maybeSingle()
    : await db
        .from(TABLE)
        .insert({
          ...values,
          user_id: userId,
          org_id: c.orgId,
          tenant_id: c.tenantId,
          client_id: c.clientId,
        })
        .select("*")
        .single();
  const connection = checked(result).data as Row | null;
  if (!connection) fail("busy", 409);
  const state = crypto.randomBytes(32).toString("base64url");
  const verifier = crypto.randomBytes(32).toString("base64url");
  // Expired state secrets are pruned on each new authorization attempt.
  checked(
    await db
      .from(STATES)
      .delete()
      .eq("user_id", userId)
      .lt("expires_at", new Date().toISOString()),
  );
  checked(
    await db.from(STATES).insert({
      state_hash: hash(state),
      session_hash: sessionHash(userId, sessionToken),
      connection_id: connection.id,
      user_id: userId,
      org_id: c.orgId,
      tenant_id: c.tenantId,
      client_id: c.clientId,
      generation,
      verifier_ciphertext: encrypt(
        verifier,
        context(connection, `pkce:${generation}`),
        c,
      ),
      expires_at: new Date(Date.now() + 10 * 60_000).toISOString(),
    }),
  );
  const url = new URL(
    `https://login.microsoftonline.com/${c.tenantId}/oauth2/v2.0/authorize`,
  );
  url.search = new URLSearchParams({
    client_id: c.clientId,
    response_type: "code",
    redirect_uri: c.callbackUrl,
    response_mode: "query",
    scope: microsoft365Scopes().join(" "),
    state,
    code_challenge: crypto
      .createHash("sha256")
      .update(verifier)
      .digest("base64url"),
    code_challenge_method: "S256",
    prompt: "select_account",
  }).toString();
  return {
    authorizationUrl: url.toString(),
    callbackOrigin: new URL(c.callbackUrl).origin,
  };
}
async function request(url: string, init: RequestInit): Promise<Response> {
  try {
    return await fetch(url, {
      ...init,
      redirect: "error",
      signal: AbortSignal.timeout(15_000),
    });
  } catch {
    return fail("provider_unavailable", 502);
  }
}
function responseError(response: Response): never {
  if (response.status === 401) fail("reconnect_required", 409);
  if (response.status === 403) fail("access_denied", 403);
  if (response.status === 429) {
    const retry = response.headers.get("retry-after");
    const seconds =
      retry && /^\d+$/.test(retry)
        ? Number(retry)
        : retry
          ? Math.ceil((Date.parse(retry) - Date.now()) / 1000)
          : 30;
    throw new Microsoft365Error(
      "rate_limited",
      429,
      Number.isFinite(seconds) && seconds > 0
        ? Math.min(300, Math.ceil(seconds))
        : 30,
    );
  }
  return fail("provider_unavailable", 502);
}
async function json(response: Response): Promise<Record<string, unknown>> {
  try {
    const value: unknown = await response.json();
    if (value && typeof value === "object" && !Array.isArray(value))
      return value as Record<string, unknown>;
  } catch {
    /* Provider payloads must never become client errors. */
  }
  return fail("provider_unavailable", 502);
}
async function token(
  c: Microsoft365Config,
  params: Record<string, string>,
  oldRefresh?: string,
) {
  const response = await request(
    `https://login.microsoftonline.com/${c.tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: c.clientId,
        client_secret: c.clientSecret,
        scope: microsoft365Scopes().join(" "),
        ...params,
      }).toString(),
    },
  );
  if (!response.ok) {
    if (response.status === 400) {
      const body = await json(response);
      if (
        ["invalid_grant", "interaction_required", "consent_required"].includes(
          String(body.error),
        )
      )
        fail("reconnect_required", 409);
    }
    responseError(response);
  }
  const body = await json(response);
  const scopes = typeof body.scope === "string" ? body.scope.split(/\s+/) : [];
  const normalized = scopes.map((scope) =>
    scope.replace(/^https:\/\/graph.microsoft.com\//i, "").toLowerCase(),
  );
  const refreshToken =
    typeof body.refresh_token === "string" ? body.refresh_token : oldRefresh;
  if (
    typeof body.access_token !== "string" ||
    !body.access_token ||
    !refreshToken ||
    String(body.token_type).toLowerCase() !== "bearer" ||
    !normalized.includes("mail.read") ||
    !normalized.includes("user.read")
  )
    fail("reconnect_required", 409);
  const expires = Number(body.expires_in);
  if (!Number.isFinite(expires) || expires < 60 || expires > 86400)
    fail("provider_unavailable", 502);
  return {
    tokens: { accessToken: body.access_token, refreshToken, scopes },
    expiresAt: new Date(Date.now() + expires * 1000).toISOString(),
  };
}
async function profile(accessToken: string) {
  const response = await request(
    "https://graph.microsoft.com/v1.0/me?$select=id,mail,userPrincipalName,displayName",
    { headers: { Authorization: `Bearer ${accessToken}` } },
  );
  if (!response.ok) responseError(response);
  const body = await json(response);
  if (typeof body.id !== "string" || !UUID.test(body.id))
    fail("provider_unavailable", 502);
  const label = [body.mail, body.userPrincipalName, body.displayName].find(
    (v) => typeof v === "string" && v.trim(),
  );
  return {
    id: body.id.toLowerCase(),
    label: typeof label === "string" ? label.slice(0, 320) : null,
  };
}
export async function completeMicrosoft365Connection(
  userId: string,
  sessionToken: string,
  input: { state: string; code?: string; error?: string },
  db: Db,
): Promise<void> {
  const c = config();
  await member(userId, c, db);
  if (!sessionToken || !/^[A-Za-z0-9_-]{43}$/.test(input.state))
    fail("invalid_state");
  // DELETE RETURNING consumes state atomically only for its authenticated session.
  const { data } = checked(
    await db
      .from(STATES)
      .delete()
      .eq("state_hash", hash(input.state))
      .eq("session_hash", sessionHash(userId, sessionToken))
      .eq("user_id", userId)
      .eq("org_id", c.orgId)
      .eq("tenant_id", c.tenantId)
      .eq("client_id", c.clientId)
      .gt("expires_at", new Date().toISOString())
      .select("*")
      .maybeSingle(),
  );
  if (!data) fail("invalid_state");
  if (input.error || !input.code || input.code.length > 16384)
    fail("access_denied", 403);
  const row = await authorized(userId, data.connection_id, c, db);
  if (row.oauth_generation !== data.generation) fail("invalid_state");
  const verifier = decrypt(
    data.verifier_ciphertext,
    context(row, `pkce:${data.generation}`),
    c,
  );
  const result = await token(c, {
    grant_type: "authorization_code",
    code: input.code,
    redirect_uri: c.callbackUrl,
    code_verifier: verifier,
  });
  const account = await profile(result.tokens.accessToken);
  if (row.account_id && row.account_id !== account.id)
    fail("account_mismatch", 409);
  await member(userId, c, db);
  if (JSON.stringify(getMicrosoft365Config()) !== JSON.stringify(c))
    fail("unavailable", 503);
  const { data: saved } = checked(
    await db
      .from(TABLE)
      .update({
        status: "connected",
        account_id: account.id,
        account_label: account.label,
        token_ciphertext: encrypt(
          JSON.stringify(result.tokens),
          context(row, "tokens"),
          c,
        ),
        token_expires_at: result.expiresAt,
        connected_at: new Date().toISOString(),
        version: crypto.randomUUID(),
        refresh_lease_id: null,
        refresh_lease_expires_at: null,
      })
      .eq("id", row.id)
      .eq("user_id", userId)
      .eq("oauth_generation", data.generation)
      .eq("version", row.version)
      .eq("enabled", true)
      .or(
        `refresh_lease_expires_at.is.null,refresh_lease_expires_at.lt.${new Date().toISOString()}`,
      )
      .select("id")
      .maybeSingle(),
  );
  if (!saved) fail("invalid_state");
}
async function access(
  userId: string,
  row: Row,
  c: Microsoft365Config,
  db: Db,
): Promise<{ row: Row; tokens: Tokens }> {
  if (row.status !== "connected" || !row.token_ciphertext)
    fail("reconnect_required", 409);
  let tokens: Tokens;
  try {
    tokens = JSON.parse(
      decrypt(row.token_ciphertext, context(row, "tokens"), c),
    ) as Tokens;
  } catch {
    checked(
      await db
        .from(TABLE)
        .update({
          status: "reconnect_required",
          token_ciphertext: null,
          token_expires_at: null,
          version: crypto.randomUUID(),
        })
        .eq("id", row.id)
        .eq("user_id", userId)
        .eq("version", row.version),
    );
    return fail("reconnect_required", 409);
  }
  if (
    row.token_expires_at &&
    Date.parse(row.token_expires_at) > Date.now() + 120_000
  )
    return { row, tokens };
  const lease = crypto.randomUUID();
  const { data: claimed } = checked(
    await db
      .from(TABLE)
      .update({
        refresh_lease_id: lease,
        refresh_lease_expires_at: new Date(Date.now() + 60_000).toISOString(),
      })
      .eq("id", row.id)
      .eq("user_id", userId)
      .eq("version", row.version)
      .eq("enabled", true)
      .or(
        `refresh_lease_expires_at.is.null,refresh_lease_expires_at.lt.${new Date().toISOString()}`,
      )
      .select("id")
      .maybeSingle(),
  );
  if (!claimed) fail("busy", 409);
  try {
    await member(userId, c, db);
    const result = await token(
      c,
      { grant_type: "refresh_token", refresh_token: tokens.refreshToken },
      tokens.refreshToken,
    );
    const account = await profile(result.tokens.accessToken);
    if (account.id !== row.account_id) fail("account_mismatch", 409);
    await member(userId, c, db);
    if (JSON.stringify(getMicrosoft365Config()) !== JSON.stringify(c))
      fail("unavailable", 503);
    const { data: saved } = checked(
      await db
        .from(TABLE)
        .update({
          token_ciphertext: encrypt(
            JSON.stringify(result.tokens),
            context(row, "tokens"),
            c,
          ),
          token_expires_at: result.expiresAt,
          version: crypto.randomUUID(),
          refresh_lease_id: null,
          refresh_lease_expires_at: null,
        })
        .eq("id", row.id)
        .eq("user_id", userId)
        .eq("version", row.version)
        .eq("refresh_lease_id", lease)
        .eq("enabled", true)
        .select("*")
        .maybeSingle(),
    );
    if (!saved) fail("reconnect_required", 409);
    return { row: saved as Row, tokens: result.tokens };
  } catch (error) {
    const needsReconnect =
      error instanceof Microsoft365Error &&
      ["reconnect_required", "account_mismatch"].includes(error.code);
    checked(
      await db
        .from(TABLE)
        .update({
          refresh_lease_id: null,
          refresh_lease_expires_at: null,
          ...(needsReconnect
            ? {
                status: "reconnect_required",
                token_ciphertext: null,
                token_expires_at: null,
              }
            : {}),
        })
        .eq("id", row.id)
        .eq("user_id", userId)
        .eq("version", row.version)
        .eq("refresh_lease_id", lease),
    );
    throw error;
  }
}
export async function checkMicrosoft365Connection(
  userId: string,
  connectionId: string,
  db: Db,
) {
  const c = config();
  const initial = await authorized(userId, connectionId, c, db);
  const { row, tokens } = await access(userId, initial, c, db);
  await member(userId, c, db);
  const current = await authorized(userId, connectionId, c, db);
  if (current.version !== row.version) fail("busy", 409);
  const response = await request(
    "https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=id",
    { headers: { Authorization: `Bearer ${tokens.accessToken}` } },
  );
  if (!response.ok) {
    if (response.status === 401)
      checked(
        await db
          .from(TABLE)
          .update({
            status: "reconnect_required",
            token_ciphertext: null,
            token_expires_at: null,
            version: crypto.randomUUID(),
          })
          .eq("id", row.id)
          .eq("user_id", userId)
          .eq("version", row.version),
      );
    responseError(response);
  }
  const body = await json(response);
  if (!Array.isArray(body.value)) fail("provider_unavailable", 502);
  await member(userId, c, db);
  const checkedAt = new Date().toISOString();
  const { data: saved } = checked(
    await db
      .from(TABLE)
      .update({ last_checked_at: checkedAt })
      .eq("id", row.id)
      .eq("user_id", userId)
      .eq("version", row.version)
      .eq("enabled", true)
      .select("id")
      .maybeSingle(),
  );
  if (!saved) fail("not_found", 404);
  return { ok: true as const, mailAccess: true as const, checkedAt };
}
export async function disconnectMicrosoft365Connection(
  userId: string,
  connectionId: string,
  db: Db,
): Promise<void> {
  // Deliberately independent of feature configuration and org membership.
  await owned(userId, connectionId, db);
  checked(
    await db.from(TABLE).delete().eq("id", connectionId).eq("user_id", userId),
  );
}

/** Dedicated encryption envelopes are also used for protected chat derivatives. */
export function sealMicrosoft365Data(value: string, aad: string): string {
  return encrypt(value, `protected:${aad}`, config());
}
export function openMicrosoft365Data(value: string, aad: string): string {
  return decrypt(value, `protected:${aad}`, config());
}
export async function getMicrosoft365ConnectionContext(
  userId: string,
  connectionId: string,
  db: Db,
) {
  const row = await authorized(userId, connectionId, config(), db);
  if (row.status !== "connected") fail("reconnect_required", 409);
  return {
    id: row.id,
    orgId: row.org_id,
    tenantId: row.tenant_id,
    generation: row.oauth_generation,
  };
}
/** Backend-only transport: credentials never leave Graph, even on /content redirects. */
export async function microsoft365GraphRequest(
  userId: string,
  connectionId: string,
  db: Db,
  path: string,
  init: RequestInit = {},
): Promise<Response> {
  const c = config();
  if (!path.startsWith("/") || path.startsWith("//") || /[\\#\r\n]/.test(path))
    fail("invalid_source", 400);
  const url = new URL(`https://graph.microsoft.com/v1.0${path}`);
  if (
    url.origin !== "https://graph.microsoft.com" ||
    !url.pathname.startsWith("/v1.0/")
  )
    fail("invalid_source", 400);
  const method = init.method || "GET";
  if (
    method !== "GET" &&
    !(
      method === "POST" &&
      ["/search/query", "/me/translateExchangeIds"].includes(
        url.pathname.slice(5),
      )
    )
  )
    fail("invalid_source", 400);
  const initial = await authorized(userId, connectionId, c, db);
  const { row, tokens } = await access(userId, initial, c, db);
  if (
    (path.startsWith("/drives/") ||
      (path === "/search/query" && String(init.body).includes("driveItem"))) &&
    (process.env.MICROSOFT365_FILES_ENABLED !== "true" ||
      !tokens.scopes.some((scope) => /(?:^|\/)files\.read\.all$/i.test(scope)))
  )
    fail("reconnect_required", 409);
  const before = await authorized(userId, connectionId, c, db);
  if (
    before.oauth_generation !== row.oauth_generation ||
    before.status !== "connected"
  )
    fail("reconnect_required", 409);
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${tokens.accessToken}`);
  const signal = init.signal
    ? AbortSignal.any([init.signal, AbortSignal.timeout(20_000)])
    : AbortSignal.timeout(20_000);
  let response: Response;
  try {
    response = await fetch(url.href, {
      ...init,
      method,
      headers,
      redirect: "manual",
      signal,
    });
  } catch {
    fail("provider_unavailable", 502);
  }
  const after = await authorized(userId, connectionId, c, db);
  if (
    after.oauth_generation !== row.oauth_generation ||
    after.status !== "connected"
  )
    fail("reconnect_required", 409);
  if (JSON.stringify(getMicrosoft365Config()) !== JSON.stringify(c))
    fail("unavailable", 503);
  if (response.status === 401) {
    checked(
      await db
        .from(TABLE)
        .update({
          status: "reconnect_required",
          token_ciphertext: null,
          token_expires_at: null,
          version: crypto.randomUUID(),
        })
        .eq("id", row.id)
        .eq("user_id", userId)
        .eq("version", row.version),
    );
  }
  if ([404, 410].includes(response.status)) fail("not_found", 404);
  if (
    !response.ok &&
    !(
      response.status === 302 &&
      method === "GET" &&
      /\/content$/.test(url.pathname)
    )
  )
    responseError(response);
  return response;
}

/** Bounded key rotation without token exchange; owner AAD and CAS remain unchanged. */
export async function rotateMicrosoft365ConnectionEncryption(
  db: Db,
  afterId?: string,
): Promise<{
  rotated: number;
  examined: number;
  skipped: number;
  nextCursor: string | null;
}> {
  const c = config();
  const active = encryptionKeys(c).active;
  if (afterId && !UUID.test(afterId)) fail("invalid_state", 400);
  let query = db
    .from(TABLE)
    .select("*")
    .eq("org_id", c.orgId)
    .eq("tenant_id", c.tenantId)
    .eq("client_id", c.clientId)
    .order("id", { ascending: true })
    .limit(100);
  if (afterId) query = query.gt("id", afterId);
  const { data } = checked(await query);
  const rows = (data ?? []) as Row[];
  let rotated = 0,
    skipped = 0;
  for (const row of rows) {
    if (
      !row.token_ciphertext ||
      row.token_ciphertext.startsWith(`v2.${active}.`)
    )
      continue;
    try {
      await authorized(row.user_id, row.id, c, db);
      const plaintext = decrypt(
        row.token_ciphertext,
        context(row, "tokens"),
        c,
      );
      const { data: saved } = checked(
        await db
          .from(TABLE)
          .update({
            token_ciphertext: encrypt(plaintext, context(row, "tokens"), c),
            version: crypto.randomUUID(),
          })
          .eq("id", row.id)
          .eq("user_id", row.user_id)
          .eq("version", row.version)
          .eq("oauth_generation", row.oauth_generation)
          .eq("enabled", true)
          .or(
            `refresh_lease_expires_at.is.null,refresh_lease_expires_at.lt.${new Date().toISOString()}`,
          )
          .select("id")
          .maybeSingle(),
      );
      if (saved) rotated++;
      else skipped++;
    } catch (error) {
      if (
        error instanceof Microsoft365Error &&
        ["not_found", "access_denied", "reconnect_required"].includes(
          error.code,
        )
      )
        skipped++;
      else throw error;
    }
  }
  return {
    rotated,
    examined: rows.length,
    skipped,
    nextCursor: rows.length === 100 ? rows[rows.length - 1].id : null,
  };
}
