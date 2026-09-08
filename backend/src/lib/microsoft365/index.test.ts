import crypto from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { getOrgRole } from "../access";
import {
  rotateMicrosoft365ConnectionEncryption,
  microsoft365GraphRequest,
  getMicrosoft365ConnectionContext,
  sealMicrosoft365Data,
  openMicrosoft365Data,
  checkMicrosoft365Connection,
  completeMicrosoft365Connection,
  disconnectMicrosoft365Connection,
  getMicrosoft365Config,
  getMicrosoft365Status,
  MICROSOFT365_SCOPES,
  startMicrosoft365Connection,
} from "./index";

vi.mock("../access", () => ({ getOrgRole: vi.fn() }));
const USER = "11111111-1111-4111-8111-111111111111";
const OTHER = "22222222-2222-4222-8222-222222222222";
const ORG = "33333333-3333-4333-8333-333333333333";
const TENANT = "44444444-4444-4444-8444-444444444444";
const CLIENT = "55555555-5555-4555-8555-555555555555";
const ACCOUNT = "66666666-6666-4666-8666-666666666666";
const CONNECTIONS = "user_microsoft365_connections";
const STATES = "microsoft365_oauth_states";
type RecordRow = Record<string, any>;
type ServiceDb = Parameters<typeof startMicrosoft365Connection>[2];

/** Exercises filtering and conditional mutations, rather than prescribing call order. */
function memoryDb() {
  const tables: Record<string, RecordRow[]> = {
    [CONNECTIONS]: [],
    [STATES]: [],
  };
  const db = {
    from(table: string) {
      let operation = "select";
      let values: RecordRow = {};
      let one = false;
      let max = Infinity;
      const filters: ((r: RecordRow) => boolean)[] = [];
      const query = {
        select(_columns?: string) {
          return query;
        },
        insert(v: RecordRow) {
          operation = "insert";
          values = v;
          return query;
        },
        update(v: RecordRow) {
          operation = "update";
          values = v;
          return query;
        },
        delete() {
          operation = "delete";
          return query;
        },
        eq(key: string, val: unknown) {
          filters.push((r) => r[key] === val);
          return query;
        },
        lt(key: string, val: string) {
          filters.push((r) => r[key] < val);
          return query;
        },
        gt(key: string, val: string) {
          filters.push((r) => r[key] > val);
          return query;
        },
        or(value: string) {
          const stamp = value.split(".lt.")[1];
          filters.push(
            (r) =>
              r.refresh_lease_expires_at == null ||
              r.refresh_lease_expires_at < stamp,
          );
          return query;
        },
        order() {
          return query;
        },
        limit(count: number) {
          max = count;
          return query;
        },
        single() {
          one = true;
          return query;
        },
        maybeSingle() {
          one = true;
          return query;
        },
        then(resolve: (value: unknown) => unknown) {
          let rows = tables[table]
            .filter((r) => filters.every((f) => f(r)))
            .slice(0, max);
          if (operation === "insert") {
            const row = {
              id: crypto.randomUUID(),
              created_at: new Date().toISOString(),
              enabled: true,
              status: "pending",
              account_id: null,
              account_label: null,
              token_ciphertext: null,
              token_expires_at: null,
              connected_at: null,
              last_checked_at: null,
              ...values,
            };
            tables[table].push(row);
            rows = [row];
          } else if (operation === "update")
            rows.forEach((r) => Object.assign(r, values));
          else if (operation === "delete") {
            tables[table] = tables[table].filter((r) => !rows.includes(r));
            if (table === CONNECTIONS)
              tables[STATES] = tables[STATES].filter(
                (s) => !rows.some((r) => r.id === s.connection_id),
              );
          }
          return Promise.resolve(
            resolve({
              data: structuredClone(one ? (rows[0] ?? null) : rows),
              error: null,
            }),
          );
        },
      };
      return query;
    },
  };
  return { db: db as unknown as ServiceDb, tables };
}
const response = (body: unknown, status = 200, headers?: HeadersInit) =>
  new Response(JSON.stringify(body), { status, headers });
const tokenResponse = (extra = {}) =>
  response({
    access_token: "private-access-token",
    refresh_token: "private-refresh-token",
    token_type: "Bearer",
    scope: "User.Read Mail.Read",
    expires_in: 3600,
    ...extra,
  });
const profileResponse = (id = ACCOUNT) =>
  response({ id, mail: "employee@example.test" });
const jwt = (sessionId: string, suffix = "signature") =>
  `header.${Buffer.from(JSON.stringify({ session_id: sessionId })).toString("base64url")}.${suffix}`;
let fetchMock: ReturnType<typeof vi.fn>;

beforeEach(() => {
  vi.stubEnv("MICROSOFT365_ENABLED", "true");
  vi.stubEnv("MICROSOFT365_CLIENT_ID", CLIENT);
  vi.stubEnv("MICROSOFT365_TENANT_ID", TENANT);
  vi.stubEnv("MICROSOFT365_ORG_ID", ORG);
  vi.stubEnv("MICROSOFT365_CLIENT_SECRET", "private-client-secret");
  vi.stubEnv(
    "MICROSOFT365_ENCRYPTION_SECRET",
    "dedicated-private-key".repeat(3),
  );
  vi.stubEnv("API_PUBLIC_URL", "https://api.example.test");
  vi.stubEnv("FRONTEND_URL", "https://app.example.test");
  vi.stubEnv("NODE_ENV", "test");
  vi.mocked(getOrgRole).mockResolvedValue("member");
  fetchMock = vi.fn();
  vi.stubGlobal("fetch", fetchMock);
});
afterEach(() => {
  vi.unstubAllEnvs();
  vi.unstubAllGlobals();
  vi.clearAllMocks();
});

async function start(db: ServiceDb, session = "verified-session") {
  const result = await startMicrosoft365Connection(USER, session, db);
  return {
    ...result,
    state: new URL(result.authorizationUrl).searchParams.get("state")!,
  };
}
async function connect(db: ServiceDb) {
  const result = await start(db);
  fetchMock
    .mockResolvedValueOnce(tokenResponse())
    .mockResolvedValueOnce(profileResponse());
  await completeMicrosoft365Connection(
    USER,
    "verified-session",
    { state: result.state, code: "code" },
    db,
  );
}

describe("Microsoft 365 connection boundary", () => {
  it("requires dedicated complete configuration and production HTTPS", () => {
    expect(getMicrosoft365Config()?.callbackUrl).toBe(
      "https://api.example.test/integrations/microsoft365/callback",
    );
    vi.stubEnv("MICROSOFT365_ENCRYPTION_SECRET", "short");
    expect(getMicrosoft365Config()).toBeNull();
    vi.stubEnv("MICROSOFT365_ENCRYPTION_SECRET", "x".repeat(32));
    vi.stubEnv("NODE_ENV", "production");
    vi.stubEnv("API_PUBLIC_URL", "http://localhost:8400");
    expect(getMicrosoft365Config()).toBeNull();
    vi.stubEnv("API_PUBLIC_URL", "https://api.example.test/api/");
    expect(getMicrosoft365Config()?.callbackUrl).toBe(
      "https://api.example.test/api/integrations/microsoft365/callback",
    );
    vi.stubEnv("API_PUBLIC_URL", "https://api.example.test/api?query=true");
    expect(getMicrosoft365Config()).toBeNull();
    vi.stubEnv("API_PUBLIC_URL", "https://api.example.test");
    vi.stubEnv("MICROSOFT365_TENANT_ID", "common");
    expect(getMicrosoft365Config()).toBeNull();
  });
  it("requests only read-only delegated scopes with PKCE and stores hashed/encrypted state", async () => {
    const { db, tables } = memoryDb();
    const result = await start(db);
    const url = new URL(result.authorizationUrl);
    expect(url.pathname).toBe(`/${TENANT}/oauth2/v2.0/authorize`);
    expect(url.searchParams.get("scope")).toBe(MICROSOFT365_SCOPES.join(" "));
    expect(url.searchParams.get("code_challenge_method")).toBe("S256");
    expect(url.searchParams.has("client_secret")).toBe(false);
    expect(result.callbackOrigin).toBe("https://api.example.test");
    expect(JSON.stringify(tables)).not.toContain(result.state);
    expect(JSON.stringify(tables)).not.toContain("verified-session");
    expect(tables[STATES][0].verifier_ciphertext.split(".")).toHaveLength(5);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("rejects non-members before initiating OAuth", async () => {
    const { db, tables } = memoryDb();
    vi.mocked(getOrgRole).mockResolvedValue(null);
    await expect(start(db)).rejects.toMatchObject({ code: "access_denied" });
    expect(tables[CONNECTIONS]).toHaveLength(0);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("rejects wrong session and owner without consuming legitimate state", async () => {
    const { db, tables } = memoryDb();
    const { state } = await start(db);
    await expect(
      completeMicrosoft365Connection(
        USER,
        "other-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    await expect(
      completeMicrosoft365Connection(
        OTHER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    expect(tables[STATES]).toHaveLength(1);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("binds OAuth to stable verified session_id across JWT refresh", async () => {
    const { db } = memoryDb();
    const { state } = await start(db, jwt(ACCOUNT));
    fetchMock
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(profileResponse());
    await completeMicrosoft365Connection(
      USER,
      jwt(ACCOUNT, "refreshed-signature"),
      { state, code: "code" },
      db,
    );
    expect((await getMicrosoft365Status(USER, db)).connection?.status).toBe(
      "connected",
    );
  });
  it("rejects expired state and changed tenant before provider calls", async () => {
    const { db, tables } = memoryDb();
    const { state } = await start(db);
    tables[STATES][0].expires_at = "2000-01-01T00:00:00.000Z";
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    tables[STATES][0].expires_at = "2999-01-01T00:00:00.000Z";
    vi.stubEnv("MICROSOFT365_TENANT_ID", OTHER);
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("consumes denied OAuth state and blocks replay", async () => {
    const { db } = memoryDb();
    const { state } = await start(db);
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, error: "provider-private-detail" },
        db,
      ),
    ).rejects.toMatchObject({
      code: "access_denied",
      message: "access_denied",
    });
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("stores encrypted tokens and exposes only safe connection status", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    expect(JSON.stringify(tables)).not.toContain("private-access-token");
    expect(JSON.stringify(tables)).not.toContain("private-refresh-token");
    const status = await getMicrosoft365Status(USER, db);
    expect(status).toMatchObject({
      available: true,
      connection: {
        status: "connected",
        accountLabel: "employee@example.test",
      },
    });
    expect(Object.keys(status.connection!)).toEqual([
      "id",
      "status",
      "accountLabel",
      "connectedAt",
      "lastCheckedAt",
    ]);
  });
  it("refuses tokens without delegated mail permission", async () => {
    const { db, tables } = memoryDb();
    const { state } = await start(db);
    fetchMock.mockResolvedValueOnce(tokenResponse({ scope: "User.Read" }));
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    expect(tables[CONNECTIONS][0].token_ciphertext).toBeNull();
  });
  it("does not silently replace an existing Microsoft account", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const ciphertext = tables[CONNECTIONS][0].token_ciphertext;
    const { state } = await start(db);
    fetchMock
      .mockResolvedValueOnce(tokenResponse())
      .mockResolvedValueOnce(profileResponse(OTHER));
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "account_mismatch" });
    expect(tables[CONNECTIONS][0].account_id).toBe(ACCOUNT);
    expect(tables[CONNECTIONS][0].token_ciphertext).toBe(ciphertext);
  });
  it("cannot resurrect a connection disconnected during token exchange", async () => {
    const { db, tables } = memoryDb();
    const { state } = await start(db);
    fetchMock
      .mockImplementationOnce(async () => {
        await disconnectMicrosoft365Connection(
          USER,
          tables[CONNECTIONS][0].id,
          db,
        );
        return tokenResponse();
      })
      .mockResolvedValueOnce(profileResponse());
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
    expect(tables[CONNECTIONS]).toHaveLength(0);
  });
  it("blocks owner mismatch and org removal before mailbox access; owner can still disconnect", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const id = tables[CONNECTIONS][0].id;
    fetchMock.mockClear();
    await expect(
      checkMicrosoft365Connection(OTHER, id, db),
    ).rejects.toMatchObject({ code: "not_found" });
    await expect(
      disconnectMicrosoft365Connection(OTHER, id, db),
    ).rejects.toMatchObject({ code: "not_found" });
    vi.mocked(getOrgRole).mockResolvedValue(null);
    await expect(
      checkMicrosoft365Connection(USER, id, db),
    ).rejects.toMatchObject({ code: "access_denied" });
    expect(
      (await getMicrosoft365Status(USER, db)).connection?.accountLabel,
    ).toBeNull();
    vi.stubEnv("MICROSOFT365_ENABLED", "false");
    await disconnectMicrosoft365Connection(USER, id, db);
    expect(tables[CONNECTIONS]).toHaveLength(0);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("checks only one message ID and returns no message data", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    fetchMock.mockClear();
    fetchMock.mockResolvedValueOnce(
      response({ value: [{ id: "private-message-id" }] }),
    );
    const result = await checkMicrosoft365Connection(
      USER,
      tables[CONNECTIONS][0].id,
      db,
    );
    expect(fetchMock.mock.calls[0][0]).toBe(
      "https://graph.microsoft.com/v1.0/me/messages?$top=1&$select=id",
    );
    expect(result).toEqual({
      ok: true,
      mailAccess: true,
      checkedAt: expect.any(String),
    });
    expect(JSON.stringify(tables)).not.toContain("private-message-id");
  });
  it("converts Graph 401 to reconnect and clears stored credentials", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    fetchMock.mockResolvedValueOnce(
      response({ error: { message: "private-provider-message" } }, 401),
    );
    await expect(
      checkMicrosoft365Connection(USER, tables[CONNECTIONS][0].id, db),
    ).rejects.toMatchObject({
      code: "reconnect_required",
      detail: "Please reconnect your Microsoft 365 account.",
    });
    expect(tables[CONNECTIONS][0]).toMatchObject({
      status: "reconnect_required",
      token_ciphertext: null,
    });
  });
  it("bounds Retry-After and suppresses provider messages", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    fetchMock.mockResolvedValueOnce(
      response({ error: { message: "private-provider-message" } }, 429, {
        "retry-after": "10000",
      }),
    );
    await expect(
      checkMicrosoft365Connection(USER, tables[CONNECTIONS][0].id, db),
    ).rejects.toMatchObject({
      code: "rate_limited",
      retryAfterSeconds: 300,
      message: "rate_limited",
    });
  });
  it("serializes concurrent refreshes with a database lease", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const row = tables[CONNECTIONS][0];
    row.token_expires_at = "2000-01-01T00:00:00Z";
    let release!: (value: Response) => void;
    let requested!: () => void;
    const waiting = new Promise<void>((resolve) => {
      requested = resolve;
    });
    fetchMock
      .mockImplementationOnce(() => {
        requested();
        return new Promise<Response>((resolve) => {
          release = resolve;
        });
      })
      .mockResolvedValueOnce(profileResponse())
      .mockResolvedValueOnce(response({ value: [] }));
    const first = checkMicrosoft365Connection(USER, row.id, db);
    await waiting;
    await expect(
      checkMicrosoft365Connection(USER, row.id, db),
    ).rejects.toMatchObject({ code: "busy" });
    await expect(start(db)).rejects.toMatchObject({ code: "busy" });
    release(tokenResponse({ refresh_token: "rotated-refresh-token" }));
    await expect(first).resolves.toMatchObject({ ok: true });
    expect(row.refresh_lease_id).toBeNull();
    expect(JSON.stringify(tables)).not.toContain("rotated-refresh-token");
  });
  it("refresh invalid_grant requires reconnect without leaking its body", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const row = tables[CONNECTIONS][0];
    row.token_expires_at = "2000-01-01T00:00:00Z";
    fetchMock.mockResolvedValueOnce(
      response(
        { error: "invalid_grant", error_description: "private-detail" },
        400,
      ),
    );
    await expect(
      checkMicrosoft365Connection(USER, row.id, db),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    expect(row).toMatchObject({
      token_ciphertext: null,
      status: "reconnect_required",
      refresh_lease_id: null,
    });
  });
  it("disconnection during refresh prevents stale writes and Graph mailbox calls", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const row = tables[CONNECTIONS][0];
    row.token_expires_at = "2000-01-01T00:00:00Z";
    fetchMock.mockClear();
    fetchMock
      .mockImplementationOnce(async () => {
        await disconnectMicrosoft365Connection(USER, row.id, db);
        return tokenResponse();
      })
      .mockResolvedValueOnce(profileResponse());
    await expect(
      checkMicrosoft365Connection(USER, row.id, db),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    expect(tables[CONNECTIONS]).toHaveLength(0);
    expect(
      fetchMock.mock.calls.some((call) =>
        String(call[0]).includes("/me/messages"),
      ),
    ).toBe(false);
  });
  it("prefers the configured connection over an older org connection", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const current = tables[CONNECTIONS][0];
    tables[CONNECTIONS].unshift({
      ...current,
      id: OTHER,
      org_id: OTHER,
      account_label: "old-account@example.test",
    });
    expect((await getMicrosoft365Status(USER, db)).connection?.id).toBe(
      current.id,
    );
    vi.stubEnv("MICROSOFT365_ORG_ID", ACCOUNT);
    expect((await getMicrosoft365Status(USER, db)).connection).toMatchObject({
      id: OTHER,
      accountLabel: null,
    });
  });
  it("tolerates an unmigrated database only while disabled", async () => {
    const query: Record<string, any> = {};
    for (const method of ["select", "eq", "order", "limit", "maybeSingle"])
      query[method] = () => query;
    query.then = (resolve: (value: unknown) => void) =>
      resolve({ data: null, error: { code: "42P01" } });
    const db = { from: () => query } as unknown as ServiceDb;
    vi.stubEnv("MICROSOFT365_ENABLED", "false");
    await expect(getMicrosoft365Status(USER, db)).resolves.toEqual({
      available: false,
      connection: null,
    });
    vi.stubEnv("MICROSOFT365_ENABLED", "true");
    await expect(getMicrosoft365Status(USER, db)).rejects.toMatchObject({
      code: "storage_unavailable",
    });
  });
  it("rejects swapped ciphertext from another connection through authenticated encryption", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    tables[CONNECTIONS][0].id = OTHER;
    fetchMock.mockClear();
    await expect(
      checkMicrosoft365Connection(USER, OTHER, db),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    expect(tables[CONNECTIONS][0]).toMatchObject({
      status: "reconnect_required",
      token_ciphertext: null,
    });
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("fails closed when membership is revoked during OAuth exchange", async () => {
    const { db, tables } = memoryDb();
    const { state } = await start(db);
    fetchMock
      .mockImplementationOnce(async () => {
        vi.mocked(getOrgRole).mockResolvedValue(null);
        return tokenResponse();
      })
      .mockResolvedValueOnce(profileResponse());
    await expect(
      completeMicrosoft365Connection(
        USER,
        "verified-session",
        { state, code: "code" },
        db,
      ),
    ).rejects.toMatchObject({ code: "access_denied" });
    expect(tables[CONNECTIONS][0].token_ciphertext).toBeNull();
  });
});

describe("protected Graph transport and rotating encryption", () => {
  it("retains old encrypted history across key rotation and binds owner AAD", () => {
    const before = sealMicrosoft365Data("synthetic history", "owner-one");
    const old = process.env.MICROSOFT365_ENCRYPTION_SECRET!;
    vi.stubEnv("MICROSOFT365_ENCRYPTION_KEY_ID", "key-2026-09");
    vi.stubEnv("MICROSOFT365_ENCRYPTION_SECRET", "new-private-key".repeat(4));
    vi.stubEnv(
      "MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS",
      JSON.stringify({ legacy: old }),
    );
    expect(openMicrosoft365Data(before, "owner-one")).toBe("synthetic history");
    expect(() => openMicrosoft365Data(before, "owner-two")).toThrow();
    const after = sealMicrosoft365Data("new history", "owner-one");
    expect(after.startsWith("v2.key-2026-09.")).toBe(true);
    expect(openMicrosoft365Data(after, "owner-one")).toBe("new history");
    vi.stubEnv("MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS", "{}");
    expect(() => openMicrosoft365Data(before, "owner-one")).toThrow();
  });
  it("rejects corrupt keyrings without exposing secret material", () => {
    vi.stubEnv("MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS", '{"bad":"secret"}');
    expect(() => sealMicrosoft365Data("body", "owner")).toThrow("unavailable");
  });
  it("preserves old mail permission while requesting files only when enabled", async () => {
    vi.stubEnv("MICROSOFT365_FILES_ENABLED", "true");
    const { db } = memoryDb();
    const result = await start(db);
    expect(
      new URL(result.authorizationUrl).searchParams.get("scope"),
    ).toContain("Files.Read.All");
  });
  it("rejects cross-owner and arbitrary external transport paths before any network IO", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const id = tables[CONNECTIONS][0].id;
    fetchMock.mockClear();
    for (const path of [
      "https://attacker.test/",
      "//attacker.test/",
      "/../../me",
      "/me#fragment",
      "/me\\bad",
    ])
      await expect(
        microsoft365GraphRequest(USER, id, db, path),
      ).rejects.toBeDefined();
    await expect(
      microsoft365GraphRequest(OTHER, id, db, "/me"),
    ).rejects.toBeDefined();
    await expect(
      microsoft365GraphRequest(USER, id, db, "/me/sendMail", {
        method: "POST",
      }),
    ).rejects.toBeDefined();
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it("uses manual redirects and preserves body format headers without sending credentials onward", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const id = tables[CONNECTIONS][0].id;
    fetchMock.mockResolvedValueOnce(new Response("{}"));
    await microsoft365GraphRequest(USER, id, db, "/me/messages/immutable", {
      headers: {
        Prefer: 'IdType="ImmutableId", outlook.body-content-type="text"',
      },
    });
    const [url, init] = fetchMock.mock.calls.at(-1)!;
    expect(url).toBe("https://graph.microsoft.com/v1.0/me/messages/immutable");
    expect(init.redirect).toBe("manual");
    expect(init.headers.get("Prefer")).toContain(
      'outlook.body-content-type="text"',
    );
    expect(init.headers.get("Authorization")).toMatch(/^Bearer /);
  });
  it("does not deliver a response when disconnected during Graph IO", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const id = tables[CONNECTIONS][0].id;
    fetchMock.mockImplementationOnce(async () => {
      await disconnectMicrosoft365Connection(USER, id, db);
      return new Response("private");
    });
    await expect(
      microsoft365GraphRequest(USER, id, db, "/me/messages/immutable"),
    ).rejects.toBeDefined();
  });
  it("marks Graph401 reconnect and blocks files when the existing grant lacks scope", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const id = tables[CONNECTIONS][0].id;
    vi.stubEnv("MICROSOFT365_FILES_ENABLED", "true");
    await expect(
      microsoft365GraphRequest(USER, id, db, "/drives/one/items/two"),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    fetchMock.mockResolvedValueOnce(new Response("{}", { status: 401 }));
    await expect(
      microsoft365GraphRequest(USER, id, db, "/me/messages/immutable"),
    ).rejects.toMatchObject({ code: "reconnect_required" });
    await expect(
      getMicrosoft365ConnectionContext(USER, id, db),
    ).rejects.toMatchObject({ code: "reconnect_required" });
  });
});

describe("credential rotation", () => {
  it("rewraps existing credentials without calling Microsoft and keeps them usable", async () => {
    const { db, tables } = memoryDb();
    await connect(db);
    const oldCipher = tables[CONNECTIONS][0].token_ciphertext;
    const oldSecret = process.env.MICROSOFT365_ENCRYPTION_SECRET!;
    vi.stubEnv("MICROSOFT365_ENCRYPTION_KEY_ID", "new-key");
    vi.stubEnv("MICROSOFT365_ENCRYPTION_SECRET", "new-secret".repeat(5));
    vi.stubEnv(
      "MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS",
      JSON.stringify({ legacy: oldSecret }),
    );
    fetchMock.mockClear();
    expect(await rotateMicrosoft365ConnectionEncryption(db)).toMatchObject({
      rotated: 1,
      skipped: 0,
      nextCursor: null,
    });
    expect(fetchMock).not.toHaveBeenCalled();
    expect(tables[CONNECTIONS][0].token_ciphertext).not.toBe(oldCipher);
    expect(
      tables[CONNECTIONS][0].token_ciphertext.startsWith("v2.new-key."),
    ).toBe(true);
    vi.stubEnv("MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS", "{}");
    fetchMock.mockResolvedValueOnce(new Response("{}"));
    await expect(
      microsoft365GraphRequest(
        USER,
        tables[CONNECTIONS][0].id,
        db,
        "/me/messages/immutable",
      ),
    ).resolves.toBeInstanceOf(Response);
    expect(await rotateMicrosoft365ConnectionEncryption(db)).toMatchObject({
      rotated: 0,
      skipped: 0,
    });
  });
});
