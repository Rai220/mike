import crypto from "node:crypto";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  assertMicrosoft365OrdinaryChatPrivate,
  findMicrosoft365ChatForOrdinaryChat,
  cleanupMicrosoft365Chats,
  createMicrosoft365Chat,
  deleteMicrosoft365Chat,
  getMicrosoft365ChatPolicy,
  listMicrosoft365Chats,
  loadMicrosoft365Chat,
  saveMicrosoft365Chat,
  rotateMicrosoft365ChatEncryption,
  type ProtectedChatPayload,
} from "./protected";
import {
  getMicrosoft365Config,
  getMicrosoft365ConnectionContext,
} from "./index";

vi.mock("./index", async (importOriginal) => {
  const original = await importOriginal<typeof import("./index")>();
  const key = crypto.randomBytes(32);
  return {
    ...original,
    getMicrosoft365ConnectionContext: vi.fn(),
    getMicrosoft365Config: vi.fn(),
    sealMicrosoft365Data(value: string, aad: string) {
      const iv = crypto.randomBytes(12);
      const cipher = crypto.createCipheriv("aes-256-gcm", key, iv);
      cipher.setAAD(Buffer.from(aad));
      return [iv, cipher.update(value), cipher.final(), cipher.getAuthTag()]
        .map((part) => part.toString("base64url"))
        .join(".");
    },
    openMicrosoft365Data(value: string, aad: string) {
      const [iv, body, end, tag] = value
        .split(".")
        .map((part) => Buffer.from(part, "base64url"));
      const cipher = crypto.createDecipheriv("aes-256-gcm", key, iv);
      cipher.setAAD(Buffer.from(aad));
      cipher.setAuthTag(tag);
      return Buffer.concat([
        cipher.update(body),
        cipher.update(end),
        cipher.final(),
      ]).toString();
    },
  };
});
const USER = "11111111-1111-4111-8111-111111111111";
const OTHER = "22222222-2222-4222-8222-222222222222";
const CONNECTION = "33333333-3333-4333-8333-333333333333";
const CONTEXT = {
  id: CONNECTION,
  orgId: "organization",
  tenantId: "tenant",
  generation: "generation",
};
const payload: ProtectedChatPayload = {
  messages: [
    {
      role: "user",
      content: "confidential corporate question",
      sourceRefs: [],
    },
  ],
  sources: [],
};
type RecordRow = Record<string, any>;
function memoryDb() {
  const rows: RecordRow[] = [];
  const shells: RecordRow[] = [];
  const grants: RecordRow[] = [];
  const hooks = { beforeUpdate: () => {}, error: false };
  const db = {
    rpc: vi.fn().mockResolvedValue({ data: 2, error: null }),
    from(table: string) {
      const tableRows = table === "chats" ? shells : table === "chat_access_grants" ? grants : rows;
      let operation = "select";
      let values: RecordRow = {};
      let single = false;
      let maximum = Infinity;
      const filters: Array<(row: RecordRow) => boolean> = [];
      const query = {
        select() {
          return query;
        },
        insert(value: RecordRow) {
          operation = "insert";
          values = value;
          return query;
        },
        update(value: RecordRow) {
          operation = "update";
          values = value;
          return query;
        },
        delete() {
          operation = "delete";
          return query;
        },
        eq(key: string, value: unknown) {
          filters.push((row) => row[key] === value);
          return query;
        },
        is(key: string, value: unknown) {
          filters.push((row) => (row[key] ?? null) === value);
          return query;
        },
        gt(key: string, value: string) {
          filters.push((row) => row[key] > value);
          return query;
        },
        order() {
          return query;
        },
        limit(value: number) {
          maximum = value;
          return query;
        },
        maybeSingle() {
          single = true;
          return query;
        },
        then(resolve: (value: unknown) => unknown) {
          if (hooks.error)
            return Promise.resolve(
              resolve({ data: null, error: { message: "private DB error" } }),
            );
          if (operation === "update") hooks.beforeUpdate();
          let found = tableRows
            .filter((row) => filters.every((filter) => filter(row)))
            .slice(0, maximum);
          if (operation === "insert") {
            tableRows.push(structuredClone(values));
            found = [values];
          }
          if (operation === "update")
            found.forEach((row) => Object.assign(row, values));
          if (operation === "delete")
            found.forEach((row) => tableRows.splice(tableRows.indexOf(row), 1));
          return Promise.resolve(
            resolve({
              data: structuredClone(single ? (found[0] ?? null) : found),
              error: null,
            }),
          );
        },
      };
      return query;
    },
  };
  return {
    db: db as unknown as Parameters<typeof createMicrosoft365Chat>[2],
    rows,
    shells,
    grants,
    hooks,
    rpc: db.rpc,
  };
}
beforeEach(() => {
  vi.stubEnv("MICROSOFT365_CHAT_ENABLED", "true");
  vi.stubEnv("MICROSOFT365_CHAT_MODEL", "gpt-5.5");
  vi.stubEnv("MICROSOFT365_RETENTION_DAYS", "30");
  vi.mocked(getMicrosoft365ConnectionContext).mockResolvedValue(CONTEXT);
  vi.mocked(getMicrosoft365Config).mockReturnValue({
    clientId: "client",
    tenantId: CONTEXT.tenantId,
    clientSecret: "secret",
    orgId: CONTEXT.orgId,
    encryptionSecret: "secret".repeat(8),
    callbackUrl: "https://api.test/callback",
    frontendOrigin: "https://app.test",
  });
});
afterEach(() => {
  vi.unstubAllEnvs();
  vi.clearAllMocks();
});

describe("protected Microsoft 365 policy and storage", () => {
  it("encrypts composer selections and dependency metadata and restores them on reload", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    const file = { filename: "PRIVATE ATTACHMENT.txt", document_id: OTHER, version_id: CONNECTION, version_number: 1 };
    const enriched: ProtectedChatPayload = {
      ...payload,
      messages: [{ ...payload.messages[0], files: [file], workflow: { id: OTHER, title: "PRIVATE WORKFLOW" }, useMicrosoft365: true, useEdgar: true, model: "gpt-5.4", reasoning: "low" }],
      preferences: { model: "gpt-5.4", reasoning: "low" },
      dependencies: { documents: [{ id: OTHER, versionId: CONNECTION, version: "hash" }], workflows: [{ id: OTHER, version: "hash" }] },
    };
    await saveMicrosoft365Chat(USER, chat, enriched, db);
    expect(JSON.stringify(rows)).not.toContain("PRIVATE ATTACHMENT");
    expect(JSON.stringify(rows)).not.toContain("PRIVATE WORKFLOW");
    expect((await loadMicrosoft365Chat(USER, chat.id, db)).payload).toEqual(enriched);
  });
  it.each([
    { preferences: { model: "gpt-5.4", apiKeys: { openai: "key" } } },
    { dependencies: { documents: [{ id: "forged", versionId: CONNECTION, version: "v1" }], workflows: [] } },
    { messages: [{ role: "user", content: "question", sourceRefs: [], files: [{ filename: "bad", document_id: "forged" }] }] },
  ])("rejects malformed encrypted capability metadata %j", async (invalid) => {
    const { db } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    await expect(saveMicrosoft365Chat(USER, chat, { ...payload, ...invalid } as ProtectedChatPayload, db)).rejects.toMatchObject({ code: "invalid_state" });
  });
  it("rotates encrypted history with CAS without changing retention or plaintext", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    const saved = await saveMicrosoft365Chat(USER, chat, payload, db);
    const before = structuredClone(rows[0]);
    const result = await rotateMicrosoft365ChatEncryption(db);
    expect(result).toEqual({
      rotated: 1,
      examined: 1,
      skipped: 0,
      nextCursor: null,
    });
    expect(rows[0].payload_ciphertext).not.toBe(before.payload_ciphertext);
    expect(rows[0].version).not.toBe(before.version);
    expect(rows[0].expires_at).toBe(before.expires_at);
    expect(rows[0].updated_at).toBe(before.updated_at);
    expect((await loadMicrosoft365Chat(USER, saved.id, db)).payload).toEqual(
      payload,
    );
  });
  it("rotation skips active envelopes and old generations and respects configured tenant", async () => {
    const { db, rows } = memoryDb();
    await createMicrosoft365Chat(USER, CONNECTION, db);
    await createMicrosoft365Chat(USER, CONNECTION, db);
    await createMicrosoft365Chat(USER, CONNECTION, db);
    rows[0].payload_ciphertext = "v2.legacy.already-active";
    rows[1].generation = "old";
    rows[2].tenant_id = "other";
    expect(await rotateMicrosoft365ChatEncryption(db)).toEqual({
      rotated: 0,
      examined: 2,
      skipped: 1,
      nextCursor: null,
    });
  });
  it("rotation cannot overwrite a concurrent turn or resurrect deleted chat", async () => {
    const { db, rows, hooks } = memoryDb();
    await createMicrosoft365Chat(USER, CONNECTION, db);
    hooks.beforeUpdate = () => {
      rows[0].version = "raced";
    };
    expect((await rotateMicrosoft365ChatEncryption(db)).skipped).toBe(1);
    hooks.beforeUpdate = () => {
      rows.splice(0);
    };
    expect((await rotateMicrosoft365ChatEncryption(db)).skipped).toBe(1);
    expect(rows).toHaveLength(0);
  });
  it("bounds rotation to 100 rows and returns a continuation cursor", async () => {
    const { db, rows } = memoryDb();
    await createMicrosoft365Chat(USER, CONNECTION, db);
    rows[0].payload_ciphertext = "v2.legacy.active";
    while (rows.length < 101)
      rows.push({ ...rows[0], id: crypto.randomUUID() });
    expect(await rotateMicrosoft365ChatEncryption(db)).toMatchObject({
      examined: 100,
      nextCursor: rows[99].id,
    });
  });
  it.each([
    "gpt-unknown",
    "gpt-5.4-lite",
    "openrouter/openai/gpt-5.5",
    "vercel/openai/gpt-5.5",
    "opencode-go/glm-5",
    "ollama/",
    "ollama/local model",
  ])("rejects unknown, renamed or brokered model %s", (model) => {
    vi.stubEnv("MICROSOFT365_CHAT_MODEL", model);
    expect(getMicrosoft365ChatPolicy()).toBeNull();
  });
  it.each([
    "gpt-5.5",
    "claude-sonnet-4-6",
    "gemini-3-flash-preview",
    "ollama/qwen3:8b",
  ])("accepts explicitly configured direct model %s", (model) => {
    vi.stubEnv("MICROSOFT365_CHAT_MODEL", model);
    expect(getMicrosoft365ChatPolicy()?.model).toBe(model);
  });
  it.each(["", "0", "91", "-1", "1.5", "30days", " 30", "030"])(
    "fails closed for retention %s",
    (days) => {
      vi.stubEnv("MICROSOFT365_RETENTION_DAYS", days);
      expect(getMicrosoft365ChatPolicy()).toBeNull();
    },
  );
  it("requires explicit enablement and model approval", () => {
    expect(getMicrosoft365ChatPolicy()).toEqual({
      model: "gpt-5.5",
      retentionDays: 30,
    });
    vi.stubEnv("MICROSOFT365_CHAT_MODEL", "");
    expect(getMicrosoft365ChatPolicy()).toBeNull();
    vi.stubEnv("MICROSOFT365_CHAT_MODEL", "gpt-5.5");
    vi.stubEnv("MICROSOFT365_CHAT_ENABLED", "false");
    expect(getMicrosoft365ChatPolicy()).toBeNull();
  });
  it("persists corporate text only in encrypted payload and lists metadata only", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    const saved = await saveMicrosoft365Chat(USER, chat, payload, db);
    expect(JSON.stringify(rows)).not.toContain("confidential");
    expect((await loadMicrosoft365Chat(USER, saved.id, db)).payload).toEqual(
      payload,
    );
    const listed = await listMicrosoft365Chats(USER, CONNECTION, db);
    expect(listed).toHaveLength(1);
    expect(listed[0]).not.toHaveProperty("payload");
    expect(listed[0]).not.toHaveProperty("title");
  });
  it("isolates another user's load, list and deletion", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    await expect(
      loadMicrosoft365Chat(OTHER, chat.id, db),
    ).rejects.toMatchObject({ code: "not_found" });
    expect(await listMicrosoft365Chats(OTHER, CONNECTION, db)).toEqual([]);
    await deleteMicrosoft365Chat(OTHER, chat.id, db);
    expect(rows).toHaveLength(1);
  });
  it("rejects generation changes and revoked membership before decrypting", async () => {
    const { db } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    vi.mocked(getMicrosoft365ConnectionContext).mockResolvedValue({
      ...CONTEXT,
      generation: "changed",
    });
    await expect(loadMicrosoft365Chat(USER, chat.id, db)).rejects.toMatchObject(
      { code: "access_denied" },
    );
    vi.mocked(getMicrosoft365ConnectionContext).mockRejectedValue(
      new Error("revoked"),
    );
    await expect(loadMicrosoft365Chat(USER, chat.id, db)).rejects.toThrow(
      "revoked",
    );
  });
  it("detects ciphertext copied across chat identities", async () => {
    const { db, rows } = memoryDb();
    await createMicrosoft365Chat(USER, CONNECTION, db);
    const second = await createMicrosoft365Chat(USER, CONNECTION, db);
    rows[1].payload_ciphertext = rows[0].payload_ciphertext;
    await expect(
      loadMicrosoft365Chat(USER, second.id, db),
    ).rejects.toMatchObject({ code: "storage_unavailable" });
  });
  it("purges expired data on load and never extends retention on save", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    const saved = await saveMicrosoft365Chat(USER, chat, payload, db);
    expect(saved.expiresAt).toBe(chat.expiresAt);
    rows[0].expires_at = new Date(Date.now() - 1).toISOString();
    await expect(loadMicrosoft365Chat(USER, chat.id, db)).rejects.toMatchObject(
      { code: "not_found" },
    );
    expect(rows).toHaveLength(0);
  });
  it("applies a reduced retention policy to existing history", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    rows[0].created_at = new Date(Date.now() - 2 * 86_400_000).toISOString();
    vi.stubEnv("MICROSOFT365_RETENTION_DAYS", "1");
    expect(await listMicrosoft365Chats(USER, CONNECTION, db)).toEqual([]);
    await expect(loadMicrosoft365Chat(USER, chat.id, db)).rejects.toMatchObject(
      { code: "not_found" },
    );
  });
  it("rejects stale writers before overwriting another answer", async () => {
    const { db } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    await saveMicrosoft365Chat(USER, chat, payload, db);
    await expect(
      saveMicrosoft365Chat(USER, chat, payload, db),
    ).rejects.toMatchObject({ code: "busy" });
  });
  it("CAS rejects a write racing another writer", async () => {
    const { db, rows, hooks } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    hooks.beforeUpdate = () => {
      rows[0].version = "other-version";
    };
    await expect(
      saveMicrosoft365Chat(USER, chat, payload, db),
    ).rejects.toMatchObject({ code: "busy" });
  });
  it("does not recreate a deleted chat during a racing save", async () => {
    const { db, rows, hooks } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    hooks.beforeUpdate = () => {
      rows.splice(0);
    };
    await expect(
      saveMicrosoft365Chat(USER, chat, payload, db),
    ).rejects.toMatchObject({ code: "busy" });
    expect(rows).toHaveLength(0);
  });
  it("allows owner deletion while feature policy is disabled", async () => {
    const { db, rows } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    vi.stubEnv("MICROSOFT365_CHAT_ENABLED", "false");
    await deleteMicrosoft365Chat(USER, chat.id, db);
    expect(rows).toHaveLength(0);
  });
  it("rejects oversized history and references with no actual source", async () => {
    const { db } = memoryDb();
    const chat = await createMicrosoft365Chat(USER, CONNECTION, db);
    await expect(
      saveMicrosoft365Chat(
        USER,
        chat,
        {
          sources: [],
          messages: [
            { role: "user", content: "a".repeat(1_048_576), sourceRefs: [] },
          ],
        },
        db,
      ),
    ).rejects.toMatchObject({ status: 413 });
    await expect(
      saveMicrosoft365Chat(
        USER,
        chat,
        {
          sources: [],
          messages: [
            {
              role: "assistant",
              content: "fabricated citation",
              sourceRefs: ["missing"],
            },
          ],
        },
        db,
      ),
    ).rejects.toMatchObject({ code: "invalid_state" });
  });
  it("returns only sanitized storage failures and runs metadata-only cleanup", async () => {
    const { db, hooks, rpc } = memoryDb();
    hooks.error = true;
    await expect(
      createMicrosoft365Chat(USER, CONNECTION, db),
    ).rejects.toMatchObject({ message: "storage_unavailable" });
    expect(await cleanupMicrosoft365Chats(db)).toBe(2);
    expect(rpc).toHaveBeenCalledWith("cleanup_microsoft365_chats", {
      retention_days: 30,
    });
  });
});


describe("ordinary Microsoft 365 chat binding", () => {
  const ordinaryId = "44444444-4444-4444-8444-444444444444";
  function fixture() {
    const value = memoryDb();
    value.shells.push({ id: ordinaryId, user_id: USER, project_id: null, org_id: null, microsoft365_protected: false });
    return value;
  }
  it("binds the encrypted payload to the ordinary identity and reuses it", async () => {
    const { db, rows, shells } = fixture();
    const first = await createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId);
    shells[0].microsoft365_protected = true;
    expect(first.ordinaryChatId).toBe(ordinaryId);
    expect(await listMicrosoft365Chats(USER, CONNECTION, db)).toEqual([]);
    expect(rows[0].ordinary_chat_id).toBe(ordinaryId);
    expect((await createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId)).id).toBe(first.id);
    expect((await findMicrosoft365ChatForOrdinaryChat(USER, ordinaryId, db))?.id).toBe(first.id);
    rows[0].ordinary_chat_id = OTHER;
    await expect(loadMicrosoft365Chat(USER, first.id, db)).rejects.toThrow();
  });
  it("rejects a foreign, shared or project chat", async () => {
    const { db, shells, grants } = fixture();
    await expect(assertMicrosoft365OrdinaryChatPrivate(OTHER, ordinaryId, db)).rejects.toMatchObject({ code: "not_found" });
    grants.push({ id: "grant", chat_id: ordinaryId });
    await expect(createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId)).rejects.toMatchObject({ code: "access_denied" });
    grants.length = 0;
    shells[0].project_id = OTHER;
    await expect(createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId)).rejects.toMatchObject({ code: "access_denied" });
  });
  it("does not recreate deleted or expired protected history", async () => {
    const { db, shells, rows } = fixture();
    const created = await createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId);
    shells[0].microsoft365_protected = true;
    rows[0].expires_at = new Date(Date.now() - 1000).toISOString();
    await expect(findMicrosoft365ChatForOrdinaryChat(USER, ordinaryId, db)).rejects.toMatchObject({ code: "not_found" });
    expect(rows).toHaveLength(0);
    await expect(createMicrosoft365Chat(USER, CONNECTION, db, ordinaryId)).rejects.toMatchObject({ code: "not_found" });
    expect(created.id).toBeTruthy();
  });
});
