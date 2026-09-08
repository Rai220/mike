import crypto from "node:crypto";
import { z } from "zod";
import type { ChatMessage } from "../chat/types";
import type { ReasoningLevel } from "../llm";
import type { Microsoft365AssistantDependencies } from "./assistantCapabilities";
import type { createServerSupabase } from "../supabase";
import { providerForModel, resolveModel } from "../llm/models";
import {
  getMicrosoft365ConnectionContext,
  getMicrosoft365Config,
  Microsoft365Error,
  openMicrosoft365Data,
  sealMicrosoft365Data,
} from "./index";
import type { SourceLocator } from "./sources";

type Db = ReturnType<typeof createServerSupabase>;
const TABLE = "microsoft365_chats";
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DAY = 86_400_000;

/** No implicit provider approval or retention policy. */
export function getMicrosoft365ChatPolicy(): {
  model: string;
  retentionDays: number;
} | null {
  const model = process.env.MICROSOFT365_CHAT_MODEL?.trim();
  const days = process.env.MICROSOFT365_RETENTION_DAYS ?? "";
  if (
    process.env.MICROSOFT365_CHAT_ENABLED !== "true" ||
    !model ||
    model.length > 200 ||
    !/^[1-9]\d?$/.test(days) ||
    Number(days) > 90
  )
    return null;
  try {
    if (
      resolveModel(model, "") !== model ||
      !["openai", "claude", "gemini", "ollama"].includes(
        providerForModel(model),
      ) ||
      (model.startsWith("ollama/") && !/^ollama\/[^\s]+$/.test(model))
    )
      return null;
  } catch {
    return null;
  }
  return { model, retentionDays: Number(days) };
}

export interface ProtectedChatPayload {
  messages: Array<{
    role: "user" | "assistant";
    content: string;
    sourceRefs: string[];
    files?: ChatMessage["files"];
    workflow?: ChatMessage["workflow"];
    useMicrosoft365?: boolean;
    useEdgar?: boolean;
    model?: string;
    reasoning?: ReasoningLevel;
  }>;
  sources: Array<{
    ref: string;
    locator: SourceLocator;
    version: string;
    title: string;
    webUrl?: string;
    fetchedAt: string;
  }>;
  preferences?: { model?: string; reasoning?: ReasoningLevel };
  dependencies?: Microsoft365AssistantDependencies;
}
export interface ProtectedChatMetadata {
  id: string;
  connectionId: string;
  createdAt: string;
  updatedAt: string;
  expiresAt: string;
  version: string;
  ordinaryChatId?: string;
}
/** Internal only: callers MUST revalidate every source ACL/version before exposing payload. */
export interface ProtectedChat extends ProtectedChatMetadata {
  payload: ProtectedChatPayload;
}
interface Row {
  ordinary_chat_id?: string | null;
  id: string;
  user_id: string;
  connection_id: string;
  org_id: string;
  tenant_id: string;
  generation: string;
  payload_ciphertext: string;
  created_at: string;
  updated_at: string;
  expires_at: string;
  version: string;
}
function policy() {
  const result = getMicrosoft365ChatPolicy();
  if (!result) throw new Microsoft365Error("unavailable", 503);
  return result;
}
function checked<T extends { error: unknown }>(result: T): T {
  if (result.error) throw new Microsoft365Error("storage_unavailable", 503);
  return result;
}
function aad(
  row: Pick<
    Row,
    "id" | "user_id" | "connection_id" | "org_id" | "tenant_id" | "generation" | "ordinary_chat_id"
  >,
) {
  return JSON.stringify([
    "mike-microsoft365-chat-v1",
    row.id,
    row.user_id,
    row.connection_id,
    row.org_id,
    row.tenant_id,
    row.generation,
    ...(row.ordinary_chat_id ? [row.ordinary_chat_id] : []),
  ]);
}
function metadata(row: Row, retentionDays: number): ProtectedChatMetadata {
  return {
    id: row.id,
    connectionId: row.connection_id,
    ...(row.ordinary_chat_id ? { ordinaryChatId: row.ordinary_chat_id } : {}),
    createdAt: row.created_at,
    updatedAt: row.updated_at,
    version: row.version,
    expiresAt: new Date(
      Math.min(
        Date.parse(row.expires_at),
        Date.parse(row.created_at) + retentionDays * DAY,
      ),
    ).toISOString(),
  };
}
const reasoningSchema = z.enum(["none", "low", "medium", "high", "xhigh", "max"]);
const preferencesSchema = z.object({ model: z.string().min(1).max(200).optional(), reasoning: reasoningSchema.optional() }).strict();
const composerMetadataSchema = z.object({
  files: z.array(z.object({ filename: z.string().min(1).max(1000), document_id: z.string().uuid(), version_id: z.string().uuid().optional(), version_number: z.number().int().positive().optional() }).strict()).max(20).optional(),
  workflow: z.object({ id: z.string().min(1).max(200), title: z.string().min(1).max(1000) }).strict().optional(),
  useMicrosoft365: z.boolean().optional(), useEdgar: z.boolean().optional(),
  model: z.string().min(1).max(200).optional(), reasoning: reasoningSchema.optional(),
}).passthrough();
const dependenciesSchema = z.object({
  documents: z.array(z.object({ id: z.string().uuid(), versionId: z.string().uuid(), version: z.string().min(1).max(200) }).strict()).max(100),
  workflows: z.array(z.object({ id: z.string().min(1).max(200), version: z.string().min(1).max(200) }).strict()).max(100),
}).strict();
function encode(payload: ProtectedChatPayload): string {
  if (
    !Array.isArray(payload.messages) ||
    payload.messages.length > 100 ||
    !Array.isArray(payload.sources) ||
    payload.sources.length > 140 ||
    payload.messages.some(
      (m) =>
        !["user", "assistant"].includes(m.role) ||
        typeof m.content !== "string" ||
        !Array.isArray(m.sourceRefs) || !composerMetadataSchema.safeParse(m).success ||
        m.sourceRefs.some((ref) => typeof ref !== "string"),
    ) ||
    payload.sources.some(
      (s) =>
        typeof s.ref !== "string" ||
        !s.locator ||
        typeof s.version !== "string" ||
        typeof s.title !== "string" ||
        typeof s.fetchedAt !== "string",
    )
  )
    throw new Microsoft365Error("invalid_state", 400);
  if ((payload.preferences !== undefined && !preferencesSchema.safeParse(payload.preferences).success) ||
      (payload.dependencies !== undefined && !dependenciesSchema.safeParse(payload.dependencies).success))
    throw new Microsoft365Error("invalid_state", 400);
  const refs = new Set(payload.sources.map((source) => source.ref));
  if (
    refs.size !== payload.sources.length ||
    payload.messages.some((message) =>
      message.sourceRefs.some((ref) => !refs.has(ref)),
    )
  )
    throw new Microsoft365Error("invalid_state", 400);
  const value = JSON.stringify(payload);
  if (Buffer.byteLength(value) > 1_048_576)
    throw new Microsoft365Error("invalid_state", 413);
  return value;
}
async function authorizedRow(userId: string, chatId: string, db: Db) {
  const currentPolicy = policy();
  if (!UUID.test(chatId)) throw new Microsoft365Error("not_found", 404);
  const { data } = checked(
    await db
      .from(TABLE)
      .select("*")
      .eq("id", chatId)
      .eq("user_id", userId)
      .maybeSingle(),
  );
  const row = data as Row | null;
  if (!row) throw new Microsoft365Error("not_found", 404);
  const context = await getMicrosoft365ConnectionContext(
    userId,
    row.connection_id,
    db,
  );
  if (
    context.orgId !== row.org_id ||
    context.tenantId !== row.tenant_id ||
    context.generation !== row.generation
  )
    throw new Microsoft365Error("access_denied", 403);
  const meta = metadata(row, currentPolicy.retentionDays);
  if (Date.parse(meta.expiresAt) <= Date.now()) {
    await deleteMicrosoft365Chat(userId, chatId, db);
    throw new Microsoft365Error("not_found", 404);
  }
  return { row, meta };
}

/** Application preflight; SQL repeats these checks under the shell row lock. */
export async function assertMicrosoft365OrdinaryChatPrivate(
  userId: string,
  ordinaryChatId: string,
  db: Db,
): Promise<{ microsoft365_protected: boolean }> {
  if (!UUID.test(ordinaryChatId)) throw new Microsoft365Error("not_found", 404);
  const { data: shell } = checked(await db.from("chats")
    .select("id,user_id,project_id,org_id,microsoft365_protected")
    .eq("id", ordinaryChatId).eq("user_id", userId).maybeSingle());
  if (!shell) throw new Microsoft365Error("not_found", 404);
  if (shell.project_id !== null || shell.org_id !== null)
    throw new Microsoft365Error("access_denied", 403);
  const { data: grants } = checked(await db.from("chat_access_grants")
    .select("id").eq("chat_id", ordinaryChatId).limit(1));
  if (grants?.length) throw new Microsoft365Error("access_denied", 403);
  return { microsoft365_protected: shell.microsoft365_protected === true };
}

/** Never resurrect expired/disconnected protected history as an ordinary chat. */
export async function findMicrosoft365ChatForOrdinaryChat(
  userId: string,
  ordinaryChatId: string,
  db: Db,
): Promise<ProtectedChat | null> {
  const shell = await assertMicrosoft365OrdinaryChatPrivate(userId, ordinaryChatId, db);
  const { data } = checked(await db.from(TABLE).select("id")
    .eq("ordinary_chat_id", ordinaryChatId).eq("user_id", userId).maybeSingle());
  if (!data) {
    if (shell.microsoft365_protected) throw new Microsoft365Error("not_found", 404);
    return null;
  }
  return loadMicrosoft365Chat(userId, data.id, db);
}

export async function createMicrosoft365Chat(
  userId: string,
  connectionId: string,
  db: Db,
  ordinaryChatId?: string,
): Promise<ProtectedChat> {
  const currentPolicy = policy();
  if (ordinaryChatId) {
    const existing = await findMicrosoft365ChatForOrdinaryChat(userId, ordinaryChatId, db);
    if (existing) {
      if (existing.connectionId !== connectionId) throw new Microsoft365Error("access_denied", 403);
      return existing;
    }
  }
  const context = await getMicrosoft365ConnectionContext(
    userId,
    connectionId,
    db,
  );
  const now = new Date().toISOString();
  const payload: ProtectedChatPayload = { messages: [], sources: [] };
  const row: Row = {
    ...(ordinaryChatId ? { ordinary_chat_id: ordinaryChatId } : {}),
    id: crypto.randomUUID(),
    user_id: userId,
    connection_id: context.id,
    org_id: context.orgId,
    tenant_id: context.tenantId,
    generation: context.generation,
    created_at: now,
    updated_at: now,
    expires_at: new Date(
      Date.now() + currentPolicy.retentionDays * DAY,
    ).toISOString(),
    version: crypto.randomUUID(),
    payload_ciphertext: "",
  };
  row.payload_ciphertext = sealMicrosoft365Data(encode(payload), aad(row));
  checked(await db.from(TABLE).insert(row));
  // Re-read after insertion to catch disconnect/membership changes during IO.
  const validated = await authorizedRow(userId, row.id, db);
  return { ...validated.meta, payload };
}

export async function listMicrosoft365Chats(
  userId: string,
  connectionId: string,
  db: Db,
): Promise<ProtectedChatMetadata[]> {
  const currentPolicy = policy();
  const context = await getMicrosoft365ConnectionContext(
    userId,
    connectionId,
    db,
  );
  const now = Date.now();
  const { data } = checked(
    await db
      .from(TABLE)
      .select("id,connection_id,created_at,updated_at,expires_at,version")
      .is("ordinary_chat_id", null)
      .eq("user_id", userId)
      .eq("connection_id", context.id)
      .eq("org_id", context.orgId)
      .eq("tenant_id", context.tenantId)
      .eq("generation", context.generation)
      .gt("expires_at", new Date(now).toISOString())
      .gt(
        "created_at",
        new Date(now - currentPolicy.retentionDays * DAY).toISOString(),
      )
      .order("updated_at", { ascending: false })
      .limit(100),
  );
  const latest = await getMicrosoft365ConnectionContext(
    userId,
    connectionId,
    db,
  );
  if (latest.generation !== context.generation)
    throw new Microsoft365Error("access_denied", 403);
  return ((data ?? []) as Row[]).map((row) =>
    metadata(row, currentPolicy.retentionDays),
  );
}

export async function loadMicrosoft365Chat(
  userId: string,
  chatId: string,
  db: Db,
): Promise<ProtectedChat> {
  const { row, meta } = await authorizedRow(userId, chatId, db);
  let payload: ProtectedChatPayload;
  try {
    payload = JSON.parse(
      openMicrosoft365Data(row.payload_ciphertext, aad(row)),
    ) as ProtectedChatPayload;
    encode(payload);
  } catch {
    throw new Microsoft365Error("storage_unavailable", 503);
  }
  return { ...meta, payload };
}

export async function saveMicrosoft365Chat(
  userId: string,
  chat: ProtectedChatMetadata,
  payload: ProtectedChatPayload,
  db: Db,
): Promise<ProtectedChat> {
  const { row, meta } = await authorizedRow(userId, chat.id, db);
  if (row.version !== chat.version) throw new Microsoft365Error("busy", 409);
  const updatedAt = new Date().toISOString();
  const version = crypto.randomUUID();
  const { data } = checked(
    await db
      .from(TABLE)
      .update({
        payload_ciphertext: sealMicrosoft365Data(encode(payload), aad(row)),
        updated_at: updatedAt,
        version,
        expires_at: meta.expiresAt,
      })
      .eq("id", row.id)
      .eq("user_id", userId)
      .eq("version", chat.version)
      .eq("generation", row.generation)
      .gt("expires_at", updatedAt)
      .select("id")
      .maybeSingle(),
  );
  if (!data) throw new Microsoft365Error("busy", 409);
  const validated = await authorizedRow(userId, row.id, db);
  if (validated.meta.version !== version)
    throw new Microsoft365Error("busy", 409);
  return { ...validated.meta, payload };
}

/** Owner can delete even after feature disablement or losing organization access. */
export async function deleteMicrosoft365Chat(
  userId: string,
  chatId: string,
  db: Db,
): Promise<void> {
  if (!UUID.test(chatId)) throw new Microsoft365Error("not_found", 404);
  checked(await db.from(TABLE).delete().eq("id", chatId).eq("user_id", userId));
}

/** Job payload contains no corporate data. Call periodically and on retention-policy reductions. */
export async function cleanupMicrosoft365Chats(db: Db): Promise<number> {
  const { data } = checked(
    await db.rpc("cleanup_microsoft365_chats", {
      retention_days: getMicrosoft365ChatPolicy()?.retentionDays ?? 90,
    }),
  );
  return typeof data === "number" ? data : 0;
}

/** Backend maintenance only. A cursor bounds each pass; rerun from the start after skipped races. */
export async function rotateMicrosoft365ChatEncryption(
  db: Db,
  afterId?: string,
): Promise<{
  rotated: number;
  examined: number;
  skipped: number;
  nextCursor: string | null;
}> {
  const config = getMicrosoft365Config();
  const activeKey = process.env.MICROSOFT365_ENCRYPTION_KEY_ID || "legacy";
  if (!config || !/^[a-zA-Z0-9_-]{1,32}$/.test(activeKey))
    throw new Microsoft365Error("unavailable", 503);
  if (afterId && !UUID.test(afterId))
    throw new Microsoft365Error("invalid_state", 400);
  let query = db
    .from(TABLE)
    .select("*")
    .eq("org_id", config.orgId)
    .eq("tenant_id", config.tenantId)
    .order("id", { ascending: true })
    .limit(100);
  if (afterId) query = query.gt("id", afterId);
  const { data } = checked(await query);
  const rows = (data ?? []) as Row[];
  let rotated = 0;
  let skipped = 0;
  for (const row of rows) {
    if (row.payload_ciphertext.startsWith(`v2.${activeKey}.`)) continue;
    try {
      const context = await getMicrosoft365ConnectionContext(
        row.user_id,
        row.connection_id,
        db,
      );
      if (
        context.orgId !== row.org_id ||
        context.tenantId !== row.tenant_id ||
        context.generation !== row.generation ||
        Date.parse(row.expires_at) <= Date.now()
      ) {
        skipped++;
        continue;
      }
      // Do not deserialize, log or put the cleartext into job payloads.
      const ciphertext = sealMicrosoft365Data(
        openMicrosoft365Data(row.payload_ciphertext, aad(row)),
        aad(row),
      );
      const { data: saved } = checked(
        await db
          .from(TABLE)
          .update({
            payload_ciphertext: ciphertext,
            version: crypto.randomUUID(),
          })
          .eq("id", row.id)
          .eq("user_id", row.user_id)
          .eq("version", row.version)
          .eq("generation", row.generation)
          .gt("expires_at", new Date().toISOString())
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
      ) {
        skipped++;
        continue;
      }
      throw error;
    }
  }
  return {
    rotated,
    examined: rows.length,
    skipped,
    nextCursor: rows.length === 100 ? rows[rows.length - 1].id : null,
  };
}
