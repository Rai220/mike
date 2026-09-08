import type { Request, Response } from "express";
import { requireMfaIfEnrolled } from "../middleware/auth";
import type { createServerSupabase } from "../lib/supabase";
import { getMicrosoft365Status, Microsoft365Error } from "../lib/microsoft365";
import {
  createMicrosoft365Chat,
  findMicrosoft365ChatForOrdinaryChat,
  getMicrosoft365ChatPolicy,
} from "../lib/microsoft365/protected";
import {
  getMicrosoft365AssistantChat,
  runMicrosoft365AssistantTurn,
  updateMicrosoft365AssistantPreferences,
} from "../lib/microsoft365/assistant";
import type { ChatMessage } from "../lib/chat/types";
import { getUserModelSettings } from "../lib/userSettings";
import { resolveEffectiveChatModel, resolveEffectiveReasoningLevel } from "../lib/modelSelection";
import { openAssistantSse } from "../lib/chat/routeStreaming";

type Db = ReturnType<typeof createServerSupabase>;
type Shell = { id: string; user_id: string | null; microsoft365_protected?: unknown };
const SAFE_ERROR = "Microsoft 365 could not complete this request. Please try again.";

export function microsoft365AssistantError(res: Response, error: unknown) {
  if (res.writableEnded || res.destroyed) return;
  const known = error instanceof Microsoft365Error;
  if (known && error.retryAfterSeconds) res.set("Retry-After", String(error.retryAfterSeconds));
  res.status(known ? error.status : 502).json({
    code: known ? error.code : "provider_unavailable",
    detail: known ? error.detail : SAFE_ERROR,
  });
}

async function verifyMfa(req: Request, res: Response): Promise<boolean> {
  let verified = false;
  await requireMfaIfEnrolled(req, res, () => { verified = true; });
  return verified;
}

function normalMessages(chat: Awaited<ReturnType<typeof getMicrosoft365AssistantChat>>) {
  return chat.messages.map((message, index) => ({
    id: `${chat.chatId}:${index}`,
    role: message.role,
    content: message.role === "assistant" ? [{ type: "content", text: message.content }] : message.content,
    files: message.files ?? null,
    workflow: message.workflow ?? null,
    useMicrosoft365: message.useMicrosoft365,
    useEdgar: message.useEdgar,
    model: message.model,
    reasoning: message.reasoning,
    citations: [],
    events: message.role === "assistant" ? [{ type: "content", text: message.content }] : null,
    microsoft365_protected: true,
  }));
}

/** Existing normal rows are pre-promotion history only; never trust the browser transcript. */
async function ordinaryHistory(chatId: string, db: Db) {
  const { data, error } = await db.from("chat_messages")
    .select("role,content").eq("chat_id", chatId).order("created_at", { ascending: true });
  if (error) throw new Microsoft365Error("storage_unavailable", 503);
  return (data ?? []).filter((row: any) => row.role === "user" || row.role === "assistant")
    .map((row: any) => ({
      role: row.role as "user" | "assistant",
      content: typeof row.content === "string" && row.content ? row.content :
        (Array.isArray(row.content) ? row.content.filter((event: any) => event.type === "content" && typeof event.text === "string").map((event: any) => event.text).join("\n") : ""),
    }));
}

export async function readMicrosoft365OrdinaryChat(
  req: Request, res: Response, shell: Shell, db: Db,
) {
  res.set("Cache-Control", "no-store");
  if (shell.user_id !== res.locals.userId) throw new Microsoft365Error("not_found", 404);
  if (!await verifyMfa(req, res)) return null;
  const chat = await findMicrosoft365ChatForOrdinaryChat(res.locals.userId, shell.id, db);
  // A permanent shell must never downgrade after expiry, disconnect or sidecar deletion.
  if (!chat) throw new Microsoft365Error("not_found", 404);
  const loaded = await getMicrosoft365AssistantChat(res.locals.userId, chat.id, db, AbortSignal.timeout(60_000), res.locals.userEmail);
  return {
    messages: normalMessages(loaded),
    expiresAt: loaded.expiresAt,
    model: loaded.model,
    reasoning: loaded.reasoning,
  };
}

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** Authorization of these references belongs to the protected capability runner. */
function validAttachments(message: ChatMessage): boolean {
  return (!message.files || (Array.isArray(message.files) && message.files.every((file) =>
    file && typeof file.filename === "string" && typeof file.document_id === "string" && UUID.test(file.document_id) &&
    (file.version_id === undefined || UUID.test(file.version_id))))) &&
    (!message.workflow || (typeof message.workflow.id === "string" && (UUID.test(message.workflow.id) || /^builtin-[a-z0-9]+(?:-[a-z0-9]+)*$/.test(message.workflow.id))));
}

export async function updateMicrosoft365OrdinaryChat(
  req: Request, res: Response, shell: Shell, db: Db,
  input: { model?: string; reasoning?: string },
) {
  res.set("Cache-Control", "no-store");
  const userId = res.locals.userId as string;
  if (shell.user_id !== userId) throw new Microsoft365Error("not_found", 404);
  if (!await verifyMfa(req, res)) return null;
  const chat = await findMicrosoft365ChatForOrdinaryChat(userId, shell.id, db);
  if (!chat) throw new Microsoft365Error("not_found", 404);
  const settings = await getUserModelSettings(userId, db);
  const policy = getMicrosoft365ChatPolicy();
  if (!policy) throw new Microsoft365Error("unavailable", 503);
  const policyFallback = !input.model && !chat.payload.preferences?.model && !settings.last_selected_chat_model;
  const resolution = policyFallback ? { ok: true as const, model: policy.model } : await resolveEffectiveChatModel({
    requested: input.model ?? chat.payload.preferences?.model,
    chatModel: chat.payload.preferences?.model,
    lastSelectedModel: settings.last_selected_chat_model,
    apiKeys: settings.api_keys, userId, db,
  });
  if (!resolution.ok) {
    res.status(resolution.status).json({ code: resolution.code, detail: resolution.detail });
    return null;
  }
  const reasoning = resolveEffectiveReasoningLevel({
    model: resolution.model, requested: input.reasoning,
    chatReasoningLevel: chat.payload.preferences?.reasoning,
    lastSelectedReasoningLevel: settings.last_selected_reasoning_level,
  });
  const updated = await updateMicrosoft365AssistantPreferences(
    userId, chat.id, db, { model: resolution.model, reasoning },
    AbortSignal.timeout(60_000), res.locals.userEmail,
  );
  return { id: shell.id, title: "Microsoft 365", model: updated.model, reasoning_level: updated.reasoning };
}

/** Runs before the ordinary route's plaintext persistence, title model, tool loop and audit. */
export async function streamMicrosoft365OrdinaryChat(
  req: Request, res: Response,
  input: { db: Db; chatId: string | null; alreadyProtected: boolean; enabled: boolean; projectId: string | null; message: ChatMessage | undefined; askInputsResponse?: unknown; model?: string; reasoning?: string; useEdgar?: boolean; chatModel?: string | null; chatReasoningLevel?: string | null },
): Promise<void> {
  res.set("Cache-Control", "no-store");
  try {
    if (!await verifyMfa(req, res)) return;
    const { db, enabled, message } = input;
    const userId = res.locals.userId as string;
    if (input.projectId) {
      res.status(409).json({ code: "microsoft365_private_chat_required", detail: "Microsoft 365 access is available in personal chats only." });
      return;
    }
    if (!message?.content?.trim() || message.content.length > 8000 ||
        !validAttachments(message) || input.askInputsResponse ||
        Buffer.byteLength(JSON.stringify(req.body)) > 65_536) {
      res.status(400).json({ code: "invalid_query", detail: "Send a question of up to 8,000 characters with valid document and workflow references." });
      return;
    }
    const policy = getMicrosoft365ChatPolicy();
    if (!policy) throw new Microsoft365Error("unavailable", 503);
    const status = await getMicrosoft365Status(userId, db);
    if (!status.available || status.connection?.status !== "connected") throw new Microsoft365Error("reconnect_required", 409);
    let shellId = input.chatId;
    if (!shellId) {
      const { data, error } = await db.from("chats").insert({ user_id: userId, project_id: null, org_id: null, title: "Microsoft 365", model: policy.model }).select("id").single();
      if (error || !data) throw new Microsoft365Error("storage_unavailable", 503);
      shellId = data.id as string;
    }
    let protectedChat = await findMicrosoft365ChatForOrdinaryChat(userId, shellId, db);
    if (!protectedChat) {
      if (input.alreadyProtected || !enabled) throw new Microsoft365Error("not_found", 404);
      protectedChat = await createMicrosoft365Chat(userId, status.connection.id, db, shellId);
    }
    const settings = await getUserModelSettings(userId, db);
    const preferences = protectedChat.payload.preferences;
    const chatModel = preferences?.model ?? (input.alreadyProtected ? null : input.chatModel);
    const policyFallback = !input.model && !chatModel && !settings.last_selected_chat_model;
    const resolution = policyFallback ? { ok: true as const, model: policy.model } : await resolveEffectiveChatModel({
      requested: input.model ?? preferences?.model,
      chatModel,
      lastSelectedModel: settings.last_selected_chat_model,
      apiKeys: settings.api_keys, userId, db,
    });
    if (!resolution.ok) {
      res.status(resolution.status).json({ code: resolution.code, detail: resolution.detail });
      return;
    }
    const model = resolution.model;
    const reasoning = resolveEffectiveReasoningLevel({
      model, requested: input.reasoning,
      chatReasoningLevel: preferences?.reasoning ?? input.chatReasoningLevel,
      lastSelectedReasoningLevel: settings.last_selected_reasoning_level,
    });
    const history = await ordinaryHistory(shellId, db);
    const stream = openAssistantSse(res, { sensitive: true });
    const emit = (value: unknown) => stream.write(`data: ${JSON.stringify(value)}\n\n`);
    try {
      emit({ type: "chat_id", chatId: shellId });
      emit({ type: "chat_title", chatId: shellId, title: "Microsoft 365" });
      emit({ type: "microsoft365", protected: true, enabled, expiresAt: protectedChat.expiresAt, model, reasoning });
      const answer = await runMicrosoft365AssistantTurn({
        userId, chatId: protectedChat.id, db, message: message.content, enabled,
        signal: AbortSignal.any([stream.signal, AbortSignal.timeout(120_000)]), ordinaryHistory: history,
        assistantOptions: {
          userEmail: res.locals.userEmail,
          files: message.files, workflow: message.workflow, useEdgar: input.useEdgar !== false,
          model, reasoning, apiKeys: settings.api_keys,
        },
      });
      if (stream.signal.aborted) return;
      const last = answer.messages.at(-1);
      if (!last || last.role !== "assistant") throw new Microsoft365Error("provider_unavailable", 502);
      emit({ type: "content", text: last.content });
      emit({ type: "citations", status: "final", citations: [] });
      stream.write("data: [DONE]\n\n");
    } catch (error) {
      if (!stream.signal.aborted) {
        emit({ type: "error", safe_to_display: true, message: error instanceof Microsoft365Error ? error.detail : SAFE_ERROR });
        stream.write("data: [DONE]\n\n");
      }
    } finally { stream.finish(); }
  } catch (error) { microsoft365AssistantError(res, error); }
}
