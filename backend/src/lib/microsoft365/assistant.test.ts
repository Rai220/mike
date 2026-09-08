import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  getMicrosoft365AssistantChat,
  getMicrosoft365AssistantSource,
  runMicrosoft365AssistantTurn,
  updateMicrosoft365AssistantPreferences,
} from "./assistant";
import {
  getMicrosoft365ChatPolicy,
  loadMicrosoft365Chat,
  saveMicrosoft365Chat,
  type ProtectedChat,
} from "./protected";
import {
  readMicrosoft365Source,
  readMicrosoft365SourceMetadata,
  searchMicrosoft365Sources,
} from "./sources";
import { streamChatWithTools, type StreamChatParams } from "../llm";
import { Microsoft365Error } from "./index";
import { normalizeMicrosoft365MailMetadata } from "./mailMetadata";
import { listMicrosoft365MailAttachments } from "./mailAttachments";
vi.mock("./mailAttachments", () => ({ listMicrosoft365MailAttachments: vi.fn() }));
import { buildMicrosoft365AssistantCapabilities, validateMicrosoft365AssistantDependencies } from "./assistantCapabilities";
vi.mock("./assistantCapabilities", () => ({ buildMicrosoft365AssistantCapabilities: vi.fn(), validateMicrosoft365AssistantDependencies: vi.fn() }));
vi.mock("./protected", () => ({
  getMicrosoft365ChatPolicy: vi.fn(),
  loadMicrosoft365Chat: vi.fn(),
  saveMicrosoft365Chat: vi.fn(),
}));
vi.mock("./sources", () => ({
  readMicrosoft365Source: vi.fn(),
  readMicrosoft365SourceMetadata: vi.fn(),
  searchMicrosoft365Sources: vi.fn(),
}));
vi.mock("../llm", () => ({ streamChatWithTools: vi.fn() }));
const db = {} as Parameters<typeof loadMicrosoft365Chat>[2];
const sourceRef = "2f8b84cc-b410-4705-82d4-c966bc5c2349";
const source = {
  locator: { kind: "mail" as const, id: "immutable-id" },
  title: "Contract renewal",
  version: "graph-v1",
  graphVersion: "graph-v1",
  fetchedAt: "2026-09-07T00:00:00Z",
  text: "The contract renews on 30 September 2026.",
};
let chat: ProtectedChat;
function answer(
  text = "Here are the latest messages.",
  citations: Array<{ ref: string; quote: string }> = [],
) {
  return { fullText: JSON.stringify({ answer: text, citations }) };
}
const turn = (
  enabled = true,
  extra: Partial<Parameters<typeof runMicrosoft365AssistantTurn>[0]> = {},
) =>
  runMicrosoft365AssistantTurn({
    userId: "owner",
    chatId: "protected",
    db,
    message: "Какие письма у меня на почте?",
    enabled,
    ...extra,
  });
function history() {
  chat.payload.sources = [
    {
      ref: sourceRef,
      locator: source.locator,
      title: source.title,
      version: source.version,
      fetchedAt: source.fetchedAt,
    },
  ];
  chat.payload.messages = [
    { role: "user", content: "List mail", sourceRefs: [sourceRef] },
    { role: "assistant", content: "Contract renewal", sourceRefs: [sourceRef] },
  ];
}
async function list(params: StreamChatParams) {
  return JSON.parse(
    (
      await params.runTools!([
        { id: "list", name: "microsoft365_recent_mail", input: {} },
      ])
    )[0]!.content,
  );
}
beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(listMicrosoft365MailAttachments).mockResolvedValue({ items: [], more: false });
  chat = {
    id: "protected",
    connectionId: "connection",
    createdAt: "2026-09-07",
    updatedAt: "2026-09-07",
    expiresAt: "2026-09-14",
    version: "cas1",
    payload: { messages: [], sources: [] },
  };
  vi.mocked(loadMicrosoft365Chat).mockImplementation(async () =>
    structuredClone(chat),
  );
  vi.mocked(saveMicrosoft365Chat).mockImplementation(
    async (_user, original, payload) => ({
      ...original,
      payload,
      version: "cas2",
    }),
  );
  vi.mocked(getMicrosoft365ChatPolicy).mockReturnValue({
    model: "claude-sonnet-4-6",
    retentionDays: 7,
  });
  vi.mocked(readMicrosoft365SourceMetadata).mockResolvedValue(source);
  vi.mocked(readMicrosoft365Source).mockResolvedValue(source);
  vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
    items: [{ locator: source.locator, title: source.title }],
    more: false,
    nextOffset: 20,
  });
  vi.mocked(streamChatWithTools).mockResolvedValue(answer());
});
describe("main Microsoft 365 assistant", () => {
  it("reads mail recipients and exposes attachment refs, then reads and cites the attachment", async () => {
    history();
    const attachment = {
      locator: { kind: "mail" as const, id: source.locator.id, attachmentId: "attachment-id" },
      title: "Budget.txt", version: "attachment-v1", graphVersion: "attachment-v1", fetchedAt: source.fetchedAt,
      attachment: { name: "Budget.txt", contentType: "text/plain", size: 123, isInline: false, type: "file" as const, readable: true },
    };
    const mail = normalizeMicrosoft365MailMetadata({ toRecipients: [{ emailAddress: { name: "Recipient", address: "recipient@example.test" } }], ccRecipients: [], isRead: false });
    vi.mocked(listMicrosoft365MailAttachments).mockResolvedValue({ items: [attachment], more: false });
    vi.mocked(readMicrosoft365Source).mockImplementation(async (_user, _connection, _db, locator) =>
      locator.kind === "mail" && locator.attachmentId
        ? { ...attachment, text: "The attached budget is 120 units." }
        : { ...source, mail, attachments: { items: [attachment], more: false } });
    vi.mocked(readMicrosoft365SourceMetadata).mockImplementation(async (_user, _connection, _db, locator) => locator.kind === "mail" && locator.attachmentId ? attachment : source);
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const parent = JSON.parse((await params.runTools!([{ id: "parent", name: "microsoft365_read", input: { ref: sourceRef } }]))[0]!.content);
      expect(parent.mail).toMatchObject({ toRecipients: [{ name: "Recipient", address: "recipient@example.test" }], ccRecipients: [], bccRecipients: null });
      expect(parent.attachments).toEqual([expect.objectContaining({ name: "Budget.txt", ref: expect.any(String), readable: true })]);
      expect(parent.attachments[0]).not.toHaveProperty("locator");
      expect(parent.attachments[0].ref).not.toBe("attachment-id");
      const ref = parent.attachments[0].ref;
      const child = JSON.parse((await params.runTools!([{ id: "child", name: "microsoft365_read", input: { ref } }]))[0]!.content);
      expect(child.text).toBe("The attached budget is 120 units.");
      return answer("The message and attachment were read", [{ ref: sourceRef, quote: "recipient@example.test" }, { ref, quote: child.text }]);
    });
    const result = await turn(true, { message: "Кому отправлено письмо и что во вложении?" });
    expect(result.sources).toHaveLength(2);
    expect(result.messages.at(-1)?.content).toContain("recipient@example.test");
    expect(result.messages.at(-1)?.content).toContain("120 units");
    const saved = vi.mocked(saveMicrosoft365Chat).mock.calls[0]![2];
    expect(saved.sources[1].locator).toEqual(attachment.locator);
    expect(saved.sources[1]).not.toHaveProperty("text");
    expect(saved.messages.at(-1)?.sourceRefs).toHaveLength(2);
  });
  it("accepts exact short metadata values but not invented recipients", async () => {
    history();
    const mail = normalizeMicrosoft365MailMetadata({ isRead: false, toRecipients: [{ emailAddress: { address: "a@b.co", name: 'A "quoted" name' } }], internetMessageHeaders: [{ name: "X-Synthetic", value: "first\nsecond" }] });
    vi.mocked(readMicrosoft365Source).mockResolvedValue({ ...source, mail });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([{ id: "read", name: "microsoft365_read", input: { ref: sourceRef } }]);
      return answer("Mail metadata", [{ ref: sourceRef, quote: "a@b.co" }, { ref: sourceRef, quote: "false" }, { ref: sourceRef, quote: 'A "quoted" name' }, { ref: sourceRef, quote: "first\nsecond" }]);
    });
    expect((await turn()).messages.at(-1)?.content).toContain("a@b.co");
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([{ id: "read", name: "microsoft365_read", input: { ref: sourceRef } }]);
      return answer("Forged recipient", [{ ref: sourceRef, quote: "forged@example.test" }]);
    });
    await expect(turn()).rejects.toMatchObject({ code: "provider_unavailable" });
  });
  it.each(["unchanged", "removed", "changed"])("revalidates sibling attachment metadata once per parent: %s", async (state) => {
    history();
    const children = ["one", "two"].map((id) => ({ ...source, ref: id,
      locator: { kind: "mail" as const, id: source.locator.id, attachmentId: id },
      attachment: { name: `${id}.txt`, contentType: "text/plain", size: 10, isInline: false, type: "file" as const, readable: true },
    }));
    chat.payload.sources.push(...children);
    chat.payload.messages[0].sourceRefs.push(...children.map((child) => child.ref));
    vi.mocked(listMicrosoft365MailAttachments).mockResolvedValue({ items: state === "removed" ? children.slice(0, 1) : children.map((child, i) => state === "changed" && i === 1 ? { ...child, version: "changed" } : child), more: false });
    const result = getMicrosoft365AssistantChat("owner", "protected", db);
    if (state === "unchanged") expect((await result).sources).toHaveLength(3);
    else await expect(result).rejects.toMatchObject({ code: state === "removed" ? "not_found" : "source_changed" });
    expect(listMicrosoft365MailAttachments).toHaveBeenCalledTimes(1);
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(1);
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
  });
  it("accepts a complete short listed attachment filename as mail evidence", async () => {
    history();
    const child = { ...source, locator: { ...source.locator, attachmentId: "file" },
      attachment: { name: "a.pdf", contentType: "application/pdf", size: 10, isInline: false, type: "file" as const, readable: true } };
    vi.mocked(listMicrosoft365MailAttachments).mockResolvedValue({ items: [child], more: false });
    vi.mocked(readMicrosoft365Source).mockResolvedValue({ ...source, attachments: { items: [child], more: false } });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([{ id: "read", name: "microsoft365_read", input: { ref: sourceRef } }]);
      return answer("Attachment listed", [{ ref: sourceRef, quote: "a.pdf" }]);
    });
    expect((await turn()).messages.at(-1)?.content).toContain("a.pdf");
  });
  it("blocks attachment content retrieval while Microsoft365 is off", async () => {
    history();
    chat.payload.sources[0].locator = { kind: "mail", id: "parent", attachmentId: "file" };
    vi.mocked(listMicrosoft365MailAttachments).mockResolvedValue({ items: [{ ...source, locator: chat.payload.sources[0].locator, attachment: { name: "file.txt", contentType: "text/plain", size: 10, isInline: false, type: "file", readable: true } }], more: false });
    await turn(false);
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    expect(vi.mocked(streamChatWithTools).mock.calls[0]![0]).not.toHaveProperty("runTools");
  });
  it.each([true, false])("keeps normal capabilities and selected model independent when Microsoft365 is %s", async (enabled) => {
    const file = { filename: "Synthetic.txt", document_id: "11111111-1111-4111-8111-111111111111", version_id: "22222222-2222-4222-8222-222222222222" };
    const dependencies = { documents: [{ id: file.document_id, versionId: file.version_id, version: "v1" }], workflows: [{ id: "workflow", version: "w1" }] };
    const execute = vi.fn().mockResolvedValue([{ tool_use_id: "edgar", content: '{"ok":true}' }]);
    const validate = vi.fn().mockResolvedValue(undefined);
    vi.mocked(buildMicrosoft365AssistantCapabilities).mockResolvedValue({
      tools: [{ type: "function", function: { name: "edgar_search_filings", description: "Search", parameters: { type: "object", properties: {} } } }],
      contextPrompt: "Selected workflow instructions", files: [file], workflow: { id: "workflow", title: "Test workflow" }, execute, validate, dependencies: () => dependencies,
    });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      expect(params).toMatchObject({ model: "gpt-5.4", reasoning: "low", apiKeys: { openai: "synthetic-server-loaded-key" }, sensitive: true, sensitiveTools: "microsoft365-assistant" });
      expect(params.systemPrompt).toContain("Selected workflow instructions");
      expect(params.tools?.some((tool) => tool.function.name === "microsoft365_recent_mail")).toBe(enabled);
      await params.runTools!([{ id: "edgar", name: "edgar_search_filings", input: { query: "Apple" } }]);
      if (enabled) await list(params);
      return answer();
    });
    const result = await turn(enabled, { assistantOptions: { files: [file], workflow: { id: "workflow", title: "Test workflow" }, useEdgar: true, model: "gpt-5.4", reasoning: "low", apiKeys: { openai: "synthetic-server-loaded-key" } } });
    expect(execute).toHaveBeenCalled();
    expect(validate).toHaveBeenCalled();
    expect(result.model).toBe("gpt-5.4");
    const payload = vi.mocked(saveMicrosoft365Chat).mock.calls[0]![2];
    expect(payload).toMatchObject({ preferences: { model: "gpt-5.4", reasoning: "low" }, dependencies, messages: [expect.objectContaining({ files: [file], workflow: { id: "workflow", title: "Test workflow" }, useMicrosoft365: enabled, useEdgar: true }), expect.any(Object)] });
    expect(JSON.stringify(payload)).not.toContain("synthetic-server-loaded-key");
    if (!enabled) expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
  });
  it("fails closed on a revoked local attachment before disclosing encrypted history", async () => {
    chat.payload.dependencies = { documents: [{ id: "document", versionId: "version", version: "hash" }], workflows: [] };
    vi.mocked(validateMicrosoft365AssistantDependencies).mockRejectedValue(new Microsoft365Error("access_denied", 403));
    await expect(getMicrosoft365AssistantChat("owner", "protected", db)).rejects.toMatchObject({ code: "access_denied" });
    expect(streamChatWithTools).not.toHaveBeenCalled();
  });
  it("updates model preferences in the encrypted payload without losing history", async () => {
    history();
    const result = await updateMicrosoft365AssistantPreferences("owner", "protected", db, { model: "gpt-5.4", reasoning: "low" });
    expect(result).toMatchObject({ model: "gpt-5.4", reasoning: "low" });
    expect(vi.mocked(saveMicrosoft365Chat).mock.calls[0]![2]).toMatchObject({ messages: chat.payload.messages, preferences: { model: "gpt-5.4", reasoning: "low" } });
    expect(streamChatWithTools).not.toHaveBeenCalled();
  });
  it.each([undefined, "all", "sentitems", "inbox"] as const)("passes recent mail scope %s and labels results", async (folder) => {
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const results = await params.runTools!([{ id: "folder", name: "microsoft365_recent_mail", input: folder ? { folder } : {} }]);
      expect(JSON.parse(results[0]!.content).folder).toBe(folder ?? "all");
      expect(params.tools?.find((tool) => tool.function.name === "microsoft365_recent_mail")?.function.parameters).toMatchObject({
        properties: { folder: { enum: ["all", "inbox", "sentitems"] } },
      });
      return answer();
    });
    await turn(true, { message: "Какие письма у меня в отправленных?" });
    expect(searchMicrosoft365Sources).toHaveBeenCalledWith("owner", "connection", db,
      expect.objectContaining({ kind: "mail", mode: "recent", folder: folder ?? "all" }), expect.any(AbortSignal));
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    expect(vi.mocked(streamChatWithTools).mock.calls[0]![0].systemPrompt).toContain("for sent/outgoing mail (Отправленные) use folder=sentitems");
  });
  it("rejects an unsupported mail folder before Graph retrieval", async () => {
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([{ id: "folder", name: "microsoft365_recent_mail", input: { folder: "../other" } }]);
      return answer();
    });
    await expect(turn()).rejects.toMatchObject({ code: "invalid_query" });
    expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("answers inbox questions through allowlisted tools and saves listed metadata dependencies without bodies", async () => {
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const found = await list(params);
      expect(found.items).toEqual([
        { ref: expect.any(String), title: source.title, kind: "mail" },
      ]);
      expect(found.items[0]).not.toHaveProperty("locator");
      return answer(`Последнее письмо: ${found.items[0].title}`);
    });
    const result = await turn();
    const params = vi.mocked(streamChatWithTools).mock.calls[0]![0];
    expect(params).toMatchObject({
      model: "claude-sonnet-4-6",
      sensitive: true,
      sensitiveTools: "microsoft365",
      maxIterations: 6,
    });
    expect(params.systemPrompt).toContain("You are Mike");
    expect(params.systemPrompt).toContain("call recent mail immediately");
    expect(params.tools?.map((tool) => tool.function.name)).toEqual([
      "microsoft365_recent_mail",
      "microsoft365_search",
      "microsoft365_read",
    ]);
    expect(params).not.toHaveProperty("callbacks");
    expect(params).not.toHaveProperty("apiKeys");
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(2);
    expect(result.sources).toHaveLength(1);
    const saved = vi.mocked(saveMicrosoft365Chat).mock.calls[0]![2];
    expect(
      saved.messages.every((message) =>
        message.sourceRefs.includes(result.sources[0]!.ref),
      ),
    ).toBe(true);
    expect(saved.sources[0]).not.toHaveProperty("text");
    expect(JSON.stringify(saved)).not.toContain(source.text);
  });
  it("reads a listed opaque ref on followup and checks exact evidence", async () => {
    history();
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const found = JSON.parse(
        (
          await params.runTools!([
            {
              id: "read",
              name: "microsoft365_read",
              input: { ref: sourceRef },
            },
          ])
        )[0]!.content,
      );
      expect(found.text).toEqual(source.text);
      return answer("The renewal is in September.", [
        { ref: sourceRef, quote: source.text },
      ]);
    });
    const result = await turn(true, { message: "Что в первом письме?" });
    expect(result.messages).toHaveLength(4);
    expect(result.messages[3]!.content).toContain(source.text);
    expect(readMicrosoft365Source).toHaveBeenCalledTimes(1);
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(2);
  });
  it.each(["Contract renewal", "Contract"])(
    "accepts a metadata-only title quotation from a listed source: %s",
    async (quote) => {
      vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
        const found = await list(params);
        return answer("Последнее письмо: Contract renewal", [
          { ref: found.items[0].ref, quote },
        ]);
      });
      const result = await turn();
      expect(result.messages[1]!.content).toContain(`“${quote}”`);
      expect(readMicrosoft365Source).not.toHaveBeenCalled();
      expect(result.messages[1]!.sourceRefs).toEqual([result.sources[0]!.ref]);
    },
  );
  it.each(["unknown-ref", "forged-title", "body-not-read"])(
    "rejects invalid metadata-only citations: %s",
    async (kind) => {
      vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
        const found = await list(params);
        return answer("Unsupported", [
          {
            ref: kind === "unknown-ref" ? sourceRef : found.items[0].ref,
            quote:
              kind === "body-not-read"
                ? source.text
                : kind === "forged-title"
                  ? "Fabricated subject"
                  : source.title,
          },
        ]);
      });
      await expect(turn()).rejects.toMatchObject({
        code: "provider_unavailable",
      });
      expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
      expect(readMicrosoft365Source).not.toHaveBeenCalled();
    },
  );
  it("accepts the complete exact short subject as metadata evidence", async () => {
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({ items: [{ locator: source.locator, title: "Hello" }], more: false, nextOffset: 20 });
    vi.mocked(readMicrosoft365SourceMetadata).mockResolvedValue({ ...source, title: "Hello" });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const found = await list(params);
      return answer("Here is the sent message.", [{ ref: found.items[0].ref, quote: "Hello" }]);
    });
    const result = await turn();
    expect(result.messages[1]!.content).toContain('“Hello”');
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
  });
  it.each(["Con", " "])("rejects a short partial or blank title quote %j", async (quote) => {
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const found = await list(params);
      return answer("A message", [{ ref: found.items[0].ref, quote }]);
    });
    await expect(turn()).rejects.toMatchObject({ code: "provider_unavailable" });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("still rejects short body quotations even when they match the entire source", async () => {
    history();
    vi.mocked(readMicrosoft365Source).mockResolvedValue({ ...source, text: "Hello", title: "Hello" });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([{ id: "read", name: "microsoft365_read", input: { ref: sourceRef } }]);
      return answer("The message says hello", [{ ref: sourceRef, quote: "Hello" }]);
    });
    await expect(turn()).rejects.toMatchObject({ code: "provider_unavailable" });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("does not accept a title quotation as body evidence after a read", async () => {
    history();
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([
        { id: "read", name: "microsoft365_read", input: { ref: sourceRef } },
      ]);
      return answer("Unsupported body claim", [
        { ref: sourceRef, quote: source.title },
      ]);
    });
    await expect(turn()).rejects.toMatchObject({
      code: "provider_unavailable",
    });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("exposes zero retrieval tools while OFF, preserves history and performs metadata-only revalidation", async () => {
    history();
    await turn(false, { message: "Напомни тему письма" });
    const params = vi.mocked(streamChatWithTools).mock.calls[0]![0];
    for (const field of [
      "tools",
      "runTools",
      "sensitiveTools",
      "apiKeys",
      "callbacks",
    ])
      expect(params).not.toHaveProperty(field);
    expect(params.systemPrompt).toContain("Microsoft 365 is OFF");
    expect(params.messages[1]!.content).toEqual("Contract renewal");
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(2);
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
  });
  it("accepts a bounded plain-language OFF reminder without exposing retrieval tools", async () => {
    history();
    const content =
      "Для обновления списка писем включите Microsoft 365 в этом чате. Сейчас могу обсудить уже показанные письма.";
    vi.mocked(streamChatWithTools).mockResolvedValue({
      fullText: ` ${content}\n`,
    });
    const result = await turn(false, {
      message: "Обнови список последних писем.",
    });
    expect(result.messages.at(-1)?.content).toEqual(content);
    const params = vi.mocked(streamChatWithTools).mock.calls[0]![0];
    expect(params).not.toHaveProperty("tools");
    expect(params).not.toHaveProperty("runTools");
    expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(2);
  });
  it.each([
    '{"answer": "unfinished',
    '[{"answer":',
    '```json\n{"answer":',
    "```unfinished",
    " ",
    "x".repeat(20_001),
  ])(
    "rejects malformed protocol or unbounded plain OFF output: case %#",
    async (fullText) => {
      vi.mocked(streamChatWithTools).mockResolvedValue({ fullText });
      await expect(turn(false)).rejects.toMatchObject({
        code: "provider_unavailable",
      });
      expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
    },
  );
  it("still requires verified structured output when Microsoft 365 is ON", async () => {
    vi.mocked(streamChatWithTools).mockResolvedValue({
      fullText: "A plain unverified inbox answer",
    });
    await expect(turn(true)).rejects.toMatchObject({
      code: "provider_unavailable",
    });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("handles an empty inbox without requiring a source quotation", async () => {
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
      items: [],
      more: false,
      nextOffset: 20,
    });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await list(params);
      return answer("Входящих писем нет.");
    });
    expect((await turn()).sources).toEqual([]);
    expect(readMicrosoft365SourceMetadata).not.toHaveBeenCalled();
  });
  it("searches files and preserves raw Graph versions independently of content hashes", async () => {
    const file = {
      locator: { kind: "file" as const, id: "file-id", driveId: "drive-id" },
      title: "Contract.docx",
      version: "etag1",
      graphVersion: "etag1",
      fetchedAt: source.fetchedAt,
    };
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
      items: [{ locator: file.locator, title: file.title }],
      more: false,
      nextOffset: 20,
    });
    vi.mocked(readMicrosoft365SourceMetadata).mockResolvedValue(file);
    vi.mocked(readMicrosoft365Source).mockResolvedValue({
      ...file,
      version: "sha256bytes",
      text: source.text,
    });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const found = JSON.parse(
        (
          await params.runTools!([
            {
              id: "search",
              name: "microsoft365_search",
              input: { kind: "file", query: "Contract" },
            },
          ])
        )[0]!.content,
      );
      const ref = found.items[0].ref;
      await params.runTools!([
        { id: "read", name: "microsoft365_read", input: { ref } },
      ]);
      return answer("September", [{ ref, quote: source.text }]);
    });
    expect((await turn()).sources[0]!.version).toEqual("etag1");
  });
  it.each([
    { name: "microsoft365_read", input: { ref: "https://evil.test/body" } },
    {
      name: "microsoft365_read",
      input: { ref: "4f886b93-06a8-4ec4-a83f-fd7e91f2ccab" },
    },
    {
      name: "microsoft365_read",
      input: { ref: sourceRef, url: "https://evil.test" },
    },
    { name: "send_email", input: { to: "attacker" } },
    {
      name: "microsoft365_search",
      input: { kind: "file", query: "Contract", url: "https://evil.test" },
    },
  ])(
    "rejects unowned refs, extra parameters and arbitrary tools: $name $input",
    async (call) => {
      history();
      vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
        // Simulate an SDK that absorbs a tool error and still offers final text.
        await params.runTools!([{ id: "bad", ...call }]).catch(() => undefined);
        return answer("Should be discarded");
      });
      await expect(turn()).rejects.toBeInstanceOf(Microsoft365Error);
      expect(readMicrosoft365Source).not.toHaveBeenCalled();
      expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
      expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
    },
  );
  it("blocks revoked history before provider calls and public hydration, including OFF", async () => {
    history();
    vi.mocked(readMicrosoft365SourceMetadata).mockRejectedValue(
      new Microsoft365Error("access_denied", 403),
    );
    await expect(turn(false)).rejects.toMatchObject({ code: "access_denied" });
    await expect(
      getMicrosoft365AssistantChat("owner", "protected", db),
    ).rejects.toMatchObject({ code: "access_denied" });
    expect(streamChatWithTools).not.toHaveBeenCalled();
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("discards all output when a listed item's ACL is revoked after generation", async () => {
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await list(params);
      vi.mocked(readMicrosoft365SourceMetadata).mockRejectedValue(
        new Microsoft365Error("access_denied", 403),
      );
      return answer("Private metadata");
    });
    await expect(turn()).rejects.toMatchObject({ code: "access_denied" });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it.each(["version", "membership", "policy", "cas"])(
    "discards generated output when %s changes",
    async (change) => {
      history();
      vi.mocked(streamChatWithTools).mockImplementation(async () => {
        if (change === "version")
          vi.mocked(readMicrosoft365SourceMetadata).mockResolvedValue({
            ...source,
            version: "v2",
          });
        if (change === "membership")
          vi.mocked(loadMicrosoft365Chat).mockRejectedValue(
            new Microsoft365Error("access_denied", 403),
          );
        if (change === "policy")
          vi.mocked(getMicrosoft365ChatPolicy).mockReturnValue(null);
        if (change === "cas") chat.version = "cas3";
        return answer("Private metadata");
      });
      await expect(turn()).rejects.toBeInstanceOf(Microsoft365Error);
      expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
    },
  );
  it("fails closed when a provider returns a forged quote", async () => {
    history();
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!([
        { id: "read", name: "microsoft365_read", input: { ref: sourceRef } },
      ]);
      return answer("Unsupported", [
        { ref: sourceRef, quote: "A fabricated quotation" },
      ]);
    });
    await expect(turn()).rejects.toMatchObject({
      code: "provider_unavailable",
    });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("copies authoritative ordinary history only on the first protected turn", async () => {
    const ordinaryHistory = [
      { role: "user" as const, content: "Earlier legal question" },
    ];
    const result = await turn(true, { ordinaryHistory });
    expect(result.messages[0]).toEqual({
      ...ordinaryHistory[0],
      sourceRefs: [],
    });
    history();
    const resumed = await turn(true, { ordinaryHistory });
    expect(
      resumed.messages.some(
        (message) => message.content === "Earlier legal question",
      ),
    ).toBe(false);
  });
  it("does not disclose or save private provider errors", async () => {
    vi.mocked(streamChatWithTools).mockRejectedValue(
      new Error("PRIVATE PROVIDER BODY"),
    );
    await expect(turn()).rejects.toMatchObject({
      code: "provider_unavailable",
      message: "provider_unavailable",
    });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("enforces question, context and tool budgets", async () => {
    chat.payload.messages = Array.from({ length: 20 }, () => ({
      role: "user",
      content: "Earlier",
      sourceRefs: [],
    }));
    await expect(turn()).rejects.toMatchObject({ code: "source_too_large" });
    chat.payload.messages = [
      { role: "assistant", content: "x".repeat(100_001), sourceRefs: [] },
    ];
    await expect(turn()).rejects.toMatchObject({ code: "source_too_large" });
    chat.payload.messages = [];
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      await params.runTools!(
        Array.from({ length: 25 }, (_, i) => ({
          id: String(i),
          name: "microsoft365_recent_mail",
          input: {},
        })),
      );
      return answer();
    });
    await expect(turn()).rejects.toMatchObject({ code: "invalid_query" });
    expect(searchMicrosoft365Sources).not.toHaveBeenCalled();
  });
  it("hydrates 20 listed metadata dependencies without downloading bodies", async () => {
    history();
    chat.payload.sources = Array.from({ length: 20 }, (_, i) => ({
      ...chat.payload.sources[0]!,
      ref: `source-${i}`,
    }));
    chat.payload.messages.forEach((message) => {
      message.sourceRefs = chat.payload.sources.map(({ ref }) => ref);
    });
    expect(
      (await getMicrosoft365AssistantChat("owner", "protected", db)).sources,
    ).toHaveLength(20);
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(20);
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
  });
  it("never exposes metadata beyond the conversation source budget", async () => {
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
      items: Array.from({ length: 21 }, (_, i) => ({
        locator: { kind: "mail" as const, id: `mail-${i}` },
        title: `Notice ${i}`,
      })),
      more: true,
      nextOffset: 20,
    });
    vi.mocked(readMicrosoft365SourceMetadata).mockImplementation(
      async (_u, _c, _db, locator) => ({
        ...source,
        locator,
        title: locator.id,
      }),
    );
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const result = await list(params);
      expect(result.items).toHaveLength(20);
      expect(result.sourceLimitReached).toBe(true);
      expect(JSON.stringify(result)).not.toContain("mail-20");
      return answer();
    });
    expect((await turn()).sources).toHaveLength(20);
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(40);
  });
  it("deduplicates file references independently of locator field order", async () => {
    const locator = { kind: "file" as const, driveId: "drive", id: "file" };
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
      items: [{ locator, title: "Contract" }],
      more: false,
      nextOffset: 20,
    });
    vi.mocked(readMicrosoft365SourceMetadata).mockResolvedValue({
      ...source,
      locator: { kind: "file", id: "file", driveId: "drive" },
    });
    vi.mocked(streamChatWithTools).mockImplementation(async (params) => {
      const first = JSON.parse(
        (
          await params.runTools!([
            {
              id: "first",
              name: "microsoft365_search",
              input: { kind: "file", query: "Contract" },
            },
          ])
        )[0]!.content,
      );
      const second = JSON.parse(
        (
          await params.runTools!([
            {
              id: "second",
              name: "microsoft365_search",
              input: { kind: "file", query: "Contract" },
            },
          ])
        )[0]!.content,
      );
      expect(second.items[0].ref).toEqual(first.items[0].ref);
      return answer();
    });
    expect((await turn()).sources).toHaveLength(1);
  });
  it("allows explicitly viewing only a used source with metadata checks before and after", async () => {
    history();
    await expect(
      getMicrosoft365AssistantSource("owner", "protected", "other", db),
    ).rejects.toMatchObject({ code: "invalid_source" });
    expect(readMicrosoft365Source).not.toHaveBeenCalled();
    const result = await getMicrosoft365AssistantSource(
      "owner",
      "protected",
      sourceRef,
      db,
    );
    expect(result.text).toEqual(source.text);
    expect(readMicrosoft365SourceMetadata).toHaveBeenCalledTimes(2);
  });
});
