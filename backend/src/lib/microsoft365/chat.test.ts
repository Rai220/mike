import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  getProtectedMicrosoft365Chat,
  searchProtectedMicrosoft365Chat,
  sendProtectedMicrosoft365Message,
  verifiedMicrosoft365Answer,
  viewProtectedMicrosoft365Source,
} from "./chat";
import {
  loadMicrosoft365Chat,
  saveMicrosoft365Chat,
  getMicrosoft365ChatPolicy,
  type ProtectedChat,
} from "./protected";
import { readMicrosoft365Source, searchMicrosoft365Sources } from "./sources";
import { streamChatWithTools } from "../llm";
vi.mock("./protected", () => ({
  loadMicrosoft365Chat: vi.fn(),
  saveMicrosoft365Chat: vi.fn(),
  getMicrosoft365ChatPolicy: vi.fn(),
}));
vi.mock("./sources", () => ({
  readMicrosoft365Source: vi.fn(),
  searchMicrosoft365Sources: vi.fn(),
}));
vi.mock("../llm", () => ({ streamChatWithTools: vi.fn() }));
const db = {} as Parameters<typeof loadMicrosoft365Chat>[2];
let chat: ProtectedChat;
const original = {
  locator: { kind: "mail" as const, id: "immutable" },
  title: "Synthetic notice",
  text: "The payment deadline is 30 September 2026.",
  version: "v1",
  fetchedAt: "2026-09-07T00:00:00Z",
  webUrl: "https://outlook.office.com/mail/id/immutable",
};
beforeEach(() => {
  vi.resetAllMocks();
  chat = {
    id: "chat",
    connectionId: "conn",
    version: "cas1",
    createdAt: "2026-09-07",
    updatedAt: "2026-09-07",
    expiresAt: "2026-09-08",
    payload: { messages: [], sources: [{ ref: "source", ...original }] },
  };
  vi.mocked(loadMicrosoft365Chat).mockImplementation(async () =>
    structuredClone(chat),
  );
  vi.mocked(saveMicrosoft365Chat).mockImplementation(
    async (_user, meta, payload) => ({ ...meta, payload }) as ProtectedChat,
  );
  vi.mocked(getMicrosoft365ChatPolicy).mockReturnValue({
    model: "gpt-5.5",
    retentionDays: 7,
  });
  vi.mocked(readMicrosoft365Source).mockResolvedValue(original);
  vi.mocked(streamChatWithTools).mockResolvedValue({
    fullText: JSON.stringify({
      answer: "Срок — 30 сентября.",
      citations: [{ ref: "source", quote: original.text }],
    }),
  });
});
function withHistory() {
  chat.payload.messages = [
    {
      role: "assistant",
      content: "protected prior answer",
      sourceRefs: ["source"],
    },
  ];
}
describe("protected corporate chat orchestration", () => {
  it("uses the approved model without generic tools, user keys or raw logs; stores transitive evidence", async () => {
    const result = await sendProtectedMicrosoft365Message("owner", "chat", db, {
      message: "When?",
      sourceRefs: ["source"],
    });
    const params = vi.mocked(streamChatWithTools).mock.calls[0][0];
    expect(params).toMatchObject({ model: "gpt-5.5", sensitive: true });
    for (const field of ["tools", "runTools", "apiKeys", "callbacks"])
      expect(params).not.toHaveProperty(field);
    expect(params.systemPrompt).toContain("untrusted");
    expect(result.messages).toHaveLength(2);
    expect(result.messages[1].content).toContain(original.text);
    expect(
      result.messages.every((message) => message.sourceRefs.includes("source")),
    ).toBe(true);
    expect(readMicrosoft365Source).toHaveBeenCalledTimes(2);
  });
  it("fails closed before model calls and history disclosure on revoked Graph access", async () => {
    withHistory();
    vi.mocked(readMicrosoft365Source).mockRejectedValue(new Error("revoked"));
    await expect(
      getProtectedMicrosoft365Chat("owner", "chat", db),
    ).rejects.toThrow();
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "Continue",
        sourceRefs: [],
      }),
    ).rejects.toThrow();
    expect(streamChatWithTools).not.toHaveBeenCalled();
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("discards generated answer if source access is revoked during model execution", async () => {
    vi.mocked(streamChatWithTools).mockImplementation(async () => {
      vi.mocked(readMicrosoft365Source).mockRejectedValue(new Error("revoked"));
      return {
        fullText: JSON.stringify({
          answer: "sensitive",
          citations: [{ ref: "source", quote: original.text }],
        }),
      };
    });
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["source"],
      }),
    ).rejects.toThrow();
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("blocks changed historical sources and changes between model input and final save", async () => {
    withHistory();
    vi.mocked(readMicrosoft365Source).mockResolvedValue({
      ...original,
      version: "v2",
    });
    await expect(
      getProtectedMicrosoft365Chat("owner", "chat", db),
    ).rejects.toMatchObject({ code: "source_changed" });
    chat.payload.messages = [];
    chat.payload.sources[0].version = "";
    vi.mocked(readMicrosoft365Source)
      .mockResolvedValueOnce(original)
      .mockResolvedValueOnce({ ...original, version: "v2" });
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["source"],
      }),
    ).rejects.toMatchObject({ code: "source_changed" });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("rejects forged refs, excessive dependencies and oversized context", async () => {
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["forged"],
      }),
    ).rejects.toMatchObject({ code: "invalid_source" });
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: Array(6).fill("source"),
      }),
    ).rejects.toMatchObject({ code: "invalid_query" });
    vi.mocked(readMicrosoft365Source).mockResolvedValue({
      ...original,
      text: "x".repeat(100001),
    });
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["source"],
      }),
    ).rejects.toMatchObject({ code: "source_too_large" });
    expect(streamChatWithTools).not.toHaveBeenCalled();
  });
  it.each([
    "not JSON",
    "null",
    JSON.stringify({ answer: "invented", citations: [] }),
    JSON.stringify({
      answer: "invented",
      citations: [{ ref: "forged", quote: original.text }],
    }),
    JSON.stringify({
      answer: "invented",
      citations: [{ ref: "source", quote: "invented exact quote" }],
    }),
  ])("never accepts unverified model quotation: %s", (raw) => {
    expect(() =>
      verifiedMicrosoft365Answer(raw, new Map([["source", original]])),
    ).toThrow();
  });
  it("revalidates the entire transitive history for a citation viewer and excludes unused picker metadata", async () => {
    await expect(
      viewProtectedMicrosoft365Source("owner", "chat", "source", db),
    ).rejects.toMatchObject({ code: "invalid_source" });
    withHistory();
    const result = await viewProtectedMicrosoft365Source(
      "owner",
      "chat",
      "source",
      db,
    );
    expect(result.text).toBe(original.text);
    expect(result.webUrl).toBe(original.webUrl);
    chat.payload.sources.push({
      ...chat.payload.sources[0],
      ref: "unused",
      title: "hidden title",
    });
    expect(
      (await getProtectedMicrosoft365Chat("owner", "chat", db)).sources,
    ).toHaveLength(1);
  });
  it("protects the chat before search and returns only opaque refs, no locators or snippets", async () => {
    vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
      items: [{ locator: original.locator, title: original.title }],
      more: true,
      nextOffset: 20,
    });
    const result = await searchProtectedMicrosoft365Chat("owner", "chat", db, {
      kind: "mail",
      query: "deadline",
    });
    expect(loadMicrosoft365Chat).toHaveBeenCalledBefore(
      searchMicrosoft365Sources,
    );
    expect(saveMicrosoft365Chat).toHaveBeenCalled();
    expect(result).toEqual({
      items: [{ ref: "source", title: original.title, kind: "mail" }],
      more: true,
      nextOffset: 20,
    });
    expect(JSON.stringify(result)).not.toContain("immutable");
    expect(streamChatWithTools).not.toHaveBeenCalled();
  });
  it("never saves when disconnected or CAS changes during model execution", async () => {
    vi.mocked(streamChatWithTools).mockImplementation(async () => {
      chat.version = "concurrent";
      return {
        fullText: JSON.stringify({
          answer: "answer",
          citations: [{ ref: "source", quote: original.text }],
        }),
      };
    });
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["source"],
      }),
    ).rejects.toMatchObject({ code: "busy" });
    expect(saveMicrosoft365Chat).not.toHaveBeenCalled();
  });
  it("blocks corporate model input when policy absent or user cancels", async () => {
    vi.mocked(getMicrosoft365ChatPolicy).mockReturnValue(null);
    await expect(
      sendProtectedMicrosoft365Message("owner", "chat", db, {
        message: "When?",
        sourceRefs: ["source"],
      }),
    ).rejects.toMatchObject({ code: "unavailable" });
    vi.mocked(getMicrosoft365ChatPolicy).mockReturnValue({
      model: "gpt-5.5",
      retentionDays: 7,
    });
    await expect(
      sendProtectedMicrosoft365Message(
        "owner",
        "chat",
        db,
        { message: "When?", sourceRefs: ["source"] },
        AbortSignal.abort(),
      ),
    ).rejects.toBeDefined();
    expect(streamChatWithTools).not.toHaveBeenCalled();
  });
});

it("retains selected source refs across search pages while dropping unused old metadata", async () => {
  chat.payload.sources.push({
    ...chat.payload.sources[0],
    ref: "unselected",
    locator: { kind: "mail", id: "old" },
  });
  vi.mocked(searchMicrosoft365Sources).mockResolvedValue({
    items: [{ locator: { kind: "mail", id: "new" }, title: "New result" }],
    more: false,
    nextOffset: 40,
  });
  await searchProtectedMicrosoft365Chat("owner", "chat", db, {
    kind: "mail",
    query: "deadline",
    offset: 20,
    selectedSourceRefs: ["source"],
  });
  const saved = vi.mocked(saveMicrosoft365Chat).mock.calls[0][2];
  expect(saved.sources.some((source) => source.ref === "source")).toBe(true);
  expect(saved.sources.some((source) => source.ref === "unselected")).toBe(
    false,
  );
});
