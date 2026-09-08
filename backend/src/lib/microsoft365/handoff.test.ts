import { beforeEach, describe, expect, it, vi } from "vitest";

const {
  getMicrosoft365Config,
  getMicrosoft365Status,
  streamChatWithTools,
  buildUserMcpTools,
} = vi.hoisted(() => ({
  getMicrosoft365Status: vi.fn(),
  getMicrosoft365Config: vi.fn(() => null as { frontendOrigin: string } | null),
  streamChatWithTools: vi.fn(async () => ({ fullText: "" })),
  buildUserMcpTools: vi.fn(async () => []),
}));
vi.mock("./index", () => ({ getMicrosoft365Config, getMicrosoft365Status }));
vi.mock("../llm", async () => ({
  ...(await vi.importActual<Record<string, unknown>>("../llm/models")),
  streamChatWithTools,
}));
vi.mock("../mcpConnectors", () => ({ buildUserMcpTools }));

import {
  getMicrosoft365ChatHandoff,
  isMicrosoft365ReadRequest,
} from "./handoff";
import { runLLMStream, AssistantStreamAbortError } from "../chat/streaming";

const query = "Какие письма у меня на почте?";
const db = {} as never;
const messages = (content: string) => [{ role: "user", content }];
const base = (content = query) => ({
  model: "gemini-3-flash-preview",
  apiMessages: messages(content),
  docStore: new Map(),
  docIndex: {},
  userId: "u1",
  db,
  write: vi.fn(),
});

beforeEach(() => {
  vi.clearAllMocks();
  getMicrosoft365Config.mockReturnValue(null);
  getMicrosoft365Status.mockResolvedValue({
    available: true,
    connection: { status: "connected" },
  });
});

describe("ordinary chat Microsoft 365 handoff", () => {
  it.each([
    query,
    "Прочитай мои письма",
    "Покажи непрочитанные письма в моей почте",
    "what emails are in my inbox?",
    "List my latest emails",
    "Find the contract in my OneDrive",
    "Найди договор на корпоративном диске",
  ])("recognizes a source read request: %s", (text) => {
    expect(isMicrosoft365ReadRequest(text)).toBe(true);
  });

  it.each([
    "Составь письмо клиенту",
    "Напиши письмо: какие письма у меня на почте?",
    "Объясни формат письма",
    "Draft an email asking what is in my inbox",
    "Translate my email",
    "Покажи документы этого проекта",
    "Какие бывают письма?",
    "What is the legal format of an email?",
    "Summarize the email attached to this chat",
  ])("does not hijack ordinary work: %s", async (text) => {
    expect(
      await getMicrosoft365ChatHandoff(messages(text), "u1", db),
    ).toBeNull();
    expect(getMicrosoft365Status).not.toHaveBeenCalled();
  });

  it("checks only the latest user turn", async () => {
    expect(
      await getMicrosoft365ChatHandoff(
        [
          ...messages(query),
          { role: "assistant", content: query },
          ...messages("Составь письмо"),
        ],
        "u1",
        db,
      ),
    ).toBeNull();
    expect(getMicrosoft365Status).not.toHaveBeenCalled();
  });

  it.each([
    { available: false, connection: { status: "connected" } },
    { available: true, connection: null },
    { available: true, connection: { status: "reconnect_required" } },
    { available: true, connection: { status: "pending" } },
  ])(
    "does not assert access when connection is unavailable: %j",
    async (status) => {
      getMicrosoft365Status.mockResolvedValue(status);
      expect(
        await getMicrosoft365ChatHandoff(messages(query), "u1", db),
      ).toBeNull();
    },
  );

  it("falls through safely if status lookup fails", async () => {
    getMicrosoft365Status.mockRejectedValue(new Error("private DB details"));
    expect(
      await getMicrosoft365ChatHandoff(messages(query), "u1", db),
    ).toBeNull();
  });

  it("returns the standard persisted SSE response without calling any LLM or MCP", async () => {
    const params = base();
    const result = await runLLMStream(params);
    expect(getMicrosoft365Status).toHaveBeenCalledWith("u1", db);
    expect(result.fullText).toContain(
      "Включите Microsoft 365 рядом с полем ввода",
    );
    expect(result.events).toEqual([{ type: "content", text: result.fullText }]);
    expect(result.citations).toEqual([]);
    expect(params.write.mock.calls.map(([chunk]) => chunk)).toEqual([
      `data: ${JSON.stringify(result.events[0])}\n\n`,
      'data: {"type":"citations","status":"final","citations":[]}\n\n',
      "data: [DONE]\n\n",
    ]);
    expect(streamChatWithTools).not.toHaveBeenCalled();
    expect(buildUserMcpTools).not.toHaveBeenCalled();
  });

  it("keeps the instruction in the same conversation regardless of origin", async () => {
    getMicrosoft365Config.mockReturnValue({
      frontendOrigin: "https://mike.example",
    });
    const result = await runLLMStream(base());
    expect(result.fullText).toContain(
      "прямо в этом чате",
    );
  });

  it("respects caller-owned DONE and English locale", async () => {
    const params = { ...base("What emails are in my inbox?"), emitDone: false };
    const result = await runLLMStream(params);
    expect(result.fullText).toContain("Enable Microsoft 365 beside the message composer");
    expect(params.write).not.toHaveBeenCalledWith("data: [DONE]\n\n");
  });

  it("does not emit an aborted handoff", async () => {
    const controller = new AbortController();
    controller.abort();
    const params = { ...base(), signal: controller.signal };
    await expect(runLLMStream(params)).rejects.toBeInstanceOf(
      AssistantStreamAbortError,
    );
    expect(params.write).not.toHaveBeenCalled();
  });

  it("keeps legal drafting on the existing LLM loop with routing awareness", async () => {
    await runLLMStream(base("Draft an email for my client"));
    expect(getMicrosoft365Status).not.toHaveBeenCalled();
    expect(streamChatWithTools).toHaveBeenCalledOnce();
    expect(streamChatWithTools.mock.calls[0]?.[0]).toMatchObject({
      systemPrompt: expect.stringContaining("Microsoft 365 toggle"),
    });
  });
});
