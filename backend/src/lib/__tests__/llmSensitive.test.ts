import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { StreamChatParams } from "../llm/types";

const mocks = vi.hoisted(() => ({
  streamText: vi.fn(),
  wrapLanguageModel: vi.fn(({ model }) => model),
  createRawLlmStreamRecorder: vi.fn(),
  logRawLlmStream: vi.fn(),
  createOpenAI: vi.fn(),
}));
vi.mock("ai", () => ({
  streamText: mocks.streamText,
  wrapLanguageModel: mocks.wrapLanguageModel,
  stepCountIs: vi.fn(),
  jsonSchema: vi.fn(),
  tool: vi.fn(),
}));
vi.mock("../llm/rawStreamLog", () => ({
  createRawLlmStreamRecorder: mocks.createRawLlmStreamRecorder,
  logRawLlmStream: mocks.logRawLlmStream,
}));
vi.mock("@ai-sdk/openai", () => ({ createOpenAI: mocks.createOpenAI }));
import {
  aiSdkFetch,
  streamAiSdk,
  assertSensitiveChatParams,
  type AiSdkAdapterConfig,
} from "../llm/aiSdk";
import { streamWithProvider } from "../llm/providers";
import { EDGAR_TOOL_NAMES } from "../chat/tools/edgarTools";

const model = {
  specificationVersion: "v4",
  provider: "openai",
  modelId: "gpt-5.5",
};
const config: AiSdkAdapterConfig = {
  provider: "openai",
  label: "OpenAI",
  model: model as AiSdkAdapterConfig["model"],
  modelId: "gpt-5.5",
};
const params: StreamChatParams = {
  sensitive: true,
  model: "gpt-5.5",
  systemPrompt: "Private sources",
  messages: [{ role: "user", content: "Private question" }],
};
function assistantParams(names = ["microsoft365_read"]): StreamChatParams {
  return {
    ...params,
    sensitiveTools: "microsoft365-assistant",
    apiKeys: { openai: "server-loaded-user-key" },
    tools: names.map((name) => ({
      type: "function",
      function: {
        name,
        description: "Read source",
        parameters: { type: "object" },
      },
    })),
    runTools: vi.fn(),
  };
}
function emit(parts: unknown[]) {
  mocks.streamText.mockReturnValue({
    stream: (async function* () {
      yield* parts;
    })(),
  });
}
function expectNoRecording() {
  expect(mocks.createRawLlmStreamRecorder).not.toHaveBeenCalled();
  expect(mocks.logRawLlmStream).not.toHaveBeenCalled();
}

beforeEach(() => {
  vi.clearAllMocks();
  vi.stubEnv("OPENAI_API_KEY", "server-key");
  vi.stubEnv("LOG_RAW_LLM_STREAM", "true");
  vi.stubEnv("RAW_LLM_STREAM_LOG_DIR", "/not-used-for-protected-requests");
  mocks.createOpenAI.mockReturnValue({ responses: vi.fn(() => model) });
});
afterEach(() => {
  vi.unstubAllEnvs();
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("protected LLM requests", () => {
  it("never records raw events even if the SDK emits them and logging is enabled", async () => {
    emit([
      { type: "raw", rawValue: { private: "body" } },
      { type: "text-delta", text: "Answer" },
    ]);
    const onContentDelta = vi.fn();
    await expect(
      streamAiSdk({ ...params, callbacks: { onContentDelta } }, config),
    ).resolves.toEqual({ fullText: "Answer" });
    expect(onContentDelta).toHaveBeenCalledWith("Answer");
    expectNoRecording();
    expect(mocks.streamText).toHaveBeenCalledWith(
      expect.objectContaining({
        include: { rawChunks: false },
        maxRetries: 0,
        telemetry: {
          isEnabled: false,
          recordInputs: false,
          recordOutputs: false,
        },
        tools: undefined,
      }),
    );
  });

  it.each(["error", "tool-error"])(
    "never records or exposes provider %s bodies",
    async (type) => {
      emit([
        { type: "raw", rawValue: "private body" },
        { type, error: new Error("private provider body") },
      ]);
      await expect(streamAiSdk(params, config)).rejects.toThrow(
        "Protected model request failed.",
      );
      expectNoRecording();
      const consoleError = vi.spyOn(console, "error");
      const options = mocks.streamText.mock.calls[0]![0];
      options.onError({ error: new Error("private provider body") });
      expect(consoleError).not.toHaveBeenCalled();
    },
  );

  it("never records synchronous SDK failures", async () => {
    mocks.streamText.mockImplementationOnce(() => {
      throw new Error("private request body");
    });
    await expect(streamAiSdk(params, config)).rejects.toThrow(
      "Protected model request failed.",
    );
    expectNoRecording();
  });

  it.each([{ apiKeys: {} }, { tools: [] }, { runTools: vi.fn() }])(
    "rejects unsafe overrides before constructing the provider: %j",
    async (override) => {
      await expect(
        streamWithProvider({ ...params, ...override }),
      ).rejects.toThrow("does not allow user API keys or tools");
      await expect(
        streamAiSdk({ ...params, ...override }, config),
      ).rejects.toThrow("does not allow user API keys or tools");
      expect(mocks.createOpenAI).not.toHaveBeenCalled();
      expect(mocks.streamText).not.toHaveBeenCalled();
      expectNoRecording();
    },
  );

  it("never retries a reasoning validation error or constructs a fallback adapter", async () => {
    emit([
      {
        type: "error",
        error: new Error("Supported values are: 'none', 'medium'."),
      },
    ]);
    await expect(
      streamWithProvider({ ...params, reasoning: "low" }),
    ).rejects.toThrow("Protected model request failed.");
    expect(mocks.streamText).toHaveBeenCalledTimes(1);
    expect(mocks.createOpenAI).toHaveBeenCalledTimes(1);
    expect(mocks.createOpenAI).toHaveBeenCalledWith(
      expect.objectContaining({ apiKey: "server-key" }),
    );
    expectNoRecording();
  });

  it("rejects unexpected tool calls without invoking callbacks", async () => {
    const onToolCallStart = vi.fn();
    emit([
      {
        type: "tool-call",
        toolName: "exfiltrate",
        toolCallId: "1",
        input: { secret: "body" },
      },
    ]);
    await expect(
      streamAiSdk({ ...params, callbacks: { onToolCallStart } }, config),
    ).rejects.toThrow("Protected model request failed.");
    expect(onToolCallStart).not.toHaveBeenCalled();
    expectNoRecording();
  });

  it("allows only explicitly opted-in Graph tool schemas and rejects callbacks or user keys", () => {
    const trusted: StreamChatParams = {
      ...params,
      sensitiveTools: "microsoft365",
      tools: [
        {
          type: "function",
          function: {
            name: "microsoft365_read",
            description: "Read source",
            parameters: { type: "object" },
          },
        },
      ],
      runTools: vi.fn(),
    };
    expect(() => assertSensitiveChatParams(trusted)).not.toThrow();
    for (const override of [
      { callbacks: { onToolCallStart: vi.fn() } },
      { apiKeys: {} },
      { sensitive: false },
      { runTools: undefined },
      { tools: [] },
      { tools: [...trusted.tools!, ...trusted.tools!] },
      {
        tools: [
          {
            type: "function" as const,
            function: {
              name: "send_email",
              description: "Send",
              parameters: {},
            },
          },
        ],
      },
    ])
      expect(() =>
        assertSensitiveChatParams({ ...trusted, ...override }),
      ).toThrow();
  });

  it("allows main-assistant read tools together with server-loaded user model keys", async () => {
    const trusted = assistantParams([
      "microsoft365_search", "microsoft365_recent_mail", "microsoft365_read",
      "read_document", "find_in_document", "list_documents", "fetch_documents",
      "list_workflows", "read_workflow", ...Object.values(EDGAR_TOOL_NAMES),
    ]);
    expect(() => assertSensitiveChatParams(trusted)).not.toThrow();
    emit([
      { type: "start-step" },
      { type: "text-delta", text: "Private planning" },
      { type: "tool-call", toolName: "edgar_read_filing", toolCallId: "1", input: {} },
      { type: "start-step" },
      { type: "raw", rawValue: "private content" },
      { type: "text-delta", text: "Final answer" },
    ]);
    await expect(streamWithProvider(trusted)).resolves.toEqual({ fullText: "Final answer" });
    expect(mocks.createOpenAI).toHaveBeenCalledWith(
      expect.objectContaining({ apiKey: "server-loaded-user-key" }),
    );
    expect(mocks.streamText).toHaveBeenCalledWith(expect.objectContaining({
      include: { rawChunks: false }, maxRetries: 0,
      telemetry: { isEnabled: false, recordInputs: false, recordOutputs: false },
    }));
    expectNoRecording();
  });

  it.each([
    "exfiltrate", "mcp_search", "send_email", "generate_docx", "edit_document",
    "replicate_document", "generate_excel", "generate_ppt", "ask_inputs",
    "courtlistener_read_case",
  ])("rejects main-assistant tool %s before constructing a provider", async (name) => {
    await expect(streamWithProvider(assistantParams([name]))).rejects.toThrow(
      "does not allow user API keys or tools",
    );
    expect(mocks.createOpenAI).not.toHaveBeenCalled();
    expect(mocks.streamText).not.toHaveBeenCalled();
    expectNoRecording();
  });

  it("rejects malformed, duplicated or unbuffered main-assistant capabilities", () => {
    const trusted = assistantParams();
    for (const override of [
      { callbacks: {} }, { sensitive: false }, { runTools: undefined },
      { tools: [] }, { tools: [...trusted.tools!, ...trusted.tools!] },
      { tools: [null] }, { tools: [{}] },
      { tools: [{ type: "function", function: null }] },
      { tools: [{ ...trusted.tools![0], type: "custom" }] },
      ...[null, [], {}, "object"].map((parameters) => ({
        tools: [{ type: "function", function: { ...trusted.tools![0]!.function, parameters } }],
      })),
      { tools: [{ type: "function", function: { ...trusted.tools![0]!.function, description: null } }] },
    ]) {
      expect(() => assertSensitiveChatParams({ ...trusted, ...override } as StreamChatParams))
        .toThrow("does not allow user API keys or tools");
    }
  });

  it.each(["microsoft365_read", "edgar_read_filing", "exfiltrate"])(
    "rejects undeclared main-assistant tool calls even if allowlisted: %s", async (toolName) => {
      emit([{ type: "tool-call", toolName, toolCallId: "1", input: {} }]);
      const trusted = assistantParams(["read_document"]);
      await expect(streamAiSdk(trusted, config)).rejects.toThrow("Protected model request failed.");
      expect(trusted.runTools).not.toHaveBeenCalled();
      expectNoRecording();
    },
  );

  it("does not retry or fall back for main-assistant provider errors", async () => {
    emit([{ type: "error", error: new Error("Supported values are: 'none', 'medium'.") }]);
    await expect(streamWithProvider({ ...assistantParams(), reasoning: "low" }))
      .rejects.toThrow("Protected model request failed.");
    expect(mocks.createOpenAI).toHaveBeenCalledTimes(1);
    expect(mocks.streamText).toHaveBeenCalledTimes(1);
    expectNoRecording();
  });

  it("removes provider warnings before the SDK's global warning logger sees them", async () => {
    emit([]);
    await streamAiSdk(params, config);
    const { middleware } = mocks.wrapLanguageModel.mock
      .calls[0]![0] as unknown as {
      middleware: {
        wrapStream: (args: {
          doStream: () => Promise<{ stream: ReadableStream }>;
        }) => Promise<{ stream: ReadableStream }>;
      };
    };
    const result = await middleware.wrapStream({
      doStream: async () => ({
        stream: new ReadableStream({
          start(controller) {
            controller.enqueue({
              type: "stream-start",
              warnings: [{ type: "other", message: "private body" }],
            });
            controller.enqueue({ type: "text-delta", delta: "Answer" });
            controller.close();
          },
        }),
      }),
    });
    const reader = result.stream.getReader();
    expect((await reader.read()).value).toEqual({
      type: "stream-start",
      warnings: [],
    });
    expect((await reader.read()).value).toEqual({
      type: "text-delta",
      delta: "Answer",
    });
  });

  it("preserves abort identity without exposing abort reasons", async () => {
    emit([{ type: "abort", reason: "private reason" }]);
    await expect(streamAiSdk(params, config)).rejects.toMatchObject({
      name: "AbortError",
      message: "Protected model request failed.",
    });
    expectNoRecording();
  });

  it("preserves ordinary chat recording behavior", async () => {
    const record = vi.fn();
    const flush = vi.fn();
    mocks.createRawLlmStreamRecorder.mockReturnValueOnce({ record, flush });
    emit([{ type: "raw", rawValue: "ordinary" }]);
    await streamAiSdk({ ...params, sensitive: false }, config);
    expect(mocks.createRawLlmStreamRecorder).toHaveBeenCalledTimes(1);
    expect(mocks.logRawLlmStream).toHaveBeenCalledTimes(1);
    expect(record).toHaveBeenCalledTimes(1);
    expect(flush).toHaveBeenCalledWith("completed");
  });

  it("fetch stream normalization does not log request or response bodies", async () => {
    vi.stubGlobal(
      "fetch",
      vi.fn().mockResolvedValue(
        new Response('data: {"text":"private response"}\n\n', {
          headers: { "content-type": "text/event-stream" },
        }),
      ),
    );
    const log = vi.spyOn(console, "log");
    const warn = vi.spyOn(console, "warn");
    const error = vi.spyOn(console, "error");
    const response = await aiSdkFetch("https://example.test/stream", {
      method: "POST",
      body: "private request",
    });
    await response.text();
    expect(log).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
    expect(error).not.toHaveBeenCalled();
    expectNoRecording();
  });
});
