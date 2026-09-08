import { afterEach, describe, expect, it, vi } from "vitest";
import { streamAiSdk, type AiSdkAdapterConfig } from "../llm/aiSdk";

// Exercise the installed SDK, including its default logging/telemetry paths.
const raw = vi.hoisted(() => ({
  createRawLlmStreamRecorder: vi.fn(),
  logRawLlmStream: vi.fn(),
}));
vi.mock("../llm/rawStreamLog", () => raw);
afterEach(() => {
  vi.restoreAllMocks();
  vi.clearAllMocks();
  vi.unstubAllGlobals();
});

describe("protected requests with the installed AI SDK", () => {
  it.each([
    ["microsoft365", "microsoft365_read"],
    ["microsoft365-assistant", "microsoft365_read"],
    ["microsoft365-assistant", "read_document"],
    ["microsoft365-assistant", "edgar_read_filing"],
    ["microsoft365-assistant", "read_workflow"],
  ] as const)("executes %s / %s without logging private data or telemetry", async (sensitiveTools, toolName) => {
    const { registerTelemetry } = await import("ai");
    const onStart = vi.fn();
    const onEnd = vi.fn();
    vi.stubGlobal("AI_SDK_TELEMETRY_INTEGRATIONS", []);
    registerTelemetry({ onStart, onEnd });
    const warn = vi.spyOn(console, "warn");
    const error = vi.spyOn(console, "error");
    const log = vi.spyOn(console, "log");
    const runTools = vi.fn(async () => [
      { tool_use_id: "call-1", content: "PRIVATE GRAPH RESULT" },
    ]);
    let iteration = 0;
    const model = {
      specificationVersion: "v4" as const,
      provider: "private-test",
      modelId: "private-model",
      supportedUrls: {},
      doGenerate: vi.fn(),
      doStream: vi.fn(async () => {
        const first = iteration++ === 0;
        return {
          stream: new ReadableStream({
            start(controller) {
              controller.enqueue({
                type: "stream-start",
                warnings: [{ type: "other", message: "PRIVATE WARNING" }],
              });
              if (first) {
                controller.enqueue({ type: "text-start", id: "progress" });
                controller.enqueue({
                  type: "text-delta",
                  id: "progress",
                  delta: "Private intermediate planning",
                });
                controller.enqueue({ type: "text-end", id: "progress" });
                controller.enqueue({
                  type: "tool-call",
                  toolCallId: "call-1",
                  toolName,
                  input: JSON.stringify({ ref: "PRIVATE SOURCE REF" }),
                });
              } else {
                controller.enqueue({ type: "text-start", id: "1" });
                controller.enqueue({
                  type: "text-delta",
                  id: "1",
                  delta: "Private answer",
                });
                controller.enqueue({ type: "text-end", id: "1" });
              }
              controller.enqueue({
                type: "finish",
                finishReason: {
                  unified: first ? "tool-calls" : "stop",
                  raw: first ? "tool_calls" : "stop",
                },
                usage: {
                  inputTokens: {
                    total: 1,
                    noCache: 1,
                    cacheRead: 0,
                    cacheWrite: 0,
                  },
                  outputTokens: { total: 1, text: 1, reasoning: 0 },
                },
              });
              controller.close();
            },
          }),
        };
      }),
    };
    await expect(
      streamAiSdk(
        {
          sensitive: true,
          sensitiveTools,
          ...(sensitiveTools === "microsoft365-assistant"
            ? { apiKeys: { openai: "server-loaded-user-key" } }
            : {}),
          model: "private-model",
          systemPrompt: "Private instructions",
          messages: [{ role: "user", content: "Private question" }],
          maxIterations: 3,
          tools: [
            {
              type: "function",
              function: {
                name: toolName,
                description: "Read corporate source",
                parameters: {
                  type: "object",
                  properties: { ref: { type: "string" } },
                  required: ["ref"],
                  additionalProperties: false,
                },
              },
            },
          ],
          runTools,
        },
        {
          model: model as AiSdkAdapterConfig["model"],
          modelId: "private-model",
          provider: "openai",
          label: "Private",
        },
      ),
    ).resolves.toEqual({ fullText: "Private answer" });
    expect(runTools).toHaveBeenCalledWith([
      {
        id: "call-1",
        name: toolName,
        input: { ref: "PRIVATE SOURCE REF" },
      },
    ]);
    expect(model.doStream).toHaveBeenCalledTimes(2);
    expect(onStart).not.toHaveBeenCalled();
    expect(onEnd).not.toHaveBeenCalled();
    expect(log).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
    expect(error).not.toHaveBeenCalled();
    expect(raw.createRawLlmStreamRecorder).not.toHaveBeenCalled();
    expect(raw.logRawLlmStream).not.toHaveBeenCalled();
  });

  it("streams without provider warnings, raw logs or registered telemetry", async () => {
    const { registerTelemetry } = await import("ai");
    const onStart = vi.fn();
    const onEnd = vi.fn();
    vi.stubGlobal("AI_SDK_TELEMETRY_INTEGRATIONS", []);
    registerTelemetry({ onStart, onEnd });
    const warn = vi.spyOn(console, "warn");
    const error = vi.spyOn(console, "error");
    const model = {
      specificationVersion: "v4" as const,
      provider: "private-test",
      modelId: "private-model",
      supportedUrls: {},
      doGenerate: vi.fn(),
      doStream: vi.fn(async () => ({
        stream: new ReadableStream({
          start(controller) {
            controller.enqueue({
              type: "stream-start",
              warnings: [{ type: "other", message: "PRIVATE WARNING" }],
            });
            controller.enqueue({ type: "text-start", id: "1" });
            controller.enqueue({
              type: "text-delta",
              id: "1",
              delta: "Answer",
            });
            controller.enqueue({ type: "text-end", id: "1" });
            controller.enqueue({
              type: "finish",
              finishReason: { unified: "stop", raw: "stop" },
              usage: {
                inputTokens: {
                  total: 1,
                  noCache: 1,
                  cacheRead: 0,
                  cacheWrite: 0,
                },
                outputTokens: { total: 1, text: 1, reasoning: 0 },
              },
            });
            controller.close();
          },
        }),
      })),
    };
    await expect(
      streamAiSdk(
        {
          sensitive: true,
          model: "private-model",
          systemPrompt: "Private source",
          messages: [{ role: "user", content: "Private question" }],
        },
        {
          model: model as AiSdkAdapterConfig["model"],
          modelId: "private-model",
          provider: "openai",
          label: "Private",
        },
      ),
    ).resolves.toEqual({ fullText: "Answer" });
    expect(model.doStream).toHaveBeenCalledTimes(1);
    expect(onStart).not.toHaveBeenCalled();
    expect(onEnd).not.toHaveBeenCalled();
    expect(warn).not.toHaveBeenCalled();
    expect(error).not.toHaveBeenCalled();
    expect(raw.createRawLlmStreamRecorder).not.toHaveBeenCalled();
    expect(raw.logRawLlmStream).not.toHaveBeenCalled();
  });

  it("does not retry or log provider errors through the SDK default callback", async () => {
    const error = vi.spyOn(console, "error");
    const model = {
      specificationVersion: "v4" as const,
      provider: "private-test",
      modelId: "private-model",
      supportedUrls: {},
      doGenerate: vi.fn(),
      doStream: vi.fn().mockRejectedValue(new Error("PRIVATE REQUEST BODY")),
    };
    await expect(
      streamAiSdk(
        {
          sensitive: true,
          model: "private-model",
          systemPrompt: "Private source",
          messages: [{ role: "user", content: "Private question" }],
        },
        {
          model: model as AiSdkAdapterConfig["model"],
          modelId: "private-model",
          provider: "openai",
          label: "Private",
        },
      ),
    ).rejects.toThrow("Protected model request failed.");
    expect(model.doStream).toHaveBeenCalledTimes(1);
    expect(error).not.toHaveBeenCalled();
    expect(raw.createRawLlmStreamRecorder).not.toHaveBeenCalled();
    expect(raw.logRawLlmStream).not.toHaveBeenCalled();
  });
});
