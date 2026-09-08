import { beforeEach, describe, expect, it, vi } from "vitest";
const fetch = vi.hoisted(() => vi.fn());
vi.mock("@/app/lib/authEvents", () => ({ authenticatedFetch: fetch }));
import {
    createMicrosoft365Chat, deleteMicrosoft365Chat, getMicrosoft365Chat,
    getMicrosoft365ChatSource, listMicrosoft365Chats, searchMicrosoft365ChatSources,
    sendMicrosoft365ChatMessage,
} from "@/app/lib/mikeApi";

beforeEach(() => {
    fetch.mockReset();
    fetch.mockImplementation(async () => new Response(JSON.stringify({}), { headers: { "Content-Type": "application/json" } }));
});

describe("protected Microsoft 365 API", () => {
    it("keeps all requests in the authenticated API path, encodes IDs, and disables caching", async () => {
        const controller = new AbortController();
        const cases = [
            [() => listMicrosoft365Chats(controller.signal), "", undefined],
            [() => createMicrosoft365Chat("connection", controller.signal), "", "POST"],
            [() => getMicrosoft365Chat("chat/1", controller.signal), "/chat%2F1", undefined],
            [() => searchMicrosoft365ChatSources("chat/1", { kind: "file", query: "query", offset: 25, selectedSourceRefs: ["selected-ref"] }, controller.signal), "/chat%2F1/search", "POST"],
            [() => sendMicrosoft365ChatMessage("chat/1", { message: "question", sourceRefs: ["opaque"] }, controller.signal), "/chat%2F1/messages", "POST"],
            [() => getMicrosoft365ChatSource("chat/1", "ref/1", controller.signal), "/chat%2F1/sources/ref%2F1", undefined],
            [() => deleteMicrosoft365Chat("chat/1", controller.signal), "/chat%2F1", "DELETE"],
        ] as const;
        for (const [call, suffix, method] of cases) {
            await call();
            const [url, init] = fetch.mock.calls.at(-1)!;
            expect(url).toBe(`/api/integrations/microsoft365/chats${suffix}`);
            expect(init).toMatchObject({ cache: "no-store", signal: controller.signal });
            expect(init.method).toBe(method);
        }
        expect(JSON.parse(fetch.mock.calls[3][1].body)).toEqual({ kind: "file", query: "query", offset: 25, selectedSourceRefs: ["selected-ref"] });
        expect(JSON.parse(fetch.mock.calls[4][1].body)).toEqual({ message: "question", sourceRefs: ["opaque"] });
    });

    it("preserves the access error status for intentional user messages", async () => {
        fetch.mockResolvedValue(new Response(JSON.stringify({ error: "Access revoked", code: "source_access_revoked" }), { status: 403, headers: { "Content-Type": "application/json" } }));
        await expect(getMicrosoft365Chat("chat")).rejects.toMatchObject({ status: 403 });
    });
});
