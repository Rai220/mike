import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const api = vi.hoisted(() => ({
    createMicrosoft365Chat: vi.fn(), deleteMicrosoft365Chat: vi.fn(),
    getMicrosoft365Chat: vi.fn(), getMicrosoft365ChatSource: vi.fn(),
    getMicrosoft365Connection: vi.fn(), listMicrosoft365Chats: vi.fn(),
    searchMicrosoft365ChatSources: vi.fn(), sendMicrosoft365ChatMessage: vi.fn(),
}));
vi.mock("@/app/lib/mikeApi", async (original) => ({ ...await original<typeof import("@/app/lib/mikeApi")>(), ...api }));
vi.mock("@/app/components/shared/PageHeader", () => ({ PageHeader: ({ children }: { children: React.ReactNode }) => <h1>{children}</h1> }));
import { Microsoft365ChatPage } from "./Microsoft365ChatPage";
import { MikeApiError } from "@/app/lib/mikeApi";

const source = { ref: "opaque-ref", title: "Quarterly forecast", kind: "mail" as const, version: "v1", fetchedAt: "2026-09-07T07:00:00Z" };
const chat = { chatId: "chat-1", expiresAt: "2099-09-08T07:00:00Z", sources: [], messages: [] };
const populated = { ...chat, sources: [source], messages: [{ role: "assistant", content: "Private answer https://evil.example/link", sourceRefs: [source.ref] }] };
const listing = { available: true, policy: { model: "approved-model", retentionDays: 7 }, chats: [{ id: "chat-1", createdAt: "2026-09-07T07:00:00Z", expiresAt: chat.expiresAt }] };

async function open() {
    render(<Microsoft365ChatPage />);
    fireEvent.click(await screen.findByRole("button", { name: /Chat 1/ }));
    await screen.findByRole("region", { name: "Choose sources" });
}
async function search() {
    fireEvent.change(screen.getByRole("textbox", { name: "Search Microsoft 365" }), { target: { value: "forecast" } });
    fireEvent.click(screen.getByRole("button", { name: "Search" }));
    await screen.findByRole("checkbox", { name: source.title });
}

describe("Microsoft365ChatPage", () => {
    beforeEach(() => {
        vi.resetAllMocks();
        api.listMicrosoft365Chats.mockResolvedValue(listing);
        api.getMicrosoft365Connection.mockResolvedValue({ available: true, connection: { id: "connection-1", status: "connected" } });
        api.getMicrosoft365Chat.mockResolvedValue(chat);
        api.searchMicrosoft365ChatSources.mockResolvedValue({ items: [source], more: false, nextOffset: 10 });
        api.sendMicrosoft365ChatMessage.mockResolvedValue(populated);
    });
    afterEach(() => { cleanup(); vi.restoreAllMocks(); });

    it("loads recent inbox metadata for a new chat without selecting or sending it", async () => {
        api.createMicrosoft365Chat.mockResolvedValue({ chatId: "chat-1" });
        render(<Microsoft365ChatPage />);
        await waitFor(() => expect(screen.getByRole("button", { name: "New chat" })).toBeEnabled());
        fireEvent.click(screen.getByRole("button", { name: "New chat" }));
        const checkbox = await screen.findByRole("checkbox", { name: source.title });
        expect(checkbox).not.toBeChecked();
        expect(api.searchMicrosoft365ChatSources).toHaveBeenCalledWith("chat-1", { kind: "mail", mode: "recent", query: "", offset: 0, selectedSourceRefs: [] }, expect.any(AbortSignal));
        expect(api.sendMicrosoft365ChatMessage).not.toHaveBeenCalled();
        expect(api.getMicrosoft365ChatSource).not.toHaveBeenCalled();
    });

    it("pages recent mail with selected refs and switches back to search when edited", async () => {
        api.searchMicrosoft365ChatSources
            .mockResolvedValueOnce({ items: [source], more: true, nextOffset: 20 })
            .mockResolvedValueOnce({ items: [{ ...source, ref: "next-ref", title: "Next inbox message" }], more: false, nextOffset: 40 })
            .mockResolvedValueOnce({ items: [], more: false, nextOffset: 0 });
        await open();
        fireEvent.click(screen.getByRole("button", { name: "Recent mail" }));
        fireEvent.click(await screen.findByRole("checkbox", { name: source.title }));
        fireEvent.click(screen.getByRole("button", { name: "Next results" }));
        await screen.findByRole("checkbox", { name: "Next inbox message" });
        expect(api.searchMicrosoft365ChatSources).toHaveBeenLastCalledWith("chat-1", { kind: "mail", mode: "recent", query: "", offset: 20, selectedSourceRefs: [source.ref] }, expect.any(AbortSignal));
        expect(screen.getByRole("button", { name: `Remove ${source.title}` })).toBeInTheDocument();
        fireEvent.change(screen.getByRole("textbox", { name: "Search Microsoft 365" }), { target: { value: "contract" } });
        fireEvent.click(screen.getByRole("button", { name: "Search" }));
        await waitFor(() => expect(api.searchMicrosoft365ChatSources).toHaveBeenLastCalledWith("chat-1", { kind: "mail", mode: "search", query: "contract", offset: 0, selectedSourceRefs: [source.ref] }, expect.any(AbortSignal)));
    });

    it("clears recent-mail mode when switching to files", async () => {
        await open();
        fireEvent.click(screen.getByRole("button", { name: "Recent mail" }));
        await screen.findByRole("checkbox", { name: source.title });
        fireEvent.click(screen.getByRole("button", { name: "Files" }));
        expect(screen.queryByRole("button", { name: "Recent mail" })).toBeNull();
        expect(screen.queryByText(/Recent inbox messages/)).toBeNull();
        await search();
        expect(api.searchMicrosoft365ChatSources).toHaveBeenLastCalledWith("chat-1", { kind: "file", mode: "search", query: "forecast", offset: 0, selectedSourceRefs: [] }, expect.any(AbortSignal));
    });

    it("requires explicit source selection and submits only opaque refs", async () => {
        await open(); await search();
        fireEvent.change(screen.getByRole("textbox", { name: "Your question" }), { target: { value: "Summarize" } });
        expect(screen.getByRole("button", { name: "Send" })).toBeDisabled();
        fireEvent.click(screen.getByRole("checkbox", { name: source.title }));
        fireEvent.click(screen.getByRole("button", { name: "Send" }));
        await screen.findByText(populated.messages[0].content);
        expect(api.sendMicrosoft365ChatMessage).toHaveBeenCalledWith("chat-1", { message: "Summarize", sourceRefs: [source.ref] }, expect.any(AbortSignal));
        expect(screen.queryByRole("link", { name: /evil/ })).toBeNull();
        expect(screen.getByRole("button", { name: `View source: ${source.title}` })).toBeInTheDocument();
    });

    it("enforces five selections and uses server pagination offset", async () => {
        const items = Array.from({ length: 6 }, (_, i) => ({ ...source, ref: `ref-${i}`, title: i ? `File ${i}` : source.title }));
        api.searchMicrosoft365ChatSources.mockResolvedValue({ items, more: true, nextOffset: 25 });
        await open(); await search();
        screen.getAllByRole("checkbox").slice(0, 5).forEach(box => fireEvent.click(box));
        expect(screen.getAllByRole("checkbox")[5]).toBeDisabled();
        fireEvent.click(screen.getByRole("button", { name: "Next results" }));
        await waitFor(() => expect(api.searchMicrosoft365ChatSources).toHaveBeenLastCalledWith("chat-1", { kind: "mail", mode: "search", query: "forecast", offset: 25, selectedSourceRefs: ["ref-0", "ref-1", "ref-2", "ref-3", "ref-4"] }, expect.any(AbortSignal)));
    });

    it("replaces search pages while preserving explicit selections across pages and queries", async () => {
        api.searchMicrosoft365ChatSources
            .mockResolvedValueOnce({ items: [source], more: true, nextOffset: 25 })
            .mockResolvedValueOnce({ items: [{ ...source, ref: "next-ref", title: "Next page file" }], more: false, nextOffset: 50 })
            .mockResolvedValueOnce({ items: [], more: false, nextOffset: 0 });
        await open(); await search();
        fireEvent.click(screen.getByRole("checkbox", { name: source.title }));
        fireEvent.click(screen.getByRole("button", { name: "Next results" }));
        await screen.findByRole("checkbox", { name: "Next page file" });
        expect(screen.queryByRole("checkbox", { name: source.title })).toBeNull();
        expect(screen.getByRole("button", { name: `Remove ${source.title}` })).toBeInTheDocument();
        fireEvent.change(screen.getByRole("textbox", { name: "Search Microsoft 365" }), { target: { value: "another query" } });
        fireEvent.click(screen.getByRole("button", { name: "Search" }));
        await waitFor(() => expect(api.searchMicrosoft365ChatSources).toHaveBeenLastCalledWith("chat-1", { kind: "mail", mode: "search", query: "another query", offset: 0, selectedSourceRefs: [source.ref] }, expect.any(AbortSignal)));
    });

    it("clears all prior output and results when a source access check fails", async () => {
        api.getMicrosoft365Chat.mockResolvedValue(populated);
        api.getMicrosoft365ChatSource.mockRejectedValue(new MikeApiError({ status: 403, code: "source_access_revoked", message: "Source access was revoked." }));
        await open(); await search();
        fireEvent.click(screen.getByRole("button", { name: `View source: ${source.title}` }));
        expect(await screen.findByRole("alert")).toHaveTextContent("Source access was revoked.");
        expect(screen.queryByText(populated.messages[0].content)).toBeNull();
        expect(screen.queryByRole("checkbox")).toBeNull();
        expect(screen.queryByRole("region", { name: "Source viewer" })).toBeNull();
    });

    it("loads exact source text through authorized viewer and refuses unsafe original URLs", async () => {
        api.getMicrosoft365Chat.mockResolvedValue(populated);
        api.getMicrosoft365ChatSource.mockResolvedValue({ ...source, text: "Exact quotation from original", webUrl: "javascript:alert(1)" });
        await open();
        fireEvent.click(screen.getByRole("button", { name: `View source: ${source.title}` }));
        expect(await screen.findByText("Exact quotation from original")).toBeInTheDocument();
        expect(api.getMicrosoft365ChatSource).toHaveBeenCalledWith("chat-1", source.ref, expect.any(AbortSignal));
        expect(screen.queryByRole("link", { name: "Open original in Microsoft 365" })).toBeNull();
    });

    it("does not display a late cancelled response", async () => {
        let resolve!: (value: typeof populated) => void;
        api.sendMicrosoft365ChatMessage.mockReturnValue(new Promise(done => { resolve = done; }));
        await open(); await search();
        fireEvent.click(screen.getByRole("checkbox", { name: source.title }));
        fireEvent.change(screen.getByRole("textbox", { name: "Your question" }), { target: { value: "Summarize" } });
        fireEvent.click(screen.getByRole("button", { name: "Send" }));
        fireEvent.click(screen.getByRole("button", { name: "Cancel request" }));
        resolve(populated);
        await waitFor(() => expect(screen.queryByText(populated.messages[0].content)).toBeNull());
        expect(api.sendMicrosoft365ChatMessage.mock.calls[0][2].aborted).toBe(true);
    });

    it("hides corporate content when the tab is hidden", async () => {
        api.getMicrosoft365Chat.mockResolvedValue(populated);
        await open();
        vi.spyOn(document, "visibilityState", "get").mockReturnValue("hidden");
        fireEvent(document, new Event("visibilitychange"));
        expect(screen.queryByText(populated.messages[0].content)).toBeNull();
    });

    it("explains unavailable policy and disables creation", async () => {
        api.listMicrosoft365Chats.mockResolvedValue({ available: false, policy: null, chats: [] });
        render(<Microsoft365ChatPage />);
        expect(await screen.findByText(/administrator must enable/)).toBeInTheDocument();
        expect(screen.getByRole("button", { name: "New chat" })).toBeDisabled();
    });
});
