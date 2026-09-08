import { createRef } from "react";
import { act, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { useUserProfile } from "@/app/contexts/UserProfileContext";
import { listWorkflows, uploadProjectDocument, uploadStandaloneDocuments } from "@/app/lib/mikeApi";
import type { Document, Workflow } from "../shared/types";
import { ChatInput, type ChatInputHandle } from "./ChatInput";

vi.mock("@/app/lib/mikeApi", () => ({
    listWorkflows: vi.fn(async () => []),
    uploadProjectDocument: vi.fn(),
    uploadStandaloneDocument: vi.fn(),
    uploadStandaloneDocuments: vi.fn(async () => []),
}));

vi.mock("@/app/contexts/UserProfileContext", () => ({
    useUserProfile: vi.fn(),
}));

vi.mock("@/app/lib/modelAvailability", () => ({
    getModelProvider: vi.fn(),
    isModelAvailable: vi.fn(() => true),
}));

vi.mock("./ModelToggle", async (importOriginal) => ({
    ...(await importOriginal<typeof import("./ModelToggle")>()),
    ModelToggle: () => (
        <>
            <button aria-label="Choose model" />
            <button aria-label="Choose reasoning" />
        </>
    ),
}));

vi.mock("./AddDocButton", () => ({
    AddDocButton: () => <button aria-label="Add documents" />,
}));
vi.mock("./UploadOverlay", () => ({ UploadOverlay: () => null }));
vi.mock("../shared/FileTypeIcon", () => ({ FileTypeIcon: () => null }));
vi.mock("../modals/AddDocumentsModal", () => ({
    AddDocumentsModal: () => null,
}));
vi.mock("./AssistantWorkflowModal", () => ({
    AssistantWorkflowModal: () => null,
}));
vi.mock("../popups/ApiKeyMissingPopup", () => ({
    ApiKeyMissingPopup: () => null,
}));

class ResizeObserverMock {
    observe() {}
    disconnect() {}
}

function mockProfile() {
    vi.mocked(useUserProfile).mockReturnValue({
        profile: {
            openRouterModels: [],
            vercelModels: [],
            openCodeGoModels: [],
            apiKeys: {},
        },
        loading: false,
        apiKeysDegraded: false,
    } as unknown as ReturnType<typeof useUserProfile>);
}

function renderInput(canSend: boolean, onSubmit = vi.fn()) {
    render(
        <ChatInput
            onSubmit={onSubmit}
            onCancel={vi.fn()}
            isLoading={false}
            canSend={canSend}
            projectId="p1"
        />,
    );
    return onSubmit;
}

describe("ChatInput canSend gating", () => {
    beforeEach(() => {
        vi.clearAllMocks();
        window.localStorage.clear();
        vi.stubGlobal("ResizeObserver", ResizeObserverMock);
        mockProfile();
    });

    it("renders a read-only composer when canSend is false", () => {
        renderInput(false);

        const textarea = screen.getByPlaceholderText(
            "Viewing only — sending needs edit access",
        );
        expect(textarea).toBeDisabled();
        expect(
            screen.getByRole("button", { name: "Send message" }),
        ).toBeDisabled();
        expect(
            screen.queryByRole("button", { name: "Add documents" }),
        ).toBeNull();
        expect(
            screen.queryByRole("button", { name: "Open workflows" }),
        ).toBeNull();
    });

    it("does not submit on Enter when canSend is false", () => {
        const onSubmit = renderInput(false);
        const textarea = screen.getByRole("combobox");

        fireEvent.keyDown(textarea, { key: "Enter" });
        expect(onSubmit).not.toHaveBeenCalled();
    });

    it("ignores window file drops when canSend is false", () => {
        renderInput(false);

        const file = new File(["x"], "dropped.pdf", {
            type: "application/pdf",
        });
        const dataTransfer = {
            types: ["Files"],
            files: [file],
        } as unknown as DataTransfer;
        fireEvent.drop(window, { dataTransfer });

        expect(uploadProjectDocument).not.toHaveBeenCalled();
    });

    it("keeps the default composer when canSend is omitted", () => {
        render(
            <ChatInput
                onSubmit={vi.fn()}
                onCancel={vi.fn()}
                isLoading={false}
                projectId="p1"
            />,
        );

        expect(
            screen.getByPlaceholderText("How can I help?"),
        ).not.toBeDisabled();
        expect(
            screen.getByRole("button", { name: "Add documents" }),
        ).toBeInTheDocument();
    });
});


describe("Microsoft 365 composer opt-in", () => {
    beforeEach(() => {
        vi.clearAllMocks();
        window.localStorage.clear();
        vi.stubGlobal("ResizeObserver", ResizeObserverMock);
        mockProfile();
    });

    it("offers private mail and files next to EDGAR, sends opt-in and resets for another chat", () => {
        const onSubmit = vi.fn();
        const props = { onSubmit, onCancel: vi.fn(), isLoading: false, chatModel: "claude-sonnet-4-6", chatReasoningLevel: null };
        const { rerender } = render(<ChatInput {...props} chatKey="first" />);
        const toggle = screen.getByRole("button", { name: "Use Microsoft 365 mail and files" });
        expect(toggle).toHaveAttribute("aria-pressed", "false");
        expect(screen.getByRole("button", { name: "Use SEC EDGAR research" })).toBeInTheDocument();
        fireEvent.click(toggle);
        expect(toggle).toHaveAttribute("aria-pressed", "true");
        expect(screen.getByRole("button", { name: "Add documents" })).toBeEnabled();
        fireEvent.change(document.querySelector("textarea")!, { target: { value: "Which messages arrived?" } });
        fireEvent.keyDown(document.querySelector("textarea")!, { key: "Enter" });
        expect(onSubmit).toHaveBeenCalledWith(expect.objectContaining({ useMicrosoft365: true, useEdgar: false }));
        expect(localStorage.getItem("microsoft365Enabled")).toBeNull();
        rerender(<ChatInput {...props} chatKey="second" />);
        expect(toggle).toHaveAttribute("aria-pressed", "false");
    });

    it.each([false, true])("retains tools and selections when Microsoft 365 switches on and off (protected=%s)", (microsoft365Protected) => {
        const onSubmit = vi.fn();
        const ref = createRef<ChatInputHandle>();
        const document: Document = {
            id: "doc-1", project_id: null, filename: "contract.pdf", file_type: "pdf",
            storage_path: null, pdf_storage_path: null, size_bytes: null,
            page_count: null, structure_tree: null, status: "ready", created_at: null,
            current_version_id: "version-1", active_version_number: 2,
        };
        const workflow = { id: "workflow-1", title: "Review agreement" };
        render(<ChatInput ref={ref} onSubmit={onSubmit} onCancel={vi.fn()} isLoading={false}
            microsoft365Protected={microsoft365Protected} chatModel="claude-sonnet-4-6"
            chatReasoningLevel="high" />);
        fireEvent.click(screen.getByRole("button", { name: "Use SEC EDGAR research" }));
        const toggle = screen.getByRole("button", { name: "Use Microsoft 365 mail and files" });

        for (const enabled of [true, false]) {
            act(() => {
                ref.current!.addDoc(document);
                ref.current!.startWorkflow(workflow, "Review this with my mail and SEC filings");
            });
            fireEvent.click(toggle);
            expect(toggle).toHaveAttribute("aria-pressed", String(enabled));
            for (const name of ["Add documents", "Open workflows", "Use SEC EDGAR research", "Choose model", "Choose reasoning"]) {
                expect(screen.getByRole("button", { name })).toBeEnabled();
            }
            expect(screen.getByRole("button", { name: "Use SEC EDGAR research" })).toHaveAttribute("aria-pressed", "true");
            expect(screen.getByText("contract.pdf")).toBeInTheDocument();
            expect(screen.getByText(workflow.title)).toBeInTheDocument();
            fireEvent.click(screen.getByRole("button", { name: "Send message" }));
            expect(onSubmit).toHaveBeenLastCalledWith(expect.objectContaining({
                files: [{ document_id: document.id, filename: document.filename, version_id: "version-1", version_number: 2 }],
                workflow, useEdgar: true, useMicrosoft365: enabled,
                model: "claude-sonnet-4-6", reasoning: "high",
            }));
        }
        expect(onSubmit).toHaveBeenCalledTimes(2);
    });

    it.each([true, false])("allows slash workflows in protected chats with access enabled=%s", async (enabled) => {
        vi.mocked(listWorkflows).mockResolvedValueOnce([{
            id: "workflow-1", metadata: { title: "Contract Intake", name: "contract-intake" },
        } as Workflow]);
        const onSubmit = vi.fn();
        render(<ChatInput onSubmit={onSubmit} onCancel={vi.fn()} isLoading={false}
            microsoft365Protected initialMicrosoft365Enabled={enabled} chatModel="claude-sonnet-4-6" />);
        const textarea = screen.getByRole("combobox");
        fireEvent.change(textarea, { target: { value: "/contract-intake" } });
        await waitFor(() => expect(screen.getByRole("option")).toBeInTheDocument());
        fireEvent.keyDown(textarea, { key: "Enter" });
        expect(onSubmit).not.toHaveBeenCalled();
        expect(textarea).toHaveValue("");
        expect(screen.getByText("Contract Intake")).toBeInTheDocument();
        fireEvent.change(textarea, { target: { value: "Review this agreement" } });
        fireEvent.keyDown(textarea, { key: "Enter" });
        expect(onSubmit).toHaveBeenCalledWith(expect.objectContaining({
            workflow: { id: "workflow-1", title: "Contract Intake" }, useMicrosoft365: enabled,
        }));
    });

    it.each([true, false])("still clears hidden protected drafts with access enabled=%s", (enabled) => {
        render(<ChatInput onSubmit={vi.fn()} onCancel={vi.fn()} isLoading={false}
            microsoft365Protected initialMicrosoft365Enabled={enabled} />);
        const textarea = screen.getByRole("combobox");
        fireEvent.change(textarea, { target: { value: "Private draft" } });
        const hidden = vi.spyOn(document, "hidden", "get").mockReturnValue(true);
        try {
            fireEvent(document, new Event("visibilitychange"));
            expect(textarea).toHaveValue("");
        } finally {
            hidden.mockRestore();
        }
    });

    it("accepts document drops while Microsoft 365 is enabled", async () => {
        render(<ChatInput onSubmit={vi.fn()} onCancel={vi.fn()} isLoading={false} initialMicrosoft365Enabled />);
        const file = new File(["x"], "dropped.pdf", { type: "application/pdf" });
        fireEvent.drop(window, { dataTransfer: { types: ["Files"], files: [file] } });
        await waitFor(() => expect(uploadStandaloneDocuments).toHaveBeenCalledWith(
            [expect.objectContaining({ file })], expect.objectContaining({ onProgress: expect.any(Function) }),
        ));
    });

    it("still respects hidden controls and read-only access with Microsoft 365 enabled", () => {
        const props = { onSubmit: vi.fn(), onCancel: vi.fn(), isLoading: false, initialMicrosoft365Enabled: true };
        const { rerender } = render(<ChatInput {...props} hideAddDocButton hideWorkflowButton hideEdgarToggle />);
        for (const name of ["Add documents", "Open workflows", "Use SEC EDGAR research"]) {
            expect(screen.queryByRole("button", { name })).not.toBeInTheDocument();
        }
        rerender(<ChatInput {...props} canSend={false} />);
        expect(screen.getByRole("combobox")).toBeDisabled();
        expect(screen.getByRole("button", { name: "Send message" })).toBeDisabled();
        expect(screen.queryByRole("button", { name: "Use Microsoft 365 mail and files" })).not.toBeInTheDocument();
    });

    it("hydrates access from delayed history without overwriting later manual choices", () => {
        const props = { onSubmit: vi.fn(), onCancel: vi.fn(), isLoading: false, chatKey: "loaded-chat" };
        const { rerender } = render(<ChatInput {...props} />);
        const toggle = screen.getByRole("button", { name: "Use Microsoft 365 mail and files" });
        rerender(<ChatInput {...props} initialMicrosoft365Enabled />);
        expect(toggle).toHaveAttribute("aria-pressed", "true");
        fireEvent.click(toggle);
        rerender(<ChatInput {...props} initialMicrosoft365Enabled={false} />);
        rerender(<ChatInput {...props} initialMicrosoft365Enabled />);
        expect(toggle).toHaveAttribute("aria-pressed", "false");
        rerender(<ChatInput {...props} chatKey="another-chat" initialMicrosoft365Enabled />);
        expect(toggle).toHaveAttribute("aria-pressed", "true");
    });

    it("does not offer Microsoft 365 inside projects", () => {
        renderInput(true);
        expect(screen.queryByRole("button", { name: "Use Microsoft 365 mail and files" })).not.toBeInTheDocument();
    });
});
