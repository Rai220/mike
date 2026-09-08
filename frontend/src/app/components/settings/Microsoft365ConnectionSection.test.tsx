import { cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { StrictMode } from "react";

const mocks = vi.hoisted(() => ({
    getMicrosoft365Connection: vi.fn(),
    connectMicrosoft365: vi.fn(),
    checkMicrosoft365MailAccess: vi.fn(),
    disconnectMicrosoft365: vi.fn(),
    needsMfaVerification: vi.fn(),
}));

vi.mock("@/app/lib/mikeApi", async (importOriginal) => ({
    ...await importOriginal<typeof import("@/app/lib/mikeApi")>(),
    ...mocks,
}));
vi.mock("@/app/components/popups/MfaVerificationPopup", () => ({
    needsMfaVerification: mocks.needsMfaVerification,
    MfaVerificationPopup: ({ open, onVerified, onCancel }: {
        open: boolean; onVerified: () => void; onCancel: () => void;
    }) => open ? <div role="dialog" aria-label="Verify identity">
        <button type="button" onClick={onVerified}>Verify</button>
        <button type="button" onClick={onCancel}>Cancel</button>
    </div> : null,
}));

import { MikeApiError } from "@/app/lib/mikeApi";
import { Microsoft365ConnectionSection } from "./Microsoft365ConnectionSection";

const connection = {
    id: "connection-1", status: "connected", accountLabel: "alice@example.com",
    connectedAt: "2026-09-07T06:00:00Z", lastCheckedAt: null,
};

describe("Microsoft365ConnectionSection", () => {
    beforeEach(() => {
        vi.resetAllMocks();
        window.history.replaceState(null, "", "/settings/connectors");
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: true, connection: null });
        mocks.needsMfaVerification.mockResolvedValue(false);
    });
    afterEach(() => {
        cleanup();
        vi.unstubAllGlobals();
        vi.restoreAllMocks();
    });

    it("shows loading and hides an unavailable unconnected feature", async () => {
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: false, connection: null });
        render(<Microsoft365ConnectionSection />);
        expect(screen.getByText("Loading connection…")).toBeInTheDocument();
        await waitFor(() => expect(screen.queryByRole("region", { name: "Microsoft 365" })).toBeNull());
    });

    it("keeps disconnect available after the feature becomes unavailable", async () => {
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: false, connection });
        render(<Microsoft365ConnectionSection />);
        expect(await screen.findByText("alice@example.com")).toBeInTheDocument();
        expect(screen.getByRole("button", { name: "Disconnect" })).toBeEnabled();
        expect(screen.queryByRole("button", { name: "Reconnect" })).toBeNull();
        expect(screen.queryByRole("button", { name: "Check mail access" })).toBeNull();
    });

    it("connects using same-tab navigation to Microsoft only", async () => {
        const assign = vi.fn();
        const realWindow = window;
        vi.stubGlobal("window", new Proxy(realWindow, {
            get(target, key) {
                if (key === "location") return { href: realWindow.location.href, assign };
                return Reflect.get(target, key);
            },
        }));
        const authorizationUrl = "https://login.microsoftonline.com/tenant/oauth2/v2.0/authorize?state=abc";
        mocks.connectMicrosoft365.mockResolvedValue({ authorizationUrl, callbackOrigin: "https://api.example.com" });
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Connect Microsoft 365" }));
        await waitFor(() => expect(assign).toHaveBeenCalledWith(authorizationUrl));
        expect(screen.getByText("Connect Outlook mail and corporate files for personal Microsoft 365 chats.")).toBeInTheDocument();
    });

    it.each([
        "http://login.microsoftonline.com/authorize",
        "https://login.microsoftonline.com.evil.example/authorize",
        "https://login.microsoftonline.com:444/authorize",
        "https://user:pass@login.microsoftonline.com/authorize",
        "javascript:alert(1)",
    ])("rejects an unsafe authorization URL: %s", async (authorizationUrl) => {
        mocks.connectMicrosoft365.mockResolvedValue({ authorizationUrl, callbackOrigin: "https://api.example.com" });
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Connect Microsoft 365" }));
        expect(await screen.findByRole("alert")).toHaveTextContent("Microsoft 365 action failed. Please try again.");
    });

    it.each([
        ["pending", "Sign-in incomplete"],
        ["reconnect_required", "Reconnect required"],
    ])("renders %s and offers reconnect", async (status, label) => {
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: true, connection: { ...connection, status } });
        render(<Microsoft365ConnectionSection />);
        expect(await screen.findByText(label)).toBeInTheDocument();
        expect(screen.getByRole("button", { name: "Reconnect" })).toBeEnabled();
        expect(screen.queryByRole("button", { name: "Check mail access" })).toBeNull();
    });

    it("checks mail access and disconnects the owned connection", async () => {
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: true, connection });
        mocks.checkMicrosoft365MailAccess.mockResolvedValue({ ok: true, mailAccess: true, checkedAt: "2026-09-07T07:00:00Z" });
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Check mail access" }));
        expect(await screen.findByText("Mail access verified.")).toBeInTheDocument();
        expect(mocks.checkMicrosoft365MailAccess).toHaveBeenCalledWith(connection.id);
        fireEvent.click(screen.getByRole("button", { name: "Disconnect" }));
        expect(await screen.findByText("Not connected")).toBeInTheDocument();
        expect(mocks.disconnectMicrosoft365).toHaveBeenCalledWith(connection.id);
    });

    it("refreshes connection state after a failed check and hides internal errors", async () => {
        mocks.getMicrosoft365Connection.mockResolvedValueOnce({ available: true, connection })
            .mockResolvedValue({ available: true, connection: { ...connection, status: "reconnect_required" } });
        mocks.checkMicrosoft365MailAccess.mockRejectedValue(new MikeApiError({ status: 500, message: "SECRET provider token response" }));
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Check mail access" }));
        expect(await screen.findByText("Reconnect required")).toBeInTheDocument();
        expect(screen.getByRole("alert")).toHaveTextContent("Microsoft 365 action failed. Please try again.");
        expect(screen.queryByText(/SECRET/)).toBeNull();
    });

    it("defers a sensitive action until MFA is verified", async () => {
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: true, connection });
        mocks.needsMfaVerification.mockResolvedValueOnce(true).mockResolvedValue(false);
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Disconnect" }));
        expect(await screen.findByRole("dialog")).toBeInTheDocument();
        expect(mocks.disconnectMicrosoft365).not.toHaveBeenCalled();
        fireEvent.click(screen.getByRole("button", { name: "Verify" }));
        expect(await screen.findByText("Not connected")).toBeInTheDocument();
    });

    it("handles server-required MFA and cancellation", async () => {
        mocks.connectMicrosoft365.mockRejectedValue(new MikeApiError({ status: 403, code: "mfa_verification_required", message: "Verify MFA" }));
        render(<Microsoft365ConnectionSection />);
        fireEvent.click(await screen.findByRole("button", { name: "Connect Microsoft 365" }));
        expect(await screen.findByRole("dialog")).toBeInTheDocument();
        fireEvent.click(screen.getByRole("button", { name: "Cancel" }));
        expect(screen.queryByRole("dialog")).toBeNull();
        expect(screen.getByRole("button", { name: "Connect Microsoft 365" })).toBeEnabled();
    });

    it("consumes a successful callback preserving other query parameters and history state", async () => {
        window.history.replaceState({ kept: true }, "", "/settings/connectors?tab=mail&microsoft365=connected#account");
        mocks.getMicrosoft365Connection.mockResolvedValue({ available: true, connection });
        render(<StrictMode><Microsoft365ConnectionSection /></StrictMode>);
        expect(await screen.findByText("Microsoft 365 connected.")).toBeInTheDocument();
        expect(window.location.search).toBe("?tab=mail");
        expect(window.location.hash).toBe("#account");
        expect(window.history.state).toEqual({ kept: true });
    });

    it("does not claim connected based solely on a query parameter", async () => {
        window.history.replaceState(null, "", "/settings/connectors?microsoft365=connected");
        render(<Microsoft365ConnectionSection />);
        await screen.findByText("Not connected");
        expect(screen.queryByText("Microsoft 365 connected.")).toBeNull();
    });

    it("shows a fixed callback error without reflecting query content", async () => {
        window.history.replaceState(null, "", "/settings/connectors?microsoft365=error");
        render(<Microsoft365ConnectionSection />);
        expect(await screen.findByRole("alert")).toHaveTextContent("Microsoft 365 sign-in could not be completed. Try connecting again.");
        expect(window.location.search).toBe("");
    });

    it("shows a safe connection loading error", async () => {
        mocks.getMicrosoft365Connection.mockRejectedValue(new Error("postgres private credentials"));
        render(<Microsoft365ConnectionSection />);
        expect(await screen.findByRole("alert")).toHaveTextContent("Failed to load Microsoft 365 connection.");
    });
});
