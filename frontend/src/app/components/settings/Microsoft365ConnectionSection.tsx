"use client";

import { useEffect, useRef, useState } from "react";
import Link from "next/link";
import { SettingsSection } from "@/app/(pages)/settings/SettingsSection";
import { PillButton } from "@/app/components/ui/pill-button";
import {
    MfaVerificationPopup,
    needsMfaVerification,
} from "@/app/components/popups/MfaVerificationPopup";
import {
    checkMicrosoft365MailAccess,
    connectMicrosoft365,
    disconnectMicrosoft365,
    getMicrosoft365Connection,
    isMfaRequiredError,
    type Microsoft365ConnectionState,
} from "@/app/lib/mikeApi";
import { userFacingApiError } from "@/app/lib/userFacingError";

type Action = "connect" | "check" | "disconnect";

const statusLabels = {
    pending: "Sign-in incomplete",
    connected: "Connected",
    reconnect_required: "Reconnect required",
};

export function Microsoft365ConnectionSection() {
    const [state, setState] = useState<Microsoft365ConnectionState | null>(null);
    const [loading, setLoading] = useState(true);
    const [busy, setBusy] = useState<Action | null>(null);
    const [pendingMfa, setPendingMfa] = useState<Action | null>(null);
    const [error, setError] = useState<string | null>(null);
    const [notice, setNotice] = useState<string | null>(null);
    const callbackResultRef = useRef<string | null | undefined>(undefined);

    useEffect(() => {
        let cancelled = false;
        const url = new URL(window.location.href);
        if (callbackResultRef.current === undefined) {
            callbackResultRef.current = url.searchParams.get("microsoft365");
        }
        const callbackResult = callbackResultRef.current;
        if (url.searchParams.has("microsoft365")) {
            url.searchParams.delete("microsoft365");
            window.history.replaceState(
                window.history.state,
                "",
                `${url.pathname}${url.search}${url.hash}`,
            );
        }
        void getMicrosoft365Connection().then((next) => {
            if (cancelled) return;
            setState(next);
            if (
                callbackResult === "connected" &&
                next.connection?.status === "connected"
            ) {
                setNotice("Microsoft 365 connected.");
            } else if (callbackResult === "error") {
                setError("Microsoft 365 sign-in could not be completed. Try connecting again.");
            }
        }).catch((err: unknown) => {
            if (!cancelled) {
                setError(userFacingApiError(err, "Failed to load Microsoft 365 connection."));
            }
        }).finally(() => {
            if (!cancelled) setLoading(false);
        });
        return () => {
            cancelled = true;
        };
    }, []);

    async function runAction(action: Action) {
        setBusy(action);
        setError(null);
        setNotice(null);
        try {
            if (await needsMfaVerification()) {
                setPendingMfa(action);
                return;
            }
            if (action === "connect") {
                const { authorizationUrl } = await connectMicrosoft365();
                const url = new URL(authorizationUrl);
                if (
                    url.origin !== "https://login.microsoftonline.com" ||
                    url.username || url.password
                ) {
                    throw new Error("Invalid authorization destination");
                }
                window.location.assign(url.href);
            } else if (state?.connection) {
                if (action === "check") {
                    const result = await checkMicrosoft365MailAccess(state.connection.id);
                    setState({
                        ...state,
                        connection: { ...state.connection, lastCheckedAt: result.checkedAt },
                    });
                    setNotice("Mail access verified.");
                } else {
                    await disconnectMicrosoft365(state.connection.id);
                    setState({ ...state, connection: null });
                    setNotice("Microsoft 365 disconnected.");
                }
            }
        } catch (err) {
            if (isMfaRequiredError(err)) {
                setPendingMfa(action);
            } else {
                setError(userFacingApiError(err, "Microsoft 365 action failed. Please try again."));
                // A failed check may have marked the account as requiring a new login.
                if (action === "check") {
                    try {
                        setState(await getMicrosoft365Connection());
                    } catch {
                        // Keep the original error and the last known connection.
                    }
                }
            }
        } finally {
            setBusy(null);
        }
    }

    if (state && !state.available && !state.connection) return null;
    const connection = state?.connection;
    const disabled = busy !== null || pendingMfa !== null;

    return (
        <SettingsSection>
            <section className="space-y-3 px-4 py-5" aria-label="Microsoft 365" aria-busy={loading || busy !== null}>
                <div>
                    <h3 className="text-sm font-medium text-foreground">Microsoft 365</h3>
                    <p className="mt-1 text-sm text-muted-foreground">Connect Outlook mail and corporate files for personal Microsoft 365 chats.</p>
                </div>
                {loading ? (
                    <p role="status" className="animate-pulse text-sm text-muted-foreground">Loading connection…</p>
                ) : state && (
                    <>
                        <div className="space-y-1 text-sm">
                            <p>{connection ? statusLabels[connection.status] : "Not connected"}</p>
                            {connection?.accountLabel && <p className="break-words text-muted-foreground">{connection.accountLabel}</p>}
                            {!state.available && <p className="text-muted-foreground">Microsoft 365 is currently unavailable. You can still disconnect your account.</p>}
                        </div>
                        {state.available && connection?.status === "connected" && <Link href="/assistant" className="inline-block rounded text-sm underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500/40 focus-visible:ring-offset-2">Open Assistant</Link>}
                        {connection && <p className="text-xs text-muted-foreground">Reconnecting makes earlier Microsoft 365 chats unavailable. Disconnecting deletes the connection and its protected chats.</p>}
                        <div className="flex flex-wrap gap-2">
                            {state.available && (
                                <PillButton tone="white" type="button" disabled={disabled} onClick={() => void runAction("connect")}>
                                    {busy === "connect" ? "Connecting…" : connection ? "Reconnect" : "Connect Microsoft 365"}
                                </PillButton>
                            )}
                            {state.available && connection?.status === "connected" && (
                                <PillButton tone="white" type="button" disabled={disabled} onClick={() => void runAction("check")}>
                                    {busy === "check" ? "Checking…" : "Check mail access"}
                                </PillButton>
                            )}
                            {connection && (
                                <PillButton tone="white" type="button" disabled={disabled} onClick={() => void runAction("disconnect")}>
                                    {busy === "disconnect" ? "Disconnecting…" : "Disconnect"}
                                </PillButton>
                            )}
                        </div>
                    </>
                )}
                {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
                {notice && <p role="status" className="text-sm text-muted-foreground">{notice}</p>}
            </section>
            <MfaVerificationPopup
                open={pendingMfa !== null}
                onCancel={() => setPendingMfa(null)}
                onVerified={() => {
                    const action = pendingMfa;
                    setPendingMfa(null);
                    if (action) void runAction(action);
                }}
            />
        </SettingsSection>
    );
}
