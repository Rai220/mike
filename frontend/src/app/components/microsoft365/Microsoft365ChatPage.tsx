"use client";

import Link from "next/link";
import { useEffect, useRef, useState } from "react";
import { PageHeader } from "@/app/components/shared/PageHeader";
import { TABLE_CHECKBOX_CLASS } from "@/app/components/shared/TablePrimitive";
import { PillButton } from "@/app/components/ui/pill-button";
import { TabPillButton } from "@/app/components/ui/tab-pill-button";
import { Input } from "@/app/components/ui/input";
import { GlassCard } from "@/app/components/ui/glass-card";
import { CitationPillUI } from "@/shared/ui/CitationPillUI";
import {
    createMicrosoft365Chat, deleteMicrosoft365Chat, getMicrosoft365Chat,
    getMicrosoft365ChatSource, getMicrosoft365Connection, listMicrosoft365Chats,
    searchMicrosoft365ChatSources, sendMicrosoft365ChatMessage,
    type Microsoft365Chat, type Microsoft365ChatList, type Microsoft365SearchItem,
    type Microsoft365SourceView,
} from "@/app/lib/mikeApi";
import { userFacingApiError } from "@/app/lib/userFacingError";

const focusClass = "rounded focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-blue-500/40 focus-visible:ring-offset-2";

function originalUrl(value?: string): string | undefined {
    try {
        const url = new URL(value ?? "");
        return url.protocol === "https:" && !url.username && !url.password ? url.href : undefined;
    } catch { return undefined; }
}

export function Microsoft365ChatPage() {
    const [listing, setListing] = useState<Microsoft365ChatList | null>(null);
    const [connectionId, setConnectionId] = useState<string | null>(null);
    const [chat, setChat] = useState<Microsoft365Chat | null>(null);
    const [kind, setKind] = useState<"mail" | "file">("mail");
    const [query, setQuery] = useState("");
    const [searchMode, setSearchMode] = useState<"search" | "recent">("search");
    const [results, setResults] = useState<Microsoft365SearchItem[]>([]);
    const [selected, setSelected] = useState<Microsoft365SearchItem[]>([]);
    const [more, setMore] = useState(false);
    const [offset, setOffset] = useState(0);
    const [message, setMessage] = useState("");
    const [viewer, setViewer] = useState<Microsoft365SourceView | null>(null);
    const [busy, setBusy] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const controller = useRef<AbortController | null>(null);
    const activeChatId = useRef<string | null>(null);

    function clearContent() {
        setChat(null); setResults([]); setSelected([]); setViewer(null);
        setQuery(""); setKind("mail"); setSearchMode("search"); setMessage(""); setMore(false); setOffset(0);
    }

    async function run(action: (signal: AbortSignal) => Promise<() => void>) {
        controller.current?.abort();
        const request = new AbortController();
        controller.current = request;
        setBusy(true); setError(null);
        try {
            const apply = await action(request.signal);
            if (!request.signal.aborted) apply();
        } catch (err) {
            if (!request.signal.aborted) {
                // Do not retain previously authorized derivatives after a failed check.
                clearContent();
                setError(userFacingApiError(err, "Microsoft 365 could not be accessed. Try again."));
            }
        } finally {
            if (controller.current === request) { controller.current = null; setBusy(false); }
        }
    }

    function refresh() {
        void run(async (signal) => {
            const [next, connection] = await Promise.all([
                listMicrosoft365Chats(signal), getMicrosoft365Connection(),
            ]);
            const current = activeChatId.current && next.available
                ? await getMicrosoft365Chat(activeChatId.current, signal) : null;
            return () => {
                setListing(next);
                setConnectionId(connection.connection?.status === "connected" ? connection.connection.id : null);
                setChat(current);
            };
        });
    }

    useEffect(() => {
        refresh();
        function visibilityChanged() {
            if (document.visibilityState === "hidden") {
                controller.current?.abort(); clearContent();
            } else { refresh(); }
        }
        document.addEventListener("visibilitychange", visibilityChanged);
        return () => {
            controller.current?.abort();
            document.removeEventListener("visibilitychange", visibilityChanged);
        };
        // The listener uses refs for the active chat and request; no sensitive state is persisted.
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    useEffect(() => {
        if (!chat) return;
        function checkExpiry() {
            if (Date.parse(chat!.expiresAt) > Date.now()) return;
            controller.current?.abort(); clearContent(); activeChatId.current = null;
            setError("This chat has expired. Create a new chat to continue.");
        }
        checkExpiry();
        const timer = window.setInterval(checkExpiry, 1000);
        return () => window.clearInterval(timer);
    }, [chat]);

    function openChat(id: string) {
        clearContent(); activeChatId.current = id;
        void run(async (signal) => {
            const next = await getMicrosoft365Chat(id, signal);
            return () => setChat(next);
        });
    }

    function newChat() {
        if (!connectionId) return;
        clearContent(); activeChatId.current = null;
        void run(async (signal) => {
            const created = await createMicrosoft365Chat(connectionId, signal);
            const next = await getMicrosoft365Chat(created.chatId, signal);
            const list = await listMicrosoft365Chats(signal);
            const recent = await searchMicrosoft365ChatSources(created.chatId, { kind: "mail", mode: "recent", query: "", offset: 0, selectedSourceRefs: [] }, signal);
            return () => {
                activeChatId.current = created.chatId; setChat(next); setListing(list);
                setSearchMode("recent"); setResults(recent.items); setOffset(recent.nextOffset); setMore(recent.more);
            };
        });
    }

    function search(append = false, mode = searchMode) {
        if (!chat || (mode === "search" && !query.trim())) return;
        const nextOffset = append ? offset : 0;
        setViewer(null); setSearchMode(mode);
        if (mode === "recent") { setKind("mail"); setQuery(""); }
        if (!append) { setResults([]); setMore(false); }
        void run(async (signal) => {
            const found = await searchMicrosoft365ChatSources(chat.chatId, { kind: mode === "recent" ? "mail" : kind, mode, query: mode === "recent" ? "" : query.trim(), offset: nextOffset, selectedSourceRefs: selected.map(source => source.ref) }, signal);
            return () => {
                setResults(found.items);
                setOffset(found.nextOffset); setMore(found.more);
            };
        });
    }

    function viewSource(ref: string) {
        if (!chat) return;
        setViewer(null);
        void run(async (signal) => {
            const source = await getMicrosoft365ChatSource(chat.chatId, ref, signal);
            return () => setViewer(source);
        });
    }

    const available = listing?.available && connectionId && listing.policy;
    const openOriginal = originalUrl(viewer?.webUrl);
    const usedSourceRefs = new Set([...(chat?.sources.map(source => source.ref) ?? []), ...selected.map(source => source.ref)]);

    return (
        <div className="flex h-full min-h-0 flex-col">
            <PageHeader>Microsoft 365</PageHeader>
            <div className="min-h-0 flex-1 overflow-y-auto p-4 space-y-4" aria-busy={busy}>
                <p className="text-sm text-muted-foreground">Personal chats with your Outlook mail and corporate files. Choose the sources you want to discuss.</p>
                {listing?.policy && <p className="text-xs text-muted-foreground">Model: {listing.policy.model}. Chats expire after {listing.policy.retentionDays} days.</p>}
                {!busy && !available && <p role="status" className="text-sm">Corporate chat is unavailable. Connect your Microsoft account in settings; your administrator must enable an approved model and retention policy.</p>}
                <Link href="/settings/connectors" className={`inline-block text-sm underline ${focusClass}`}>Microsoft 365 settings</Link>
                {error && <p role="alert" className="text-sm text-destructive">{error}</p>}
                {busy && <div role="status" className="flex items-center gap-3 text-sm text-muted-foreground"><span className="animate-pulse">Checking access and loading…</span><PillButton tone="white" type="button" onClick={() => { controller.current?.abort(); controller.current = null; setBusy(false); clearContent(); }}>Cancel request</PillButton></div>}
                <div className="flex flex-wrap gap-2">
                    <PillButton type="button" tone="black" disabled={!available || busy} onClick={newChat}>New chat</PillButton>
                    <PillButton type="button" tone="white" disabled={busy} onClick={() => { clearContent(); refresh(); }}>Refresh access</PillButton>
                </div>
                <div className="grid gap-4 lg:grid-cols-[15rem_minmax(0,1fr)]">
                    <nav aria-label="Microsoft 365 chats" className="space-y-2">
                        {listing?.chats.map((item, index) => <div key={item.id} className="flex items-center gap-2">
                            <button type="button" disabled={busy || !available} aria-current={chat?.chatId === item.id ? "page" : undefined} onClick={() => openChat(item.id)} className={`min-w-0 flex-1 p-2 text-left text-sm hover:bg-app-surface-hover ${focusClass}`}>
                                Chat {index + 1}<span className="block text-xs text-muted-foreground">{new Date(item.createdAt).toLocaleString()}</span>
                            </button>
                            <PillButton type="button" tone="white" disabled={busy} aria-label={`Delete chat ${index + 1}`} onClick={() => {
                                clearContent(); if (activeChatId.current === item.id) activeChatId.current = null;
                                void run(async (signal) => { await deleteMicrosoft365Chat(item.id, signal); const list = await listMicrosoft365Chats(signal); return () => setListing(list); });
                            }}>Delete</PillButton>
                        </div>)}
                        {listing && !listing.chats.length && <p className="text-sm text-muted-foreground">No personal chats yet.</p>}
                    </nav>
                    <div className="min-w-0 space-y-4">
                        {chat && <>
                            <GlassCard><section aria-label="Choose sources" className="space-y-3 p-4">
                                <h2 className="text-sm font-medium">Choose sources</h2>
                                <p className="text-xs text-muted-foreground">Each chat supports up to 5 distinct sources, including earlier messages, and 20 questions. Start a new chat to use an updated source version.</p>
                                <div className="flex gap-2">{(["mail", "file"] as const).map(value => <TabPillButton key={value} active={kind === value} disabled={busy} onClick={() => { setKind(value); setSearchMode("search"); setResults([]); setMore(false); setOffset(0); }}>{value === "mail" ? "Mail" : "Files"}</TabPillButton>)}</div>
                                {kind === "mail" && <PillButton type="button" tone="white" disabled={busy} onClick={() => search(false, "recent")}>Recent mail</PillButton>}
                                <form className="flex gap-2" onSubmit={event => { event.preventDefault(); search(false, "search"); }}>
                                    <Input aria-label="Search Microsoft 365" maxLength={500} value={query} disabled={busy} onChange={event => { setQuery(event.target.value); setSearchMode("search"); setResults([]); setMore(false); setOffset(0); }} placeholder={kind === "mail" ? "Search your mail" : "Search corporate files"} />
                                    <PillButton type="submit" tone="white" disabled={busy || !query.trim()}>Search</PillButton>
                                </form>
                                {searchMode === "recent" && <p className="text-xs text-muted-foreground">Recent inbox messages, newest first. Select the messages you want to discuss.</p>}
                                <div className="space-y-2">{results.map(item => <label key={item.ref} className="flex items-start gap-2 text-sm break-words">
                                    <input type="checkbox" className={TABLE_CHECKBOX_CLASS} checked={selected.some(source => source.ref === item.ref)} disabled={busy || (usedSourceRefs.size >= 5 && !usedSourceRefs.has(item.ref))} onChange={event => setSelected(previous => event.target.checked ? [...previous, item] : previous.filter(source => source.ref !== item.ref))} />{item.title}
                                </label>)}</div>
                                {more && <PillButton type="button" tone="white" disabled={busy} onClick={() => search(true)}>Next results</PillButton>}
                                <p className="text-xs text-muted-foreground">{selected.length} selected · {usedSourceRefs.size} of 5 sources used in this chat</p>
                                {selected.map(item => <div key={item.ref} className="flex items-center gap-2 text-xs"><span className="min-w-0 break-words">{item.title}</span><button type="button" disabled={busy} className={`underline ${focusClass}`} aria-label={`Remove ${item.title}`} onClick={() => setSelected(previous => previous.filter(source => source.ref !== item.ref))}>Remove<span className="sr-only"> {item.title}</span></button></div>)}
                            </section></GlassCard>
                            <section aria-label="Conversation" className="space-y-3">
                                {chat.messages.map((entry, index) => <GlassCard key={index}><article className="space-y-2 p-4"><p className="text-xs font-medium text-muted-foreground">{entry.role === "user" ? "You" : "Mike"}</p><p className="whitespace-pre-wrap break-words text-sm">{entry.content}</p><div className="flex flex-wrap gap-2">{entry.sourceRefs.map(ref => {
                                    const sourceIndex = chat.sources.findIndex(source => source.ref === ref);
                                    const source = chat.sources[sourceIndex];
                                    return source ? <CitationPillUI key={ref} disabled={busy} aria-label={`View source: ${source.title}`} active={viewer?.ref === ref} className={focusClass} onClick={() => viewSource(ref)}>{sourceIndex + 1}</CitationPillUI> : null;
                                })}</div></article></GlassCard>)}
                                {!chat.messages.length && <p className="text-sm text-muted-foreground">Choose recent mail or search for a source, select it, then ask a question.</p>}
                            </section>
                            <form className="space-y-2" onSubmit={event => {
                                event.preventDefault(); if (!message.trim()) return;
                                setViewer(null);
                                void run(async (signal) => { const next = await sendMicrosoft365ChatMessage(chat.chatId, { message: message.trim(), sourceRefs: selected.map(source => source.ref) }, signal); return () => { setChat(next); setMessage(""); }; });
                            }}>
                                <label htmlFor="microsoft365-question" className="text-sm font-medium">Your question</label>
                                <textarea id="microsoft365-question" value={message} maxLength={8000} disabled={busy} onChange={event => setMessage(event.target.value)} rows={3} className={`w-full bg-app-surface p-3 text-sm ${focusClass}`} />
                                <PillButton type="submit" tone="black" disabled={busy || !message.trim() || (!selected.length && !chat.messages.length)}>Send</PillButton>
                            </form>
                        </>}
                        {viewer && <GlassCard><section aria-label="Source viewer" className="space-y-3 p-4"><div className="flex items-start justify-between gap-2"><h2 className="break-words text-sm font-medium">{viewer.title}</h2><PillButton type="button" tone="white" onClick={() => setViewer(null)}>Close source</PillButton></div><p className="text-xs text-muted-foreground">Checked {new Date(viewer.fetchedAt).toLocaleString()} · Version {viewer.version}</p><pre className="max-h-96 overflow-y-auto whitespace-pre-wrap break-words font-sans text-sm">{viewer.text}</pre>{openOriginal && <a href={openOriginal} target="_blank" rel="noopener noreferrer" referrerPolicy="no-referrer" className={`text-sm underline ${focusClass}`}>Open original in Microsoft 365</a>}</section></GlassCard>}
                    </div>
                </div>
            </div>
        </div>
    );
}
