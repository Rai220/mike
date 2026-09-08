# PRD: protected Microsoft 365 storage and execution

## Current product contract

The user explicitly requested Microsoft 365 as a toggle in the main Assistant,
like EDGAR. The initial separate `/assistant/microsoft365` page and subsequent
ordinary-chat handoff are superseded. The current UI/flow is defined in the
[unified Assistant PRD](microsoft365-unified-chat-prd.md). Account connection
remains in Settings → Connectors; the legacy chat route redirects to Assistant.

The security problem remains the same: Graph ACL checks authorize a fresh read
but cannot retract corporate text already stored in ordinary messages, exports,
shared documents or other users' histories. Keeping the same UI/chat identity
therefore requires protected persistence and guarded ordinary-chat operations.

## Persistence and authority

`microsoft365_chats` remains a backend-only encrypted store. Its optional unique
`ordinary_chat_id` links a private, owned ordinary chat to one protected payload.
AES-GCM binds owner, organization, tenant, connection/generation and linked chat
identity. Source subjects, provenance, messages and quotations remain encrypted.
The ordinary shell exposes the generic title `Microsoft 365`.

Migration 03 permanently sets `chats.microsoft365_protected` and rejects owner
changes, shared grants, project/organization moves, title changes and new
plaintext message writes. Row locks serialize promotion against sharing and
plaintext inserts. The linked identity and expiry cannot be replaced or extended.
Existing ordinary history is copied into the encrypted payload once; all later
turns use that payload, including turns with Microsoft retrieval OFF.

The normal chat GET hydrates protected history only through owner/connection,
expiry and source ACL/version validation. The normal SSE message route uses
protected generation/persistence rather than ordinary message inserts. Sharing,
export, title mutation and related ordinary routes also guard protected shells;
SQL guards remain the final boundary for direct browser/database writes.

## Retrieval and model execution

The same Mike system prompt gains a Microsoft capability section. ON exposes only
three read-only Graph tools: recent inbox, mail/file search and source read by
opaque reference. No generic MCP, public search/EDGAR, document tools, write
operations, BYOK, broker fallback or model-selected arbitrary URL requests are
allowed in a protected conversation. Source metadata/text are untrusted evidence.

OFF exposes no retrieval tools and performs no automatic source-body reads.
Already displayed history can be discussed after metadata-only access/version
checks. Turning OFF does not remove source dependencies, downgrade encryption,
allow sharing, restore ordinary tools or erase history. Explicit authorized source
viewing is a separate read operation.

Before history/model disclosure and after generation, every transitive source
dependency receives metadata access/version validation. Reading an original
requires a current version; quotations must occur verbatim in originals read in
that turn. Metadata-only listing and OFF history discussion may have no citations.
The normal SSE route releases the completed answer only after validation and
encrypted save, with corporate text rendered as inert plain text.

Limits for the unified flow: 20 source references, 20 user questions, 100,000
context characters, 24 tool calls and six model iterations per turn. Existing
bounded Office/PDF parsers and Graph transport checks are reused without new
dependencies. Unsupported files fail explicitly; source changes or access
revocation block dependent history.

## Policy and lifecycle

An explicit valid `MICROSOFT365_CHAT_MODEL` and `MICROSOFT365_RETENTION_DAYS`
(1–90) are required with `MICROSOFT365_CHAT_ENABLED=true`. Only supported exact
direct provider model IDs are accepted. RnD currently uses `claude-sonnet-4-6`
and seven-day retention. Configuration does not independently attest processing
region, provider contracts or backup deletion.

Retention starts with protected-payload creation and never extends through
activity. Owner deletion, disconnect and membership removal purge protected
credentials/history as applicable; expiry checks reject reads before background
cleanup. The durable maintenance worker cleans expired states/history and stale
connection generations every ten minutes. Versioned keys and batch CAS rotation
remain unchanged; operator procedures are in [setup](microsoft365-setup.md).

Human memory/screenshots, infrastructure logs, provider retention and independent
backup copies remain outside application row-deletion guarantees.

## Verification and remaining scope

Current local verification covers the unified route/tool toggle, encrypted
history, guarded ordinary readers/writers, privacy races, source validation,
provider allowlists and UI behavior. Real disposable PostgreSQL 16 tests validate
migrations 01–03, write guards and fresh/incremental schema convergence with
synthetic auth stubs. Exact totals are in the [verification record](microsoft365-verification.md).

Earlier RnD evidence proved persistent OAuth/mail access and a selected-mail
answer/quotation through the separate chat. The unified version is now deployed:
same-chat listing/body read/quotation, retrieval OFF with metadata-only checks,
encrypted history reload and zero ordinary plaintext messages were verified live.
A synthetic browser test also passed. A real empty file Search returned HTTP 200
and the deployed adapter successfully parsed it without a `hits` field. Separate
intermittent 20-second Graph timeouts were safely reported. Refresh/reconnect, nonempty file
access/downloads, two-user ACL changes and operational backup/proxy acceptance
remain separate checks.

Deferred: attachments/full threads, shared mailboxes, Word add-in, Microsoft SSO,
synchronization/indexing, sending mail and modifying Microsoft files. Each needs
an explicit extension of the source, permission and protected-derivative contract.
