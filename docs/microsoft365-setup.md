# Microsoft 365 protected chats in Mike

Mike connects a corporate account in Settings → Connectors. In the main
`/assistant` chat, turn Microsoft 365 on alongside the existing chat capabilities
and ask about Outlook mail or accessible OneDrive/SharePoint files. Mike lists,
searches and reads sources itself. The conversation keeps its ordinary chat ID;
its history moves into linked encrypted storage. The old `/assistant/microsoft365`
route redirects to Assistant, and no separate Microsoft sidebar entry is needed.

Once Microsoft 365 is used, the conversation stays private even after the toggle
is turned off. It cannot be shared, moved to a project, exported or used with
generic MCP, web/EDGAR, document-generation tools or user-supplied model keys.
OFF disables new Microsoft retrieval; it does not erase already shown history
or make that history ordinary plaintext.

The previous separate-chat version was deployed on RnD and verified with a
persistent Microsoft sign-in, live inbox listing and a source-based Sonnet
answer. The unified Assistant is now deployed with migration 03 and verified
through a same-chat ON → list/read → OFF → history reload sequence. RnD uses
direct `claude-sonnet-4-6` with seven-day
retention. File access, refresh/reconnect, two-user ACL acceptance and operational
requirements remain separate checks. See the [deployment record](microsoft365-rnd-deployment.md).

## Operator setup

1. Apply these migrations, in order, through the normal
   [deployment procedure](deployment.md):

   - `backend/migrations/20260907_01_microsoft365_connections.sql`
   - `backend/migrations/20260907_02_microsoft365_protected_chats.sql`
   - `backend/migrations/20260907_03_microsoft365_unified_chat.sql`

   Fresh installations already include all three shapes in `backend/schema.sql`.
   Do not replay unrelated historical migrations. Migration 02 deletes existing
   Microsoft connections whose owners no longer belong to their organization
   before adding the membership foreign key. Migration 03 links encrypted history
   to ordinary chats and adds permanent privacy/plaintext/share write guards.
   Verify that migration before serving the unified Assistant code.
2. Select the Mike organization and corresponding Entra tenant. Members connect
   their own Microsoft accounts. Organization administrators do not inherit
   access to another member's mail, files, chat history, or credentials.
3. Configure the following backend-only values in the deployment secret store:

   ```dotenv
   MICROSOFT365_ENABLED=true
   MICROSOFT365_CLIENT_ID=<application-client-id>
   MICROSOFT365_TENANT_ID=<directory-tenant-id>
   MICROSOFT365_CLIENT_SECRET=<application-secret-value>
   MICROSOFT365_ORG_ID=<mike-organization-uuid>
   MICROSOFT365_ENCRYPTION_SECRET=<dedicated-random-secret-at-least-32-characters>
   MICROSOFT365_ENCRYPTION_KEY_ID=key1
   MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS={}
   FRONTEND_URL=https://mike.example.com
   API_PUBLIC_URL=https://mike.example.com/api
   MICROSOFT365_FILES_ENABLED=false
   MICROSOFT365_CHAT_ENABLED=false
   MICROSOFT365_CHAT_MODEL=
   MICROSOFT365_RETENTION_DAYS=
   ```

   Generate a dedicated encryption secret with `openssl rand -hex 32`. Keep
   secrets stable across restarts. The ignored diagnostic file
   `backend/.env.microsoft365.local` is not automatically loaded by the product
   backend. Never use Next.js public variables or browser storage for these
   values.
4. Register an Entra **Web** redirect matching the complete backend callback,
   including the `/api` prefix if used:

   ```text
   https://mike.example.com/api/integrations/microsoft365/callback
   ```

   Preserve existing redirects. The localhost canary callback on port 8400 is a
   separate registration and does not validate this HTTPS product callback.
   Only the commercial Microsoft cloud is implemented.
5. Configure delegated Graph `User.Read`, `Mail.Read`, and `offline_access`.
   For corporate file search, set `MICROSOFT365_FILES_ENABLED=true`, grant
   delegated `Files.Read.All`, and reconnect existing accounts to obtain the
   added scope. `Files.Read` alone is insufficient for Graph Search `driveItem`.
   This does not request application permissions or automatically add
   `Sites.Read.All`. Tenant consent, user assignment, service licensing, and
   Conditional Access must permit the selected account and scopes.
6. After approval of the model/provider, processing location, and retention,
   configure `MICROSOFT365_CHAT_MODEL` to an exact supported model ID and
   `MICROSOFT365_RETENTION_DAYS` to an integer from 1 through 90, then set
   `MICROSOFT365_CHAT_ENABLED=true`. There is no default model or retention.
   Accepted providers are direct OpenAI, Anthropic, Google, or an explicitly
   configured Ollama endpoint; OpenRouter, Vercel gateway, and OpenCode Go models
   are rejected. Renamed model IDs and fallback resolution are rejected too.
   Provision the corresponding server credential/endpoint. The setting itself
   does not prove contractual approval or enforce a provider's processing region.
7. Run Mike's durable jobs worker with the same configuration as the API. The
   existing `db_jobs` schema and a running worker are required; Redis is optional.
   Embedded workers are the default. If the API uses `WORKERS_MODE=none`, run
   `node dist/worker.js` separately. See
   [background jobs](deployment.md#background-jobs-and-redis).
8. Deploy/restart backend, frontend, and workers. Reverse proxies/APM must omit
   OAuth callback query strings and protected request/response bodies. The
   application strips callback query values before auth/error logging, sets
   `no-store`/`no-referrer`, and disables raw LLM recording and SDK telemetry for
   protected requests. These controls do not configure infrastructure logs.

## User flow and limits

Sign in to Mike, complete enrolled MFA, and connect Microsoft 365 in Settings →
Connectors. Consent returns to the same Mike session. **Check mail access** only
checks permission to list messages and returns status/time; it reads no body.
Microsoft OAuth connects an external source and does not replace Mike sign-in.

Open Assistant, use a private conversation and enable Microsoft 365. Ask, for
example, “Какие письма у меня на почте?”; Mike can list recent subjects, then read
a relevant message when asked about its contents. The same conversation supports
follow-ups such as “Что в первом письме?”. No manual source-picker step is needed.
Shared and project conversations cannot be converted to this private mode.

The ON toggle exposes `microsoft365_recent_mail`, `microsoft365_search` and
`microsoft365_read`. These tools accept bounded search input or server-issued
opaque references, never arbitrary Graph URLs/credentials. Search is on demand;
it is not an exhaustive archive scan or synchronized index. The unified chat
allows up to **20 source references, 20 questions and 100,000 context characters**,
with at most 24 tool calls and six model iterations per turn. Start a new private
conversation when these limits are reached.

Turning Microsoft 365 OFF removes retrieval tools and prevents automatic new
mail/file body reads. Mike can discuss the already-shown conversation. The server
still checks metadata access and versions before disclosing retained history and
before/after generation; revoked or changed dependencies block that history.
Explicitly opening an original through an authorized source endpoint is a
separate read operation. OFF is not a network-isolation or data-erasure switch.

Mail is read as individual messages. Supported file types are TXT, Markdown,
CSV, PDF, DOCX, XLSX, and PPTX, with download, archive expansion, extraction,
text/context, and execution-time limits. Unsupported, oversized, encrypted, or
unreadable content is rejected. Scanned PDF OCR is not implemented.

Mike must read an original before discussing its contents; citations are checked
as verbatim substrings of originals read during that turn. Metadata-only inbox
lists and discussion with retrieval OFF may have no quotations. Corporate text
and quotations render as inert plain text, without executable model-generated
links or markup.

The normal Assistant SSE endpoint buffers corporate output until the answer,
source access and versions have passed validation and encrypted persistence has
succeeded. No partial corporate answer is released before these checks. Source
changes or revocation require a new chat for the new version. Already seen
information cannot be retracted from a person or screenshot; this is not a
device-level copy-control feature.

## Credentials, expiry, and deletion

- Tokens, PKCE verifiers, and protected chat payloads use AES-256-GCM with
  authenticated owner/tenant/connection context. Source titles and conversation
  text are encrypted together; the ordinary chat shell uses only the generic
  title `Microsoft 365`, with no source subjects or answer excerpts.
  The tables are backend-only with RLS and revoked browser grants.
- One-time OAuth state expires after ten minutes. Its state/session identifiers
  are hashed. Refresh uses an on-demand lease and conditional writes; normal
  operation does not poll mail or refresh all tokens in the background.
- Membership removal immediately cascades to connection credentials, OAuth
  states, and protected chats. Disconnect also cascades those records. Owner,
  connection generation, and conditional-version checks prevent stale writes
  from recreating deleted data. Disconnect does not revoke the Entra consent
  grant; that remains a Microsoft user/admin operation.
- The encrypted payload is linked to the ordinary chat ID, with a permanent
  privacy flag on the ordinary shell. Existing ordinary history is copied into
  the payload once; subsequent messages, including OFF turns, remain encrypted.
  Database triggers reject new plaintext messages and sharing grants. Deleting
  encrypted history or disconnecting does not reset that shell into an ordinary
  writable conversation or restart retention.
- Retention starts when the protected payload is created. Conversation activity never extends
  it. Shorter active retention applies to existing history; increasing the
  configured retention does not extend a stored expiry. Expired chats are denied
  immediately on access, even before their physical cleanup.
- Workers enqueue `microsoft365.maintenance` at startup and every ten minutes.
  The job removes expired history and OAuth states and purges stale connection
  generations. Keep `MICROSOFT365_ORG_ID` configured while disabling the feature
  so cleanup continues. Disabling UI/connection flags does not delete history
  immediately. Re-encryption requires valid enabled connection configuration;
  expiry cleanup runs even when that configuration is disabled.
- Backup expiry, snapshot deletion, restored-database cleanup, infrastructure
  logs, and provider-side retention require an operator policy. Database row
  deletion is not a claim that those independent copies have been erased.

## Encryption-key rotation

Key IDs are versioned envelope labels, not Azure secret IDs. New installations
should set a stable ID such as `key1`; pre-versioned ciphertext uses `legacy`.
Changing a secret while retaining its old ID makes old ciphertext unreadable.

1. Preserve the current secret and its current ID in the secret store.
2. Generate a new secret, choose a **different** `MICROSOFT365_ENCRYPTION_KEY_ID`,
   and set `MICROSOFT365_ENCRYPTION_SECRET` to that new value. Add the previous
   value under its old ID in `MICROSOFT365_PREVIOUS_ENCRYPTION_KEYS`, a JSON
   object such as `{"key1":"<previous-secret>"}`. For unversioned data, use
   `{"legacy":"<previous-secret>"}`. Up to eight previous keys are accepted.
3. Roll configuration consistently to API and workers. During a rolling
   deployment, every still-serving replica must be able to read both old and
   new envelope IDs; stage the new key in readers' key rings before promoting
   it to the active writer key.
4. Maintenance rewraps connection tokens and chat ciphertext in batches of at
   most 100 rows, using cursors and CAS checks. It preserves chat retention and
   activity timestamps. Inspect maintenance results and retry skipped races or
   reconnect-required records; completion of one batch is not full rotation.
5. Keep old keys until all cursor pages finish, skipped records are resolved,
   and no remaining connection/chat envelope references an old ID. Pending
   OAuth verifiers are not batch-rewritten: allow their ten-minute lifetime and
   a successful expiry cleanup to finish. Retire old keys only after these
   checks. Independently decide how encrypted backups remain readable or expire.

Correct key-ring rotation does not require reconnecting every account. Removing
an old key too early can require reconnecting credentials and can make retained
chat history unreadable.

## Verification and live acceptance

Automated coverage exercises connection/session ownership, source selection and
versions, exact quotations, encrypted history, CAS/disconnect races, retention,
rotation, provider privacy, routes, and frontend behavior. Disposable PostgreSQL
16 tests exercise real DDL replay, RLS/grants, ownership constraints, write guards,
cleanup, and cascades, with matching fresh/incremental Microsoft DDL. The fixture
uses synthetic auth dependencies; it is not a full deployed Supabase or real
corporate-directory test. See [TODO](../TODO.md) for the final verification record.

Live evidence includes the successful HTTPS product sign-in, persistent mail
access and an earlier separate-chat inbox/answer/quotation/history probe. Those
checks used the deployed direct Sonnet provider and deleted the temporary chat.
The current unified ON/OFF flow passed a synthetic browser smoke and separate
live server-runner checks, including zero new body/list/search calls while OFF
and zero ordinary plaintext messages. A real file Search returned HTTP 200 with
no hits, and the deployed adapter successfully returned an empty list. Separate
intermittent 20-second Graph timeouts were safely reported; they do not invalidate
the successful empty-result parser check. Refresh,
reconnect, nonempty file Search/downloads, two-user revoked-source isolation and
proxy/backup controls remain to be checked; successful single-account mail
access does not prove them.
Deferred: message attachments, automatic full-thread expansion, shared mailboxes,
Word add-in, Microsoft SSO for Mike, delta/subscription synchronization, semantic
RAG/indexing, sending mail, and modifying Microsoft files. These are outside this
unified private Assistant MVP. See the [unified-chat PRD](microsoft365-unified-chat-prd.md)
and [integration plan](microsoft-365-integration-plan.md).

Exact local commands, counts and verification limits: [verification record](microsoft365-verification.md).

A new OAuth attempt changes the connection generation. Existing protected chats
from the previous generation become inaccessible and are removed by maintenance,
even if the new sign-in is cancelled. Reconnecting retains the Microsoft account
binding, but does not restore the old protected conversation.
