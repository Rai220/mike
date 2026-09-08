# RnD deployment — Microsoft 365, 2026-09-07

Deployment history after the user explicitly requested the server rollout.
The unified Assistant is deployed and its mail ON/OFF flow is live-verified.
Earlier separate-chat rollouts below are historical evidence. Final file-search
parser verification is recorded separately below.

- Target: SSH alias `rnd`, `/opt/mike`.
- Current chat entry: https://lispenard.nl.tuna.am/assistant
- Legacy `/assistant/microsoft365` redirects to `/assistant` in the unified version.
- Connection settings: https://lispenard.nl.tuna.am/settings/connectors
- Web callback used for the successful Mike sign-in: `https://lispenard.nl.tuna.am/api/integrations/microsoft365/callback`.

## Initial separate-chat rollout — historical

Transferred 49 feature files, merging against the existing server tree. Existing
EDGAR, branding, UI and deployment changes were preserved. Backend/frontend
package manifests and lockfiles already matched locally; no dependency changes
were needed. Both Docker images built successfully on RnD.

Before changes, saved a custom-format database dump, source/config archive and
previous image IDs under `/opt/mike-deploy-backups/microsoft365-20260907T091527Z`.
Backups and configuration are private. Previous images remain tagged
`mike-backend:before-m365` and `mike-frontend:before-m365`.

Applied only `20260907_01_microsoft365_connections.sql` and
`20260907_02_microsoft365_protected_chats.sql`, in order, using
`psql -v ON_ERROR_STOP=1`. The full fresh schema was not replayed.

Created an empty Microsoft 365 organization for the existing Kostya account:
`5b48ef5f-bb18-4191-905a-4fb4b9781269`. That account is its only member/admin.
Other existing accounts were not added. Bound it to the supplied Microsoft tenant.

Stored the supplied application credentials and a newly generated dedicated
encryption key in the server's private `backend/.env`; verified Compose passes
the values without printing them. Enabled Microsoft connection and file scopes;
corporate model processing was disabled in this initial deployment. Embedded
worker mode is `thread`. The follow-up rollout below enables the existing user-selected
model with a seven-day retention period; rollout verification passed.

Ran `docker compose up -d --no-deps backend frontend`; database, auth, storage,
Redis and the existing outer access controls were preserved.

## Initial rollout verification — historical

- Backend and frontend containers are running the rebuilt images.
- Backend health: HTTP 200.
- New frontend route through local frontend: HTTP 200.
- Microsoft connection and chat APIs through local frontend: HTTP 401 without
  a Mike session, as required.
- Initial backend service check for Kostya: configuration valid and connection
  available. Subsequent user sign-in established a persistent connection:
  `status: connected`; `checkMicrosoft365Connection` returned `mailAccess: true`.
- A real synthetic `claude-sonnet-4-6` request through the existing direct provider
  returned the expected marker. No corporate source was included in that probe.
- `microsoft365.maintenance` durable job: `done`.
- The public origin has an existing outer authentication gate; unauthenticated
  public requests are not evidence of an absent route. Use the normal server login.

The deployment manifest is `/opt/mike/.microsoft365-deployment.json`.
Local test evidence remains in [verification](microsoft365-verification.md).

## Separate mail entry fix — historical, deployed and verified

After connecting, the user asked the ordinary assistant about their inbox and
received a denial of email capability. The ordinary chat did not expose the
protected chat, and corporate model processing was still disabled on RnD.

That follow-up code added a Microsoft 365 sidebar entry, a recent-inbox picker
loaded when creating a protected chat, and a deterministic handoff from ordinary
mail requests before the general LLM/tool loop. Corporate metadata and content
stay inside the protected chat; listing recent mail does not select sources or
invoke the model.

The verified runtime configuration uses the user's existing selected model and
existing direct Anthropic credential:

```dotenv
MICROSOFT365_CHAT_ENABLED=true
MICROSOFT365_CHAT_MODEL=claude-sonnet-4-6
MICROSOFT365_RETENTION_DAYS=7
```

Transferred 12 scoped source/test files after comparing them against the previous
deployment and repository baseline; no merge conflicts or unrelated overrides.
Saved source/config and previous image IDs under
`/opt/mike-deploy-backups/mail-entry-20260907T112552Z`. Previous images are tagged
`mike-backend:before-mail-entry` and `mike-frontend:before-mail-entry`.
Both Docker production builds passed, then backend/frontend were recreated with
`docker compose up -d --no-deps backend frontend`. No SQL migration was needed.

Live checks after restart:

- Both containers running; health and protected frontend route HTTP 200.
- Connection and protected-chat APIs return HTTP 401 without a Mike session.
- Persistent connection active; runtime policy exactly matches the values above.
- Ordinary `runLLMStream` with “Какие письма у меня на почте?” emits the absolute
  protected-chat link and completes the stream before the general LLM/tool loop.
- A temporary protected chat listed one real inbox message and had seven-day expiry.
- Selected that message, generated an answer with a verified source quotation,
  reloaded encrypted history, and opened the freshly checked source viewer.
- Deleted the temporary chat in a finally block. No mail subject/body/model answer
  or credentials were printed in the verification output.

## Unified Assistant rollout — deployed and mail flow verified

The user's updated requirement replaces the separate page and handoff with a
Microsoft 365 toggle in the ordinary Assistant, alongside existing capabilities
such as EDGAR. The same chat ID and SSE route are retained. With the toggle ON,
Mike can list recent inbox mail, search mail/files and read sources through three
read-only Graph tools. OFF removes these tools and new body retrieval; existing
history still requires metadata-only ACL/version checks. Other external/document
tools remain unavailable in a conversation that has used Microsoft 365.

Encrypted history is linked through `microsoft365_chats.ordinary_chat_id`.
Migration `20260907_03_microsoft365_unified_chat.sql` makes the ordinary chat's
protected flag permanent and guards sharing, project/organization moves,
plaintext inserts and title leaks. Existing ordinary history is copied once into
the encrypted payload; subsequent turns remain protected, including OFF turns.
The legacy Microsoft route redirects to Assistant; the extra sidebar link is gone.

The existing direct `claude-sonnet-4-6` model and seven-day retention are retained.

Before the rollout, saved the custom PostgreSQL dump (853,838 bytes), scoped
source/config archive and previous image IDs under
`/opt/mike-deploy-backups/unified-m365-20260907T130837Z`. Previous images use the
`before-unified-m365` tag. The initial scoped transfer contained 40 source/test
files; the updated browser spec was transferred separately. Applied migration 03 with `ON_ERROR_STOP`; backend/frontend Docker builds
passed and both services were restarted. Backend health and `/assistant` returned
HTTP 200; API requests without a Mike session returned HTTP 401.

Live server checks used one temporary ordinary chat ID throughout:

- ON recent-mail tool listed one source without reading its body.
- A follow-up read the body and returned a validated quotation.
- OFF handled a request to refresh the list with **zero new source-body, list or
  search calls**; only metadata ACL/version checks ran.
- Encrypted history reloaded; ordinary `chat_messages` contained zero plaintext
  messages. Deleting the temporary ordinary chat cascaded to the linked payload.
- Runtime policy remained direct `claude-sonnet-4-6` with seven-day retention.

Two real model response differences were fixed and redeployed: citations based
on listed titles in metadata-only replies, and plain prose in retrieval-OFF
replies. Source-body quotation validation remains required when bodies are read.

The current synthetic Chromium main-Assistant test passed: same chat ID, toggle
ON/OFF, reload, inert corporate text and mobile layout. It complements the live
server-runner checks; it does not establish a live browser OAuth flow. Verification
output contained status/counts/booleans, not credentials or corporate content.

### Empty file-search parser — deployed and live-verified

The final deployed artifact includes the `providerHits` parser fix. A live file
Search for `test` returned Graph HTTP 200 in 776 ms, with only `total: 0` and
`moreResultsAvailable: false` in the result container and no `hits` field. The
actual `searchMicrosoft365Sources` adapter completed successfully with `items: []`.
This proves the empty result is handled correctly; no matching file was available
for a live body-download check.

Separate intermittent 20-second Graph timeouts were observed and safely reported.
They are provider-request timeouts, not failures to parse the successful empty
response. No claim of automatic retry or removal of upstream timeouts is made.
Final backend/frontend builds and restarts completed successfully; the final
broad backend run passed 352 tests in 14 files.

## Remaining corporate acceptance

The HTTPS sign-in and persistent mail connection have succeeded. Refresh and
reconnect, nonempty file Search/download access, two-user ACL acceptance, and
operational backup/retention requirements remain
separate checks. The server requests delegated `Files.Read.All`; file access is
not proven by successful mail access.

## Sent-mail folder fix — 2026-09-07

The recent-mail native tool previously queried Inbox only. The main Assistant now
uses `folder=all` by default, `sentitems` for sent/outgoing questions, and `inbox`
for explicit incoming-mail questions. Results identify their folder scope. The
legacy adapter default remains Inbox for compatibility. Only IDs and subjects
are listed; body reads still require the separate read tool.

Live Graph checks returned two messages across the mailbox and one in Sent Items.
The first runner checks exposed a second issue: a valid exact five-character
subject citation was rejected by the eight-character quotation minimum. Complete,
nonblank short titles are now accepted only for metadata-only answers; partial
short title quotes and short body quotes remain rejected.

Verification: 139 tests passed in the four targeted sources, assistant, legacy
chat and ordinary-route suites; `npm run build --prefix backend` and
`git diff --check` passed. Backend Docker build/restart succeeded on RnD, with
backup `/opt/mike-deploy-backups/sent-mail-20260907T135207Z`; no SQL was needed.
The final live main-Assistant runner invoked Sent Items, returned one source and
an answer in 6,158 ms, and accepted the exact short title citation. Plaintext
`chat_messages` count was zero; the temporary chat was deleted. Probe output
contained counts/booleans, not subjects, message bodies, answers or credentials.

## Independent composer capabilities — 2026-09-07

Supersedes the earlier main-Assistant restrictions above. The Microsoft 365
switch now controls only new Graph retrieval. Attachments, slash/selected
workflows, EDGAR, model and reasoning controls remain available in ON and OFF
states; toggling Microsoft 365 does not clear selections. Reload restores the
last submitted access switch and encrypted composer metadata. A later manual
switch change takes precedence over delayed history hydration.

The protected runner composes Graph tools with explicitly allowed local
read/find/list/fetch, workflow and EDGAR tools. Selected workflows supply
instructions; local documents use the bounded, isolated file parser. Document,
project and workflow access plus version fingerprints are checked before use and
before exposing persisted history. Model/reasoning preferences, attachment and
workflow metadata stay in the encrypted payload. User-selected models use keys
loaded from the authenticated user's backend settings; the environment model is
only the fallback when no model has been selected. Credentials are never saved
in the conversation. Sensitive SDK logging/recording/telemetry remain disabled.

The legacy Graph-only runner keeps its narrower contract. Sharing, project moves
and plaintext chat persistence remain blocked for protected conversations.
Workflow operations that create/edit documents are not introduced by this patch:
the runner can apply workflow instructions and draft text, but cannot claim to
have created a document artifact. General MCP and mutation tools remain absent.

Validation: 436 backend tests passed in 15 files, followed by 13 capability tests
after aligning document search with the shared whitespace/budget contract.
These suites overlap and must not be summed. Backend TypeScript build passed.
38 targeted frontend tests, TypeScript and scoped ESLint passed. The broader
frontend run encountered five export tests failing because jsdom Blob.text was
unavailable; the changed getChat mapping tests passed. Synthetic Chromium passed
one test covering ON/reload ON/OFF/reload OFF, independent enabled controls,
EDGAR selection, inert source text and desktop/mobile layout. Screenshots were
inspected; the browser API was mocked.

RnD backend/frontend Docker builds and restarts succeeded. No SQL migration was
needed. Source/image backup:
`/opt/mike-deploy-backups/combined-m365-20260907T142457Z`.
A live main-runner request used Graph Sent Items, a synthetic stored attachment,
a selected synthetic workflow and public EDGAR research together (16,385 ms).
The response passed workflow/budget checks; reload preserved the attachment,
workflow, selected Sonnet model and low reasoning. Ordinary `chat_messages`
contained zero plaintext messages. A following OFF turn still read the local
document and used EDGAR with zero new Graph list calls. Temporary chat,
document/version, workflow and object were deleted. Output contained only
counts/booleans, not mail subjects/bodies, answers or credentials.

## Full mail metadata and attachments — 2026-09-07

The main Assistant now reads the standard v1.0 message properties, including
From/Sender, To/Cc/Bcc/Reply-To, timestamps, read/draft/receipt flags, importance,
categories, follow-up flag, message/thread/folder identifiers, Internet headers,
body and uniqueBody. Missing values are distinct from explicitly empty lists.
The previous read projection omitted recipients; live Graph confirmed two To
recipients on the sent message. Full metadata is included in tool evidence, and
exact scalar citations handle short values and JSON escaping safely.

Every mail read enumerates attachments even if hasAttachments is false (inline
attachments are excluded from that Graph flag). Names, MIME types, sizes, inline
status, type and reading limitations are provided with opaque attachment refs.
The existing isolated parser reads PDF/DOCX/XLSX/PPTX/TXT/CSV/MD up to 10 MiB.
Item attachments expose message/event/contact data; nested attachment bodies are
not recursively read. Reference attachments and unsupported files are listed
without external downloads. No OCR, legacy Office, or arbitrary-format support
is claimed. Listings are capped at 100 with an explicit more flag; the chat
permits 20 parent sources plus 100 attachments within its existing context limit.

Attachment fingerprints bind parent changeKey and attachment metadata. Access
and versions are rechecked around content reads and before save/history display.
History revalidation groups attachment metadata requests by parent. All source
locators and derived content remain in encrypted chat storage; OFF performs no
new attachment-body retrieval. No dependencies, OAuth scopes or SQL changed.

Validation: 421 tests passed in 15 backend files (Microsoft365 suites, three
Microsoft365 route suites, llmSensitive and llmSensitiveSdk); TypeScript build
and git diff --check passed. Includes real synthetic DOCX parsing, bounded
streaming, unsafe paths/pagination, parent changes, removed/changed attachments,
short/escaped metadata citations and invented-recipient rejection.

Backend Docker build/restart completed on RnD; health HTTP 200. Backup:
`/opt/mike-deploy-backups/mail-details-20260907T145455Z`.
A live main-agent request read the actual latest Sent Items message and named
both To recipients in 13,410 ms; encrypted reload passed, plaintext chat_messages
count was zero. The two accessible real messages contain no attachments, so
actual Microsoft attachment download is not yet verified.

A separate disposable process used synthetic Graph transport with a real DOCX,
the live selected Sonnet provider, actual parser and encrypted storage. The
agent named both synthetic recipients and extracted the expected budget in
24,789 ms, downloading the attachment once. Reload preserved its locator,
plaintext count remained zero, and a subsequent OFF turn made zero additional
attachment downloads. Both temporary chats were deleted. Corporate mail values,
answers and credentials were not emitted in probe output.
