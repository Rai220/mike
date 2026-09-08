# Microsoft 365 implementation verification — 2026-09-07

The protected mail/file web MVP is implemented and its initial version was
deployed on RnD, including migrations 01/02. It remains disabled by default in
example configuration. The checklist in [TODO](../TODO.md) distinguishes code
verification, live evidence, operator configuration, and deferred extensions.
See the [deployment record](microsoft365-rnd-deployment.md) for server facts and
the historical mail-entry follow-up and the completed unified mail rollout. No dependencies were added or installed for
this work.

## Initial separate-chat backend checks — historical

```bash
npm test --prefix backend -- \
  src/lib/microsoft365 \
  src/middleware/microsoft365Body.test.ts \
  src/routes/__tests__/microsoft365.test.ts \
  src/routes/__tests__/microsoft365Chats.test.ts \
  src/lib/__tests__/llmSensitive.test.ts \
  src/lib/__tests__/llmSensitiveSdk.test.ts \
  src/lib/__tests__/llmProviders.test.ts \
  src/lib/__tests__/llmModels.test.ts \
  src/lib/dbq/__tests__ \
  src/lib/__tests__/access.test.ts \
  src/routes/__tests__/auth.test.ts \
  src/__tests__/workerEntrypoint.test.ts
npm run build --prefix backend
node --test backend/scripts/microsoft365-check.test.mjs
```

Result: **311 tests in 19 files passed**, backend build passed, and **7 canary
unit tests passed**. This includes existing provider/auth/access/job regressions,
real synthetic Office/PDF extraction, parser admission/termination limits,
connection/ACL/CAS/retention boundaries, exact-quote validation, and installed-SDK
checks for sensitive logging/telemetry/retry suppression. Graph and LLM responses
in these automated suites are synthetic; separate live probes are recorded below.

## Initial separate-chat SQL checks — historical

```bash
python3 backend/scripts/microsoft365-schema-check.py
```

Result: **84 assertions passed**, both migrations replay successfully, and
HEAD schema + migrations 01/02 has the same fingerprint as the current fresh
schema (**5332 lines**). Tests cover RLS and actual denied browser-role SELECT,
backend-only grants, composite ownership, immutable retention, membership and
connection cascades, OAuth orphan/expiry cleanup, and retention cleanup.

The script used cached PostgreSQL 16 with minimal Supabase auth/role stubs in a
uniquely named Docker container: no published ports, network disabled, tmpfs data,
1 CPU/512 MiB, and automatic deletion. The container was removed. This is real SQL
execution, but does not validate Supabase 17, GoTrue, PostgREST or the deployed
HTTP stack.

## Initial separate-chat frontend/browser checks — historical

```bash
npm test --prefix frontend -- \
  src/app/components/microsoft365 \
  src/app/components/settings/Microsoft365ConnectionSection.test.tsx
npm test --prefix frontend -- src/app/lib/mikeApi.test.ts -t Microsoft365
npm run build --prefix frontend
npm run lint --prefix frontend -- \
  src/app/components/microsoft365 \
  src/app/components/settings/Microsoft365ConnectionSection.tsx \
  src/app/components/settings/Microsoft365ConnectionSection.test.tsx \
  src/app/lib/mikeApi.ts \
  'src/app/(pages)/assistant/microsoft365/page.tsx' \
  'src/app/(pages)/settings/connectors/page.tsx'
```

Result: **28 component/API tests + 4 connection API tests passed**. Frontend
production build, including type checking and the `/assistant/microsoft365` route,
passed. Targeted lint: **0 errors**, one existing `selectedConnector`
`react-hooks/exhaustive-deps` warning on the connectors page.

The complete pre-existing `mikeApi.test.ts` suite has five previously reproduced
`blob.text is not a function` failures on unchanged HEAD. The targeted result
above does not claim the entire frontend suite passes.

`e2e/microsoft365.spec.ts`: **1 Chromium browser test passed** against the built
frontend, covering mail/file selection, preserving refs over search pages,
question/quote/source viewer, original URL, revoked-access clearing, absence of
share/export controls, and mobile overflow. Browser page errors: zero.
Desktop/mobile/source screenshots were inspected. Every API request was mocked
and non-local network requests were blocked.

The root Playwright package was not installed in this checkout. This run reused
an already installed Playwright 1.59.1 runner and cached Chromium via a temporary
isolated config at `/tmp/mike-microsoft365-browser/playwright.config.cjs`, with no
auth setup dependency and empty storage state. It used port 3337; the test server
was stopped afterward. The version declared by the repository was not validated
by that run. With normal root dev dependencies installed, the committed spec can
be run using the repository test runner with `--no-deps`; use the documented local
frontend and avoid real-auth setup for this synthetic spec.

## Final hygiene and live limits

`git diff --check` passed. The local credential file is ignored and mode 600;
no credential values were included in changes or test output. New migrations
are reflected in `backend/schema.sql`.

The [live localhost check](microsoft365-connection-check.md) previously confirmed
HTTP 200 for token/profile/mail and reading one mail body. Tokens were discarded.
Subsequently, the user completed the persistent Mike sign-in through the HTTPS
callback. On RnD, service status returned `available: true`, `status: connected`,
and a real `checkMicrosoft365Connection` returned `mailAccess: true`. A real
synthetic request to the user's selected `claude-sonnet-4-6` through the existing
direct provider returned the expected marker; no corporate content was included.

The following remain unverified: refresh/reconnect, nonempty Graph file Search and
file downloads, two-user corporate ACL acceptance, proxy/APM behavior, backup deletion
and production latency/load. See
the [setup guide](microsoft365-setup.md) for concrete deployment requirements.

## Separate mail entry follow-up — historical, deployed and verified

That follow-up introduced recent-inbox metadata paging, a sidebar entry and
recent-mail button, and an ordinary-chat handoff before the general model/tool
loop. New protected chats load recent mail with no automatic selection or model
call. Ordinary chats receive only a link, not corporate metadata or content.

```bash
npm test --prefix backend -- src/lib/microsoft365/handoff.test.ts src/lib/microsoft365/sources.test.ts src/lib/microsoft365/chat.test.ts src/routes/__tests__/microsoft365Chats.test.ts
npm test --prefix frontend -- src/app/components/microsoft365/Microsoft365ChatPage.test.tsx src/app/components/shared/AppSidebar.microsoft365.test.tsx
npm run lint --prefix frontend -- src/app/components/microsoft365/Microsoft365ChatPage.tsx src/app/components/shared/AppSidebar.tsx src/app/components/shared/AppSidebar.microsoft365.test.tsx src/app/lib/mikeApi.ts
git diff --check
```

Final focused run: **120 backend tests in 4 files, 13 frontend tests in 2 files**
passed. The handoff implementation also passed 45 targeted handoff/streaming
regression tests. Frontend type checking and targeted lint passed. Backend and
frontend Docker production builds passed on RnD; both services were restarted.
The earlier browser run above predates this follow-up; no new browser run is claimed.

Post-restart runtime policy: `claude-sonnet-4-6`, seven-day retention. Health/page
HTTP 200; both integration APIs HTTP 401 without authentication. The exact user
prompt passed through real ordinary `runLLMStream` and returned the protected
chat link and completed stream.

A temporary protected chat listed one real inbox message. A selected-source model
request succeeded with a validated quotation; encrypted history reloaded and the
source viewer reread the original successfully. The temporary chat was deleted.
Only status/counts/booleans were emitted; source content and model output were not
printed. This verifies one account/message through backend domain services and
the deployed provider, not the full browser OAuth flow or two-user ACL acceptance.


## Unified Assistant — deployed and verified mail path

The separate page/handoff has been superseded by the Microsoft 365 toggle in
ordinary Assistant. The same chat ID/SSE route uses an encrypted linked payload;
the permanent protected flag prevents sharing, project moves, plaintext future
messages and title leaks. ON exposes only the three read-only Graph tools;
OFF removes retrieval tools/body reads while metadata ACL checks protect history.

Fresh focused results reported for the current code:

- **352 backend tests in 14 files passed**.
- **57 frontend tests in 8 files passed**.
- **112 real PostgreSQL assertions passed**, with migrations 01–03 and matching
  fresh/upgraded fingerprints (**5396 lines**). New cases cover immutable private
  state, linked history ownership and blocked plaintext/share writes.
- Backend build, frontend TypeScript check and targeted lint passed.

Subsequent focused runs passed **95 source/runner tests**, **117 route/runner
tests**, and **16 frontend hook tests** after stripping client-supplied protected
history. These suites overlap earlier runs and must not be added into one total.
The final broad backend run passed: **352 tests in 14 files**.

Backend/frontend Docker production builds passed on RnD. Migration 03 was applied
successfully with `ON_ERROR_STOP`; both services restarted and health/main
Assistant returned HTTP 200. Unauthenticated API requests returned HTTP 401.

The updated synthetic Chromium test passed (**1 test**): ordinary Assistant chat
ID preserved across toggle ON/OFF and reload, inert corporate text, and mobile
layout. It reused the previously documented cached Playwright/Chromium runner
and isolated config, without adding dependencies. API responses were mocked.

Live deployed server-runner evidence uses one ordinary chat ID throughout:

- ON listed one inbox source without body retrieval, then read it in a follow-up
  and produced a validated quotation.
- OFF handled a refresh request with **zero new body/list/search calls**; only
  metadata ACL/version checks ran.
- Encrypted history reloaded, plaintext `chat_messages` count was zero, and
  deleting the temporary ordinary chat cascaded to the protected payload.
- The configured direct provider remained `claude-sonnet-4-6`, retention seven days.

Real model differences for metadata-title citations and OFF plain-prose answers
were fixed and redeployed. Body-based replies still require validated quotations.
No source subject/body, model output or credentials were printed by the probes.
This proves the single-account mail path through the deployed runner and storage,
not a live browser sign-in or multi-user corporate ACL acceptance.

### Empty file-search parser — deployed and live-verified

The deployed artifact contains the `providerHits` fix. A live Graph file Search
for `test` returned HTTP 200 in 776 ms: `total: 0`, `moreResultsAvailable: false`,
and no `hits` field. Calling the actual deployed `searchMicrosoft365Sources`
succeeded with `items: []`. Final backend and frontend builds/restarts completed
successfully.

Separate intermittent Graph requests reached the 20-second timeout and were
safely reported. Those upstream request failures are distinct from the successful
empty-result parser check; no automatic-retry guarantee is claimed. No matching
file was found for live body download. Nonempty file Search/downloads,
refresh/reconnect, two-user ACL and operational acceptance remain unverified.

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
