# PRD: Microsoft 365 toggle in the main Assistant

## Outcome

A connected user stays in the ordinary Assistant, enables Microsoft 365 beside
existing capabilities such as EDGAR, and asks “Какие письма у меня на почте?”.
Mike lists recent inbox subjects and can read/search the user's mail and accessible
corporate files in subsequent turns. There is no separate mail assistant, source
picker or handoff. The chat ID and normal SSE transport remain the same.

## Toggle contract

- ON: expose recent-mail, search and source-read Graph tools to Mike. Read an
  original before answering about its contents; validate verbatim quotations.
- OFF: remove retrieval tools and automatic new body reads. Discuss already-shown
  history only after metadata ACL/version checks. Do not claim fresh inbox access.
- Once used, the conversation stays private and encrypted in both states. OFF
  does not enable sharing, project moves, exports, MCP, EDGAR/web or document tools.
- Accounts connect/reconnect in Settings → Connectors. Legacy
  `/assistant/microsoft365` redirects to `/assistant`; no extra sidebar item.

## Implementation boundaries

Reuse `microsoft365_chats` as an encrypted payload linked by `ordinary_chat_id`.
Migration 03 permanently protects the ordinary shell and guards plaintext writes,
sharing, project/organization linkage and titles. Retain existing conversation
context by copying ordinary history once, then store every later turn encrypted.
Never put source subjects or generated summaries into the ordinary title.

Reuse Mike's prompt and the existing direct provider in sensitive mode, with an
explicit Graph-only tool allowlist. Buffer output until post-generation source
checks and encrypted persistence succeed, then emit the normal SSE response.
See the [protected storage/execution PRD](microsoft365-protected-chat-prd.md).

## Acceptance

1. The exact inbox question invokes recent mail with ON in the main Assistant;
   a follow-up reads a previously listed message without manual source selection.
2. OFF exposes no source retrieval tools or automatic body reads; history context
   survives and ACL checks still fail closed on revoked/changed sources.
3. Reload hydrates the same chat's encrypted history; new corporate and OFF turns
   never enter ordinary plaintext messages or list titles.
4. Sharing/move/export and direct plaintext/share writes remain blocked, including
   concurrent promotion attempts and after disconnect/deletion/expiry.
5. Provider/logging/tool boundaries, migration replay, UI toggle behavior and
   normal non-Microsoft chat regressions pass focused tests.
6. Deploy migration 03 and rebuilt services on RnD, then verify the live main-chat
   route and ON/OFF behavior. **Completed for the single-account mail path**:
   same-chat listing/body read, OFF without new body/list/search calls, encrypted
   reload, zero ordinary plaintext messages and temporary-chat cascade cleanup.

The existing RnD policy is direct `claude-sonnet-4-6` with seven-day protected
retention. A real file Search returned HTTP 200 with no hits and the deployed
adapter returned an empty list successfully. Separate intermittent 20-second
Graph timeouts were safely reported. Nonempty file access/downloads,
refresh/reconnect, two-user ACL exercises and
backup/provider operational guarantees are not established by mail-only tests.

## Amendment: independent composer options (2026-09-07)

Microsoft 365 ON/OFF controls only Graph retrieval. Existing attachment, EDGAR,
workflow, model and reasoning selections remain usable and survive reload. The
main protected runner composes explicit read tools while retaining encrypted
history, current source authorization and version checks. User-selected models
use authenticated server-loaded user settings; the configured model is a fallback.
The prior OFF-no-tools and fixed-model acceptance statements are superseded for
the ordinary Assistant. Workflow instructions can run and produce drafted text;
document mutation/artifact generation and arbitrary MCP remain outside this
protected runner. See the updated verification protocol for live evidence.
