# PRD: Microsoft 365 account connection

## Problem

The live localhost canary has verified delegated `User.Read`, `Mail.Read`, profile access and the body of one message (all HTTP 200). Mike still has no persistent, authenticated way to connect a user's corporate Microsoft account. Canary tokens were deliberately not retained.

## Proposed solution

Implement the first integration milestone: an opt-in Microsoft 365 connection in Settings → Connectors, scoped to the signed-in Mike user and one configured Mike organization / Entra tenant. Store credentials encrypted on the backend. Support connect/reconnect, status, an explicit mailbox access check that returns metadata only, and disconnect.

## Non-goals / out of scope for this milestone

Do not expose mail bodies to chat, other tools, the document library, shared projects, exports, or the frontend. Do not add file permissions, Microsoft login to Mike, application mail permissions, external dependencies, background indexing, or remote database migrations. Protected chat/source provenance and approved LLM policy remain prerequisites to enabling mail in chat.

## Technical approach / affected areas

- Backend: dedicated `lib/microsoft365/` Graph/OAuth service and authenticated routes at `/integrations/microsoft365`.
- Tenant-specific Authorization Code + PKCE, scopes `offline_access User.Read Mail.Read`, no ID token flow. The existing Mike session remains the authentication boundary; `/me` resolves the Microsoft account from the exchanged Graph token. Bind state to the Mike user, organization, tenant, app and request session. One-time, expiring state; validate before network requests.
- Operator configuration explicitly enables the feature and binds `MICROSOFT365_ORG_ID` to `MICROSOFT365_TENANT_ID`. Check current org membership on connect, callback, status/check and refresh. Disconnect remains available to the connection owner after membership/config changes.
- Encrypted access/refresh tokens and PKCE state, backend-only tables, fresh-schema + dated migration. Atomic state consumption and refresh lease / conditional writes prevent concurrent refresh and stale callback resurrection after disconnect. Never read credentials through direct browser database access.
- API: `GET /integrations/microsoft365`; `POST /integrations/microsoft365/connect`; `GET /integrations/microsoft365/callback`; `POST /integrations/microsoft365/:connectionId/check`; `DELETE /integrations/microsoft365/:connectionId`.
- Status: `{ available: boolean, connection: null | { id, status: 'pending' | 'connected' | 'reconnect_required', accountLabel: string | null, connectedAt: string | null, lastCheckedAt: string | null } }`. No tokens, plaintext provider errors or message metadata.
- Connect: `{ authorizationUrl, callbackOrigin }`; callback requires the matching Mike session and redirects to `/settings/connectors?microsoft365=connected|error`. UI uses same-tab navigation, avoiding popup lifecycle dependencies.
- Check: `{ ok: true, mailAccess: true, checkedAt: string }` only, using `/me/messages?$top=1&$select=id` and verifying delegated `Mail.Read`. Existing canary is the evidence of body access; this product endpoint does not read message bodies.
- Frontend: reuse settings primitives and safe API error mapping. MFA verification for sensitive connection mutations follows existing connector UX. Hide the feature when unavailable; do not imply chat access is enabled.

## Success metrics / verification

- Targeted backend tests cover owner/org/tenant/session mismatch, state expiry/replay, callback/disconnect and refresh races, safe Graph failures, and no sensitive output. Targeted frontend tests cover unavailable/disconnected/connected states, connect/check/disconnect, error handling and callback notifications.
- Backend build, relevant frontend lint/typecheck and existing auth/connector regressions pass. Validate SQL offline and through a disposable local database if available; report any real-Supabase/schema-drift gap explicitly.
- The original delegated canary succeeds live. Persistent product connection needs a separately registered HTTPS callback and a second login after deployment; do not claim those have run based on the localhost canary.

## Alternatives considered

Reusing the generic MCP URL for Graph was rejected because Graph does not implement MCP. Automatically attaching mail to the existing chat was rejected because persisted/shared answers and exports currently lack Microsoft source ACL semantics.

## Open questions

Corporate approved LLM, history retention/export policy, production callback registration and pilot organization assignment are deployment/chat-integration prerequisites. The PRD is local; no external issue has been published.
