# PRD: Self-Hosted Outlook Mailbox-Write MCP Server

**Owner:** Keith Mendoza
**Status:** Active
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, Claude Code CLI
**Repo location:** `sapidus-writeback-mcp/outlook-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)

---

> **Deploying this yourself?** Every `example.com` in this document is a placeholder for your own Microsoft 365 tenant's domain — swap it in as you go. "You"/"your" throughout refers to whoever operates this deployment.

## 1. Problem Statement

Anthropic's official Microsoft 365 connector handles Outlook and calendar read and search well on a Pro plan, but it has no write capability available to individual plans: write tools are only enabled through **Organization settings > Connectors**, a Team/Enterprise-only surface with no documented self-service path for Free/Pro/Max users. In practice that leaves two gaps in your workflow that the connector can't fill:

1. **Email drafting** — Claude needs to create a short, text-only Outlook draft from chat context, which you review and send yourself.
2. **Calendar management** — Claude needs to create, edit, and (with confirmation) delete events on your calendar.

This server fills exactly those write gaps. Read and search stay with the M365 connector — no duplication. At the volumes involved, Azure hosting cost is effectively $0/month; the trade is engineering/maintenance time rather than a plan upgrade.

## 2. Goals

- Let Claude create an Outlook draft (subject + body, plain text by default, HTML when needed) in your mailbox.
- Let Claude create, update, and delete calendar events on your calendar.
- Require an explicit in-chat confirmation before any calendar deletion.
- Work identically from Claude Desktop, Claude Cowork, and Claude Code CLI.
- Request the minimum Graph permissions needed: `Mail.ReadWrite` and `Calendars.ReadWrite`, nothing else.
- Keep monthly hosting cost negligible (target: under $2/month).
- Stay scoped to write actions the M365 connector can't do — no re-implementing read/search.

## 3. Non-Goals

- **Read and search (email or calendar)** — handled by the existing M365 connector, deliberately not duplicated here.
- **Sending email** — you send manually. No `Mail.Send` scope requested, so sending is not technically possible even by accident.
- **Attachments** — not needed; drafts with attachments are rejected.
- **Multi-user / multi-tenant** — single-user tool scoped to your own mailbox only.
- **Teams, SharePoint, OneDrive** — out of scope entirely. (If any of these need write access later, that's a new server folder with its own Entra app — see section 14 and `../REPO-CONVENTIONS.md` — not new scopes on this app.)
- **High availability / SLA** — best-effort personal-use uptime is fine.

## 4. Users

- **Primary and only user:** you, using your own Microsoft 365 tenant.

## 5. Use Cases

| # | Frequency | Description |
|---|-----------|-------------|
| UC1 | recurring | Working from context already in the conversation (using the M365 connector for any read/search), Claude drafts a short, text-only email. You review it in Outlook and send it yourself. |
| UC2 | ad hoc | Claude creates a calendar event (title, time, location, notes) on your calendar. |
| UC3 | ad hoc | Claude updates an existing calendar event (reschedule, edit details). |
| UC4 | ad hoc | Claude deletes a calendar event — only after you explicitly confirm in chat. |

## 6. Functional Requirements — MCP Tools

### `create_draft`
- **Input:** to address(es), subject, body, optional `isHtml` flag (defaults to `false` — plain text; set `true` to send `body` as HTML, e.g. for a table).
- **Output:** confirmation + Outlook draft ID/link so you can open it in Outlook.
- **Graph call:** `POST /me/messages` (saved as a draft, not sent).
- **Scope required:** `Mail.ReadWrite`.
- **Hard constraints:**
  - Reject any request that includes attachments.
  - No `Mail.Send` scope requested, so the server cannot send — only draft.

### `update_draft`
- **Input:** draft message ID, plus any fields to change (to, subject, body), optional `isHtml` flag (same default/behavior as `create_draft`; only applies when `body` is also being changed).
- **Output:** confirmation + updated draft ID/link.
- **Graph call:** `PATCH /me/messages/{id}`.
- **Scope required:** `Mail.ReadWrite` (already in scope for `create_draft`).
- **Hard constraint:** reject any update that adds attachments.
- **ID source:** the draft ID can come from this server's own `create_draft` response earlier in the conversation, or from the M365 connector's read tools searching the Drafts folder. Verified in Phase 0 testing: drafts share the same `/me/messages` ID space as regular mail, so an ID the connector finds by searching the Drafts folder is directly usable here — no translation needed.
- **Not confirmation-gated:** unlike `delete_event`, editing a draft isn't destructive — nothing is lost, you just see the updated draft in Outlook.

### `create_event`
- **Input:** subject, start (ISO 8601 + timezone), end, optional location, optional body/notes, optional attendees.
- **Output:** confirmation + event ID/link.
- **Graph call:** `POST /me/events`.
- **Scope required:** `Calendars.ReadWrite`.

### `update_event`
- **Input:** event ID, plus any fields to change (subject, start, end, location, body, attendees).
- **Output:** confirmation + updated event ID/link.
- **Graph call:** `PATCH /me/events/{id}`.
- **Scope required:** `Calendars.ReadWrite`.

### `delete_event`
- **Input:** event ID.
- **Behavior:** **two-step, confirmation-gated.** The tool must not delete on first call. On the first invocation it returns the event's details (subject, date/time) and a confirmation token, and does nothing destructive. It performs the actual `DELETE` only on a second call that echoes back that token — ensuring you've confirmed the specific event in chat.
- **Graph call:** `DELETE /me/events/{id}` (only on confirmed second call).
- **Scope required:** `Calendars.ReadWrite`.
- **Rationale:** delete is the only destructive, hard-to-notice action in the set; create/update are self-evident on the calendar, delete is not.

**Note on event IDs:** since this server doesn't read/search, event IDs for `update_event` and `delete_event` come from the M365 connector's calendar search in the same conversation. Verified in Phase 0 testing: an event ID returned by the connector's calendar search resolves via this app's own Graph credentials against `GET /me/events/{id}` — confirming both hit the same `/me/events` ID space, no translation needed.

## 7. Authentication & Authorization

Two independent auth boundaries — don't conflate them.

### 7a. MCP server -> Microsoft Graph (mailbox + calendar writes)

**Reusing Anthropic's existing Enterprise App — investigated, not viable. Confirmed in Phase 0.** Admin consent for the official M365 connector only instantiates the *service principal* (Enterprise Application object) for "M365 MCP Client for Claude" / "M365 MCP Server for Claude" inside your tenant. The underlying **App Registration** — and the client secret/certificate needed to mint tokens as that app — is owned and held by Anthropic, not by tenant admins. There's no supported mechanism for a third-party server to authenticate using someone else's registered client credentials. Verified live with a throwaway `http://localhost` redirect against both service principals' client IDs, each failing a different way: one returned **AADSTS50011** (redirect URI not registered for that app), the other **AADSTS700016** (no service principal for that app ID in this tenant at all). Both confirm this server cannot mint delegated Graph tokens using Anthropic's app IDs.

**Path forward: the MCP server registers and owns its own Entra App.**
- **Display name: "Outlook Writeback MCP."** Deliberately does *not* include "Claude" — the tenant already has Anthropic's own "M365 MCP Client for Claude" / "M365 MCP Server for Claude" service principals in Enterprise Applications, and a third "Claude"-named app owned by this project would be both confusing to distinguish at a glance and misleading about affiliation, since this app isn't an Anthropic product. If you want to note which client calls it, that goes in the app's description/publisher-notes field (free text), not the searchable display name.
- Single-tenant App Registration, restricted to your own tenant only.
- Delegated permissions: `Mail.ReadWrite` and `Calendars.ReadWrite` — the only two Graph scopes needed.
- Explicitly do **not** request `Mail.Send`, `Mail.Read`, `Calendars.Read` (read is the connector's job), `Files.*`, or any Teams/SharePoint scope.
- Admin consent granted once by you (as Global Admin of the tenant).
- OAuth 2.0 Authorization Code flow with PKCE; refresh token cached in Azure Key Vault, never in app storage or logs.

### 7b. Claude -> MCP server (who's allowed to call the tools)

A separate, smaller problem: only you will ever call this server, so it needs no multi-user OAuth infrastructure. Two options, in order of preference:

1. **Static bearer token.** Claude Code CLI accepts a custom `Authorization` header directly (`claude mcp add --transport http ... --header "Authorization: Bearer <token>"`) — no OAuth dance. Simplest to build and maintain.
2. **Minimal self-issued OAuth 2.1 server.** If the Desktop/Cowork custom connector UI requires an OAuth flow rather than a static header (Phase 0 check — see open questions), implement a small Authorization Code + PKCE endpoint in the same Function app, backed by a single hardcoded user record. Satisfies "the MCP server does its own authentication" without a general-purpose multi-tenant provider.

The token protecting boundary 7b is unrelated to, and must not be confused with, the Graph credentials in 7a. Any transport/auth code that implements this boundary is the one thing in this server that's a candidate for the monorepo's shared `shared/` folder (see `../REPO-CONVENTIONS.md`) — the 7a Graph credentials never are.

## 8. Architecture Overview

```
Claude (Desktop / Cowork / Code CLI)
        |  HTTPS, MCP Streamable HTTP transport
        |  (bearer token or OAuth — see 7b)
        v
Azure Function App: <server>-func
  (resource group: <server>-rg)
        |  MCP tool handlers: create_draft, create_event,
        |                     update_event, delete_event (confirm-gated)
        |  OAuth 2.0 Authorization Code + PKCE to Graph (see 7a)
        v
Microsoft Graph API  (/me/messages, /me/events)
        v
your Outlook mailbox + calendar (your tenant)
```

- **Compute:** Azure Functions, C# (isolated worker model), HTTP trigger, Consumption (or Flex Consumption) plan.
- **Resource naming:** this server's Azure resources are named to trace directly back to the repo folder — resource group `<server>-rg`, Function App `<server>-func`, Entra app "Outlook Writeback MCP." Every future server folder in the monorepo gets its own resource group and Function App under the same `<server>-rg` / `<server>-func` pattern (see `../REPO-CONVENTIONS.md`) — resources are never shared across servers, even though the code lives in one repo.
- **Secrets:** Azure Key Vault holds the Graph app's client secret/certificate and the Claude-facing bearer token or OAuth signing key. The Function App never stores these directly:
  - The Function App authenticates to Key Vault via **Managed Identity** — no stored bootstrap credential.
  - No `local.settings.json` file is used in production. `local.settings.json` is local-dev-only, is listed in `.gitignore`, and a `local.settings.json.example` with placeholder keys ships in the repo for contributors. Deployment reads Application Settings, never a file in the repo.
- **State:** minimal — the cached Graph refresh token lives in **Azure Key Vault** (§7a), read and written live via the Key Vault Secrets SDK through the Function App's Managed Identity, not an app-setting Key Vault reference — the app must be able to write the rotated token back after every redemption, and app-setting references are read-only. (An earlier draft of this section said Table Storage; Key Vault is correct, since this is a secret.) **Resolved in Phase 2:** `delete_event`'s pending-confirmation state needs no storage at all. `DeleteConfirmationTokenService` (`OutlookWriteback.Graph/Confirmation/`) issues a stateless, HMAC-signed token carrying the event ID and a 5-minute expiry; the first `delete_event` call mints one, the second call must echo it back, and `EventDeletionService` only calls Graph's `DELETE` once that token validates for the same event ID. Unforgeability is what makes the gate real — a client can't compute its own token and skip the round-trip. The signing key is a static secret (no rotation/write-back needed), so unlike the refresh token it's provisioned as a plain Key Vault-reference app setting rather than read live via the Secrets SDK — see `DEPLOYMENT.md`.
- **Network:** must be a public HTTPS endpoint. Claude Desktop, Cowork, and claude.ai connect to remote MCP servers from Anthropic's cloud infrastructure, not from your local device — a private/VPN-only endpoint won't work for those clients. A custom domain (`outlook-writeback.example.com`, following a flat `<server>.example.com` per-server naming convention — see `DEPLOYMENT.md`'s Phase 4 section) is live and, since Phase 4's cutover, is the **only** hostname that supports a fresh OAuth authorize for this server — `<server>-func.azurewebsites.net` still serves traffic but no longer works for OAuth clients.

## 9. Client Setup (per surface)

- **Claude Code CLI:** `claude mcp add --transport http outlook-write https://outlook-writeback.example.com/mcp --header "Authorization: Bearer <token>"` (or `/mcp` to run the OAuth flow, if that path is used instead).
- **Claude Desktop / Cowork:** Customize > Connectors > "+ Add custom connector" -> server URL -> Advanced settings for OAuth Client ID/Secret if the OAuth path is used.

## 10. Non-Functional Requirements

- **Cost:** negligible. At single-user volume this sits deep inside Azure Functions' free grant; set a low Cost Management budget alert on `<server>-rg` as a tripwire (notification-only — it cannot throttle or disable the Function App, so it never risks blocking a legitimate call mid-month). Tag this resource group (and every future server's resource group) with a common tag — e.g. `project: sapidus-writeback-mcp` — so cost is still visible as one rollup in Cost Management despite each server living in its own resource group (see `../REPO-CONVENTIONS.md`).
- **Security:** two Graph scopes only (`Mail.ReadWrite`, `Calendars.ReadWrite`); no send capability; secrets never logged; drafts with attachments rejected; calendar deletes confirmation-gated; single-tenant app registration.
- **Reliability:** best-effort; no uptime SLA needed for a personal tool.
- **Logging:** log tool invocations (name, timestamp, success/failure) for personal debugging — never email or event body content.

## 11. Open Questions

1. ~~Does the Desktop/Cowork custom-connector UI accept a static bearer token, or does it require a full OAuth handshake?~~ **Resolved.** Observation already showed no raw Authorization-header field in the "Add custom connector" dialog — only a server URL plus optional **OAuth Client ID/Secret**. Phase 3's spike (throwaway `<server>-spike-func`/`<server>-spike-rg`, isolated from production per `../REPO-CONVENTIONS.md`'s per-server invariant) supplied the live round-trip this question was waiting on, and changed the answer to §7b option 2 itself: **Azure Functions' built-in Easy Auth (App Service Authentication v2, `authsettingsV2`) fronted by a new Entra ID app satisfies the full MCP OAuth contract with zero hand-rolled `/authorize`/`/token` code** — cheaper than the minimal self-issued OAuth 2.1 server this section originally called for. Confirmed live, against the spike app:
   - `GET /.well-known/oauth-protected-resource[/runtime/webhooks/mcp]` is served **natively by Easy Auth v2** — no custom RFC 9728 shim endpoint needed.
   - An unauthenticated request returns `401` with a correctly-formed `WWW-Authenticate: Bearer ... resource_metadata="https://.../.well-known/oauth-protected-resource/runtime/webhooks/mcp"` header.
   - A request bearing a valid Entra-issued token for the new app succeeds (`200`, full MCP `initialize` response) — but only once the MCP extension's own separate `mcp_extension` system-key gate is relaxed via `host.json`'s `extensions.mcp.system.webhookAuthorizationLevel: "Anonymous"`. The app-setting equivalent (`AzureFunctionsJobHost__extensions__mcp__system__webhookAuthorizationLevel`) was tried first and confirmed persisted, but had no observable effect even after a full restart — only the `host.json` literal, republished, actually changed behavior.
   - **Coexistence finding:** the system-key layer and Easy Auth's platform layer don't compose — once the system-key gate is set to `Anonymous` so Easy-Auth-validated tokens get through, it stops providing any protection of its own, so the existing `x-functions-key` path Claude Code CLI uses today would need to keep working through a gate that's now wide open, which defeats its purpose. **Decision:** migrate Claude Code CLI to the same Entra OAuth flow as Desktop/Cowork (Entra has no DCR/CIMD, so the CLI registers with a manually-supplied `--client-id`/`--client-secret`/`--callback-port` per `claude mcp add`), dropping `x-functions-key` entirely. Unifies all three clients on one auth mechanism — also the more agent-agnostic outcome per `REPO-CONVENTIONS.md` §9, since nothing about the flow is Claude-specific (Entra ID is a general-purpose OAuth 2.1 AS; any conforming client can use the same registration).

   Full command-level findings (exact `az ad app`/`az rest` sequence, `authsettingsV2` payload, validation checks) are in `DEPLOYMENT.md`. **Resolved — applied to production and validated.** The spike's Entra app was repointed to production-only (`identifierUris` stripped of the spike hostname) rather than replaced. Two things the spike didn't catch surfaced during the production rollout: Easy Auth's `WWW-Authenticate` header only advertises `scope`/`resource_metadata` (the RFC 9728 discovery fields) once the undocumented `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` app setting is present — without it, auth still works but discovery is silently degraded; and Claude Desktop/Cowork's redirect URI (`https://claude.ai/api/mcp/auth_callback`) does need to be registered (under `web.redirectUris`, since Desktop is a confidential client, unlike the CLI's public/PKCE registration) — the original assumption that Anthropic's client needed no redirect URI registration was untested and wrong. Both are now in `DEPLOYMENT.md`. Live validation passed against both real Claude Code CLI (OAuth migration, a real `create_draft` call) and real Claude Desktop (custom connector, live sign-in).
2. ~~Are event IDs returned by the M365 connector's calendar search directly usable as Graph `/me/events/{id}` identifiers by this server?~~ **Resolved in Phase 0:** yes — confirmed live against a real connector-returned event ID (see §6 note on event IDs).

## 12. Milestones

- **Phase 0 — Spike (½–1 day): complete.** ~~Register your own Entra app ("Outlook Writeback MCP") with `Mail.ReadWrite` + `Calendars.ReadWrite` and confirm consent~~ done. ~~Prototype `POST /me/messages` and `POST /me/events` locally~~ done — live E2E checks passed (draft created, event created). ~~Verify connector event IDs work here (Q2)~~ done — resolved yes. ~~Confirm Anthropic's Enterprise App can't be reused (throwaway token request)~~ done — AADSTS50011/AADSTS700016 (see §7a).
- **Phase 1 — Email MVP: complete.** ~~Deploy `create_draft` and `update_draft` to Azure~~ done — `<server>-func` on Flex Consumption, using the official GA Azure Functions MCP extension (`Microsoft.Azure.Functions.Worker.Extensions.Mcp` 1.4.0), which didn't exist when this PRD's §8 architecture was drafted and supersedes its original hand-rolled-HTTP-trigger sketch. ~~Wire up to Claude Code CLI (simplest auth path)~~ done — the extension's built-in `mcp_extension` system key over `x-functions-key` satisfies §7b with no custom bearer/OAuth code. ~~Validate by drafting and revising a real reconciliation email~~ done — both tools called from real, separate Claude Code CLI invocations against the deployed endpoint. Also solved along the way: non-interactive Graph auth for the deployed (non-browser) service, via a new `SilentGraphCredential`/`GraphTokenEndpointClient` pair that silently redeems a refresh token cached in Key Vault (read and write-back, since Entra rotates it on every redemption), seeded once via a new `OutlookWriteback.Bootstrap` console tool. See `DEPLOYMENT.md` for the full runbook.
- **Phase 2 — Calendar: complete.** ~~Add `create_event`, `update_event`, and the confirm-gated `delete_event`~~ done — `create_event` and `update_event` mirror the draft tools' shape (`OutlookGraphClient.CreateEventAsync`/`UpdateEventAsync`, both now also supporting an optional attendee list); `update_event` follows `update_draft`'s partial-update pattern (only supplied fields change) and replaces the attendee list entirely rather than merging, per Graph PATCH semantics. ~~Validate the delete two-step flow specifically~~ done, via a stateless design rather than a storage service: `DeleteConfirmationTokenService` mints an HMAC-signed, event-ID-bound, 5-minute-TTL token on the first `delete_event` call, and `EventDeletionService` only calls `DELETE /me/events/{id}` when a second call's token validates for that same event ID — covered by unit tests on the token service (valid round-trip, wrong event ID, tampered/wrong-key signature, expired, malformed) and an integration test proving the Graph client is never invoked when the token fails validation. See §8's State bullet for the storage-question resolution.
- **Phase 3 — Multi-client: complete.** ~~Determine whether the Desktop/Cowork custom connector's OAuth-only requirement can be met without hand-rolling an Authorization Code + PKCE server~~ done — validated against a throwaway `<server>-spike-func` (see §11 Q1): Azure Functions' built-in Easy Auth fronted by a new Entra app satisfies the full MCP OAuth contract, including the RFC 9728 protected-resource-metadata document, with no custom endpoint code. ~~Apply this configuration to production, migrate the CLI, validate against real Claude Desktop/Cowork~~ done — production `<server>-func` now runs Easy Auth via the repointed "Outlook Writeback MCP Connector" app, Claude Code CLI is migrated off `x-functions-key` onto the same Entra OAuth flow (public-client/PKCE registration, no client secret), and both the CLI and real Claude Desktop were validated live against production. Two gaps the spike hadn't surfaced were found and fixed during production rollout: the `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` app setting (required for RFC 9728 discovery fields in the `401`'s `WWW-Authenticate` header) and Desktop/Cowork's redirect URI needing explicit registration under `web.redirectUris`. Full detail in `DEPLOYMENT.md`. The spike resource group was deleted once production validation passed.
- **Phase 4 — Hardening:** Add cost/budget alerts (Azure Cost Management + Billing → Budgets, scoped to `<server>-rg`; a plain budget alert emails a threshold breach directly — no Application Insights or Action Group needed for that alone) — **done.** A low monthly budget was set on `<server>-rg` via the portal; it's notification-only and cannot throttle or disable the Function App, so it can't interrupt a legitimate call mid-month even at a tight threshold. Tidy logging; optional custom domain — **done, as a cutover.** `outlook-writeback.example.com` (a flat `<server>.example.com` naming convention, chosen over a shared `mcp.example.com` label that would collide once future servers like a hypothetical onedrive-writeback need their own) is live via a CNAME + TXT pair added at the domain's DNS provider and Flex Consumption's site-scoped App Service Managed Certificate (currently in preview), and `outlook-write` connects through it with a real, cold OAuth authorize — not just a cached-token reconnect. Getting there took two wrong attempts and one right one on this self-referencing Entra app ("Outlook Writeback MCP Connector"): adding the new hostname's `identifierUris` alongside the old ones broke the existing connection with `AADSTS90009` (ambiguous self-reference across 4 candidate URIs); leaving `identifierUris` and the Function App's `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` setting alone passed discovery checks but failed a fresh authorize with `AADSTS9010010` (the PRM document's dynamic `resource` field and the static advertised `scope` disagreed). The fix that worked: **swap**, not add — `identifierUris` replaced with exactly the new hostname's two forms, and `WEBSITE_AUTH_PRM_DEFAULT_WITH_SCOPES` updated to match, in the same change. That resolves the mismatch but is a full cutover: `<server>-func.azurewebsites.net` no longer supports a fresh OAuth authorize as a result — an accepted tradeoff since there was exactly one client to repoint. Full runbook and all three findings in `DEPLOYMENT.md`'s Phase 4 section.

## 13. Known Limitations

### 13a. Claude-side OAuth refresh gap for external-IdP custom connectors (boundary 7b)

Per [Anthropic's connector authentication docs](https://claude.com/docs/connectors/building/authentication), Claude's MCP client is documented to refresh OAuth tokens "reactively on a 401 response, with a proactive refresh up to five minutes before the stored expiry." In practice this doesn't happen for this server, because of an open, unresolved bug: **[anthropics/claude-ai-mcp#228](https://github.com/anthropics/claude-ai-mcp/issues/228)**, filed against custom connectors backed by an external IdP — Entra ID is the exact case the issue cites. Affected connectors never get either the reactive or proactive refresh attempted: the client-side proxy reconnects the transport and reports "Connected," but forwards the now-expired access token on the next tool call anyway, which fails. As of this writing there's no confirmed fix or timeline from Anthropic.

This lands squarely on `outlook-writeback`: boundary 7b runs on Easy Auth + Entra ID (§7b, §11 Q1), which is the affected external-IdP shape, Entra access tokens default to roughly a one-hour lifetime, and this server's real usage is roughly weekly (UC1, §5). Almost every invocation after a multi-day gap is therefore likely to hit an expired token that the documented refresh behavior should have — but currently doesn't — prevent.

**Why `static_headers` auth isn't a workaround here.** That auth type would sidestep the whole OAuth-expiry problem, since there'd be no token to refresh. But the "Add custom connector" dialog in Claude Desktop, Cowork, and claude.ai only exposes Name, Remote MCP server URL, and an optional OAuth Client ID/Secret pair — there is no raw Authorization-header field through which a static bearer token could be supplied. Until that UI adds one, `static_headers` isn't reachable from these clients regardless of what the server itself supports. Revisit this section if that changes.

**Mitigation adopted (does not fix the underlying bug — that's outside this project's control):**

- **(a) Tool-level guidance.** Every tool's MCP description (`Functions/CreateDraftTool.cs`, `UpdateDraftTool.cs`, `CreateEventTool.cs`, `UpdateEventTool.cs`, `DeleteEventTool.cs`) now tells the calling model that an authentication/401-style failure means it should tell the user the connector likely needs reconnecting, rather than silently retrying or failing the call.
- **(b) Open action item.** Check whether the "Outlook Writeback MCP Connector" Entra app's Conditional Access policy or a custom token-lifetime policy can extend the default ~1-hour access-token lifetime. A longer-lived token shrinks how often a refresh is needed in the first place, which reduces exposure to #228 — it does not close the underlying gap, since any refresh that does eventually become necessary would still hit the same bug.

## 14. Future Consideration: Auth Pattern for Additional MCP Servers

**Decided:** one Entra App Registration per server, kept as an invariant, not a starting point to revisit later. When a new write capability is needed (OneDrive, Tasks, etc.), it gets its own folder in the `sapidus-writeback-mcp` monorepo, its own Entra App Registration with only the scopes it needs, its own resource group, and its own Function App — never new scopes bolted onto this app's registration. This keeps section 2's "minimum permissions" goal and section 10's "two scopes only" posture true as an ongoing property of this server, not just a fact about how it launched.

Repo structure is a separate axis from this and doesn't change the decision: living in a monorepo (`sapidus-writeback-mcp/outlook-writeback/`, `sapidus-writeback-mcp/onedrive-writeback/`, etc.) is purely a code-organization convenience. It shares CI/tooling and a small amount of genuinely capability-agnostic code (the 7b Claude-to-server transport/auth layer, in `sapidus-writeback-mcp/shared/`) — it never shares Entra apps, resource groups, Key Vault secrets, or Graph scopes across server folders. See `../REPO-CONVENTIONS.md` for the full naming and isolation rules every server in the monorepo follows.

The alternative considered and rejected — a shared "Client" + per-server "Server" apps via On-Behalf-Of, mirroring Anthropic's own `M365 MCP Client for Claude` / `M365 MCP Server for Claude` pair (section 7a) — remains available to revisit, but only if maintaining N independent OAuth flows across servers becomes an actual, felt friction point, not as a default direction.
