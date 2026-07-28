# CLAUDE.md — outlook-writeback

Server-specific guidance for working in this folder. Cross-server rules live in the root `CLAUDE.md` and `../REPO-CONVENTIONS.md` — read those first.

## Status

All phases (spike, email MVP, calendar, multi-client OAuth, custom domain + cost hardening) are complete. Full implementation history and design rationale live in `docs/archive/` (completed/superseded specs) and `DEPLOYMENT.md` (the operational runbook) — not duplicated here. In-progress specs, if any, live in `docs/active/`.

`docs/active/prd-multi-recipient.md` is implemented and merged in code but not yet deployed to the production Function App — see the root `CLAUDE.md` PRD-lifecycle rule (`## Document maintenance`) for why it hasn't moved to `docs/archive/` yet.

## Runtime

.NET 10 (`net10.0`), isolated worker model.

## Build/test commands

Run from the repo root (`sapidus-writeback-mcp.slnx`):

- `dotnet build` — builds all four projects (`OutlookWriteback.Graph`, `OutlookWriteback.Graph.Tests`, `OutlookWriteback` (the Functions app), `OutlookWriteback.Bootstrap`).
- `dotnet test` — runs all three NUnit tiers:
  - `Category=Unit` — pure payload-mapping logic and the non-interactive auth machinery (`GraphTokenEndpointClient`, `SilentGraphCredential`), no I/O.
  - `Category=Integration` (`OutlookGraphClientIntegrationTests`) — exercises `OutlookGraphClient`'s real request-building/response-deserialization through the Graph SDK against a stubbed `HttpMessageHandler` (`TestSupport/StubHttpMessageHandler.cs`). No network, no credentials, fast (~1s total) — this is the tier for the normal dev loop.
  - `Category=E2E` (`OutlookGraphClientE2ETests`) — real Microsoft Graph, real mailbox. Self-skips via `Assert.Ignore` unless the environment variables below are set; requires the real "Outlook Writeback MCP" Entra app registration and admin consent from the user. Cannot run in CI:
    - `OUTLOOK_WRITEBACK_TENANT_ID` / `OUTLOOK_WRITEBACK_CLIENT_ID` — the Entra app's IDs.
    - `OUTLOOK_WRITEBACK_TEST_TO_ADDRESS` — a real mailbox address, for the draft-creation check.
- `dotnet test --filter "Category!=E2E"` — Unit + Integration only, safe and fast for CI.
- `func start` (from this folder) — local smoke test against real Azure dependencies (Key Vault, Entra app) via `local.settings.json` + your own `az login` session; prints the discovered MCP tool list on startup, a much faster feedback loop than deploy-and-poll.

The E2E tier uses `InteractiveBrowserCredential`, which requires the "Outlook Writeback MCP" Entra app to be registered as a **public client** with `http://localhost` listed under **Mobile and desktop** redirect URIs (not a web/confidential registration) — see the doc comment on `OutlookGraphClient.CreateWithInteractiveBrowserAuth`. This tier still only hits the library directly, not the deployed Azure endpoint — extending it to a deployed-E2E tier (speaking real MCP Streamable HTTP to the live endpoint) is a reasonable stretch item, not yet built.

## Releases and versioning

This server's canonical version lives in `version.txt` (semver, pre-1.0). [release-please](https://github.com/googleapis/release-please) mirrors it automatically into `host.json`'s `extensions.mcp.serverVersion` — the MCP protocol field clients read — on every release; don't hand-edit `serverVersion` directly, it'll be overwritten by the next release PR.

release-please watches Conventional Commit messages on `main` and keeps an up-to-date release PR open scoped to this folder (`outlook-writeback/CHANGELOG.md` + `version.txt` + the `host.json` mirror — see root `release-please-config.json`). Merging that PR cuts an `outlook-writeback-vX.Y.Z` tag and GitHub Release. Pre-1.0, breaking changes bump minor, not major (`bump-minor-pre-major`, set repo-wide).

The release PR is created with the default `GITHUB_TOKEN`, which doesn't itself trigger further Actions workflows on the commits/tags it creates — if a future build/deploy-on-tag workflow needs to fire automatically off a release-please tag, it'll need a PAT or GitHub App token instead.

## Key points for a future implementer

- Scopes: `Mail.ReadWrite` and `Calendars.ReadWrite` only. No `Mail.Send` — the server can create/update drafts but structurally cannot send.
- Tools: `create_draft`, `update_draft`, `create_event`, `update_event`, `delete_event`.
- `delete_event` is the only destructive tool and is two-step, confirmation-gated: the first call returns event details + a confirmation token and does nothing destructive; the actual `DELETE` only happens on a second call that echoes that token back.
- Drafts with attachments are rejected outright (no attachment support).
- Event/draft IDs are expected to come from the M365 connector's read/search tools in the same conversation (shared Graph ID space) — this server never implements its own read/search.
- Compute: Azure Functions, C# (isolated worker model), HTTP trigger, Flex Consumption plan.
