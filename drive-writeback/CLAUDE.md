# CLAUDE.md — drive-writeback

Server-specific guidance for working in this folder. Cross-server rules live in the root `CLAUDE.md` and `../REPO-CONVENTIONS.md` — read those first.

## Status

**Phase 0 (validation spikes) in progress.** No Functions host project, resource group, Function App, or Key Vault exist yet — those are Phase 1+ per `docs/active/PRD-drive-write.md` §11's phasing. Right now this folder is just the Graph client library and its E2E spike tests, mirroring how `outlook-writeback`'s own Phase 0 worked (see that server's `docs/archive/prd-create-update.md` §12 Milestones): the spike lives permanently as real library code plus a live-Graph test tier, not a throwaway script.

Live spike findings get written back into `docs/active/PRD-drive-write.md` §11/§12 as they're confirmed — see that doc for which of the six Phase 0 checklist items are still open.

## Runtime

.NET 10 (`net10.0`).

## Build/test commands

Run from the repo root (`sapidus-writeback-mcp.slnx`):

- `dotnet build` — builds `DriveWriteback.Graph` and `DriveWriteback.Graph.Tests`.
- `dotnet test --filter "Category!=E2E"` — currently a no-op (every test in this folder is `Category=E2E` right now; there's no Unit/Integration tier yet since there's no non-Graph logic to test in isolation until Phase 1's tool surface exists). Still the correct command to run in CI — it stays green by matching nothing.
- `dotnet test --filter Category=E2E` — real Microsoft Graph, real OneDrive. Self-skips via `Assert.Ignore` unless the environment variables below are set; requires the real "Drive Writeback MCP" Entra app registration and admin consent (see `DEPLOYMENT.md`, or run `scripts/Register-EntraApp.ps1`). Cannot run in CI:
  - `DRIVE_WRITEBACK_TENANT_ID` / `DRIVE_WRITEBACK_CLIENT_ID` — the Entra app's IDs.

The E2E tier uses `InteractiveBrowserCredential`, which requires the "Drive Writeback MCP" Entra app to be registered as a **public client** with `http://localhost` listed under **Mobile and desktop** redirect URIs (not a web/confidential registration) — see the doc comment on `DriveGraphClient.CreateWithInteractiveBrowserAuth`. Each spike test creates its own uniquely-named file/folder under the signed-in user's OneDrive root and deletes it in a `finally` block, so repeated runs don't accumulate junk.

## Releases and versioning

This server's canonical version lives in `version.txt` (semver, pre-1.0), tracked by its own [release-please](https://github.com/googleapis/release-please) package (own `CHANGELOG.md`, own release PR — see root `release-please-config.json`). No `host.json` version mirror yet, since there's no `host.json` until Phase 1 creates the Functions host project — that gets wired in then, following `outlook-writeback`'s `extra-files` pattern.

## Key points for a future implementer

- Scopes: `Files.ReadWrite.All` and `Sites.Read.All` only (PRD §6) — no `Sites.ReadWrite.All`. If any Phase 0 spike 403s under these two, that's a direct signal §6 needs revisiting.
- `DriveGraphClient` (in `DriveWriteback.Graph/DriveGraphClient.cs`) is this phase's spike artifact, not the final tool surface — it implements exactly what the Phase 0 spikes need: `CreateFolderPathAsync` (mkdir -p), `UploadTextContentAsync`, `GetItemByPathAsync`, `ReplaceTextContentAsync` (If-Match), `DeleteItemAsync`. Every method takes an optional `driveId`, mirroring the PRD §5 addressing model (omitted → signed-in user's own OneDrive).
- The Graph SDK's generated `Me.Drive.Root` only exposes `GetAsync`/`Content` — no `Children`, no `ItemWithPath`. Folder/file operations route through `Drives[driveId].Items["root"]` instead (Graph's documented `"root"` item-id alias), which does support both. Verified by compiling against the real `Microsoft.Graph` 6.2.0 assembly, not assumed — see the doc comment on `DriveGraphClient` itself if this needs revisiting on a future SDK upgrade.
- No Native AOT (decided against — see PRD §10). Full `Microsoft.Graph` SDK, matching `outlook-writeback`.
- `scripts/Register-EntraApp.ps1` is the idempotent, repeatable version of `DEPLOYMENT.md`'s Entra app registration steps — read the doc first, the script exists so you don't have to hand-type the commands on a re-run. Whether `scripts/` stays as an ongoing directory is an open item, to be decided once `DEPLOYMENT.md` has been through a couple of real runs.
