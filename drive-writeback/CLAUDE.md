# CLAUDE.md — drive-writeback

Server-specific guidance for working in this folder. Cross-server rules live in the root `CLAUDE.md` and `../REPO-CONVENTIONS.md` — read those first.

## Status

**Phase 1 ("Safe creates") implemented — pending deployment.** `get_item`, `create_folder`, and `create_file` are built end-to-end: `DriveGraphClient` (Graph layer), `DriveWriteService` (dry-run-aware service layer, dry-run **on** by default), and their `Functions/` tool classes, all wired into a real Functions host project (`Program.cs`, `host.json`) plus a `DriveWriteback.Bootstrap` console app for one-time refresh-token seeding. `dotnet test --filter "Category!=E2E"` is green (45 Unit/Integration tests, all offline, no live tenant needed).

**Nothing has actually been deployed.** No resource group, Function App, or Key Vault exists yet — `DEPLOYMENT.md`'s Phase 1 sections are runbook text for a future session to execute, not something this pass ran. Per this repo's doc-maintenance convention, this Status section stays "pending deployment" (not "shipped") until that actually happens.

**The `@microsoft.graph.conflictBehavior`-on-content-PUT question is unresolved.** `create_file`'s `conflict_behavior` handling does not depend on the answer — it enforces `fail`/`rename`/`replace` entirely client-side (pre-check via `TryGetItemByPathAsync`, no reliance on Graph's query parameter) — but the open question itself is still open. A probe test exists (`DriveGraphClientE2ETests.ContentPut_conflictBehavior_query_parameter_is_an_open_question`) and has never been run against a live tenant. Two more E2E tests are similarly unrun: `DriveWriteServiceE2ETests.ContentPut_to_a_missing_parent_is_an_open_question` (does a content PUT to a missing parent auto-vivify it or 404 — moot for this server's own behavior since `DriveGraphClient.CreateFileAsync` guards against it before ever reaching Graph, but worth recording) and the live `conflict_behavior`/dry-run-zero-mutations checks in that same fixture. Run `dotnet test --filter Category=E2E` against a real tenant and write the findings back into `docs/active/PRD-drive-write.md` §11/§12 once someone does.

Phase 0's original three OneDrive spikes (mkdir-p, eTag/cTag `If-Match`, item-ID shape) remain confirmed passing from that earlier session — see `docs/active/PRD-drive-write.md` §11 for the historical record. SharePoint validation is still deferred, no team site available yet; Phase 1 is OneDrive-only.

## Runtime

.NET 10 (`net10.0`).

## Build/test commands

Run from the repo root (`sapidus-writeback-mcp.slnx`):

- `dotnet build` — builds the whole solution: `DriveWriteback.Graph`, `DriveWriteback.Graph.Tests`, `DriveWriteback` (the Functions host), and `DriveWriteback.Bootstrap`.
- `dotnet test --filter "Category!=E2E"` — 45 tests, Unit + Integration tiers, entirely offline, no credentials needed. This is the tier that runs in CI.
- `dotnet test --filter Category=E2E` — 7 tests, real Microsoft Graph, real OneDrive. Self-skips via `Assert.Ignore` unless the environment variables below are set; requires the real "Drive Writeback MCP" Entra app registration and admin consent (see `DEPLOYMENT.md` — the sole source of truth for provisioning this, no separate script). Cannot run in CI:
  - `DRIVE_WRITEBACK_TENANT_ID` / `DRIVE_WRITEBACK_CLIENT_ID` — the Entra app's IDs.
  - **Only 3 of the 7 have ever actually been run against a live tenant** (the original Phase 0 spikes). The other 4, all added this pass, are unrun — see the Status section above.

The E2E tier uses `InteractiveBrowserCredential`, which requires the "Drive Writeback MCP" Entra app to be registered as a **public client** with `http://localhost` (E2E tests) and `http://localhost:8500/` (`DriveWriteback.Bootstrap`) listed under **Mobile and desktop** redirect URIs (not a web/confidential registration) — see `DEPLOYMENT.md`. Each spike test creates its own uniquely-named file/folder under the signed-in user's OneDrive root and deletes it in a `finally` block, so repeated runs don't accumulate junk.

## Releases and versioning

This server's canonical version lives in `version.txt` (semver, pre-1.0), tracked by its own [release-please](https://github.com/googleapis/release-please) package (own `CHANGELOG.md`, own release PR — see root `release-please-config.json`). `host.json`'s `$.extensions.mcp.serverVersion` (currently `0.1.0`, matching `version.txt`) is now wired into that package's `extra-files`, following `outlook-writeback`'s pattern — a release-please bump patches both in lockstep.

## Key points for a future implementer

- Scopes: `Files.ReadWrite.All` and `Sites.Read.All` only (PRD §6) — no `Sites.ReadWrite.All`.
- `DriveGraphClient` (in `DriveWriteback.Graph/DriveGraphClient.cs`) now carries both the original Phase 0 spike methods (`CreateFolderSegmentAsync`, `GetItemByPathAsync`, `GetItemByIdAsync`, `ReplaceTextContentAsync`, `DeleteItemAsync`) and Phase 1's real tool-surface methods (`CreateFolderPathAsync`, `TryGetItemByPathAsync`, `ResolveFolderPathAsync`, `CreateFileAsync`, `GetItemAsync`, `CreateWithSilentRefreshAuth`). Every method takes an optional `driveId`, mirroring the PRD §5 addressing model (omitted → signed-in user's own OneDrive).
- **`UploadTextContentAsync` is still Phase-0-spike-only, deliberately not migrated.** The plan for this pass called for moving the E2E fixture's helper onto `CreateFileAsync` and deleting `UploadTextContentAsync` once it did — that migration was deliberately deferred: the three original Phase 0 E2E tests that use it are the only currently-verified-green live tests in this folder, and this pass couldn't re-run them to confirm the swap didn't regress anything. `CreateFileAsync` is additive alongside it, not a replacement yet. Do the migration (and then delete `UploadTextContentAsync`) once someone can run the E2E tier live and confirm the swap.
- **Prefer id-based addressing over colon-path addressing for anything just created in the same call chain.** Real Graph behavior, found the hard way (PRD §11): colon-path resolution (`Items["root"].ItemWithPath(path)`) of a folder immediately after creating it can lag, and the failure mode isn't a clean 404 — a `/children` POST against an unresolved colon-path silently lands at the drive root instead. `CreateFolderPathAsync`/`ResolveFolderPathAsync` chain by the `id` each create/lookup returns rather than re-deriving path strings.
- **Path-vs-ID discriminator (`DrivePath.LooksLikeItemId`):** an id never contains `/`, but that alone is insufficient (a root-level filename like `"notes.md"` also contains none). The heuristic adds no-dot/all-alphanumeric/length≥20 (the one observed live OneDrive item id was 34 chars). `GetItemAsync`/`get_item`'s response states which interpretation was used, so a misclassification is visible rather than silent — this is a heuristic, not a guarantee, and a hyphenated SharePoint id (not yet observed) could misclassify as a path.
- **Dry-run (`DriveWriteOptions.DryRun`, default `true`) lives entirely in `DriveWriteService`** — `DriveGraphClient` stays a dumb, dry-run-agnostic Graph wrapper. Flip via the `DRIVE_WRITEBACK_DRY_RUN` app setting/env var, no code change needed.
- The Graph SDK's generated `Me.Drive.Root` only exposes `GetAsync`/`Content` — no `Children`, no `ItemWithPath`. Folder/file operations route through `Drives[driveId].Items["root"]` (or `Items[itemId]` once an id is known) instead (Graph's documented `"root"` item-id alias), which does support both. Verified by compiling against the real `Microsoft.Graph` 6.2.0 assembly, not assumed — see the doc comment on `DriveGraphClient` itself if this needs revisiting on a future SDK upgrade. Confirmed the same way: `Content.ContentRequestBuilder`'s `PutAsync`/`ToPutRequestInformation` accept only `Microsoft.Kiota.Abstractions.DefaultQueryParameters` — no typed `conflictBehavior` support — which is why `create_file` enforces conflict behavior client-side instead.
- No Native AOT (decided against — see PRD §10). Full `Microsoft.Graph` SDK, matching `outlook-writeback`.
- No provisioning script — `DEPLOYMENT.md` is the sole source of truth for standing this up or rebuilding it, mirroring `outlook-writeback`'s posture exactly. An earlier `scripts/Register-EntraApp.ps1` wrapper was tried and removed: it hit a real eventual-consistency issue in its admin-consent verification (see `DEPLOYMENT.md` step 5's "known issue" note) that a second copy of the same logic just added a place for it to hide.
