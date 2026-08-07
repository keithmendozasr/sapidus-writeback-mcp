# PRD — `drive-writeback` MCP Server

**Owner:** Keith Mendoza
**Status:** Completed
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, claude.ai
**Repo location:** `sapidus-writeback-mcp/drive-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)

---

## 1. Problem statement

Anthropic's official Microsoft 365 connector provides read access to SharePoint documents and pages, but write operations against OneDrive and SharePoint document libraries are either absent or gated behind higher-tier plans. This server fills the write gap for file and folder manipulation, mirroring the pattern already established by `outlook-writeback`.

Graph models OneDrive for Business and SharePoint document libraries under a single `driveItem` resource; **what's actually shipped (Phases 1–2) is OneDrive-only.** SharePoint document library support was part of this doc's original scope but is not built — it's tracked separately in `docs/active/PRD-sharepoint-support.md`, including the addressing/permissions differences and library-level behaviors that matter once SharePoint is in scope. (Personal Microsoft/MSA account support is a separate, unexplored gap — see `docs/active/PRD-oss-deployment.md`.)

---

## 2. Goals

1. Create a new file in a specified directory, or in a drive root.
2. Create a directory, with `mkdir -p` semantics (intermediate segments created as needed, idempotent).
3. Rename a file or directory.
4. Move a file or directory.
5. Replace a file's content.
6. Delete a file or directory.

All six are implemented and shipped against the user's OneDrive (Phases 1–2, both deployed and confirmed working end to end from Claude Desktop and claude.ai).

**SharePoint document library support was part of this goal's original scope but is not built.** The running server rejects SharePoint document library drives outright (`ResolveDriveIdAsync`, `DriveGraphClient.cs`) rather than operate against them without proper hardening in place — see `docs/active/PRD-sharepoint-support.md`, which now owns that scope.

## 3. Non-goals (v1)

- Resumable/chunked uploads for files > 4 MB. Tracked in `docs/active/PRD-extended-write-features.md`.
- Binary file content (images, PDFs, Office documents) — text/UTF-8 only in v1. Tracked in `docs/active/PRD-extended-write-features.md`.
- Copy operations (async, long-running). Tracked in `docs/active/PRD-extended-write-features.md`.
- SharePoint **lists**, list items, site pages, or metadata column values. Document library files and folders only. (Applies regardless of whether `docs/active/PRD-sharepoint-support.md` ever lands — lists stay out of scope either way.)
- Sharing links, permissions management, or any operation that grants access to another principal. **Explicitly out of scope indefinitely** — see §9.
- Read/search of file *content*. Handled by the official M365 connector.
- Directory/file **discovery and selection** — enumerating sites, drives, or folder contents to find a target. This server assumes the drive/site/item to act on has already been identified by the caller (e.g. via the official M365 connector); it resolves and validates a given target, it does not browse for one.
- Partial/diff-based content edits. Graph only supports whole-content replacement.
- Permanent delete. Deletes route to the recycle bin, deliberately.
- Cross-drive moves (Graph `PATCH` cannot do this; it requires copy + delete). Tracked in `docs/active/PRD-extended-write-features.md`.

---

## 4. Tool surface

Seven tools. This server does not browse — it assumes the target drive/site/item has already been identified by the caller (typically via the official M365 connector) and passed in as a path or ID. `get_item` is the one read tool, and it exists as a pre-write safety check, not a discovery mechanism.

### 4.1 Write tools

| Tool | Purpose | Destructive? |
|---|---|---|
| `create_file` | Create a new file at a path with inline text content | No (default `fail` on conflict) |
| `create_folder` | Create a folder, `mkdir -p` semantics | No (idempotent) |
| `update_file_content` | Replace the full content of an existing file | Yes — requires `if_match` |
| `rename_item` | Change an item's name in place | Reversible |
| `move_item` | Change an item's parent folder (same drive only) | Reversible |
| `delete_item` | Move an item to the recycle bin | Yes — requires explicit in-chat confirmation |

### 4.2 Resolution / read tool

| Tool | Purpose |
|---|---|
| `get_item` | Fetch metadata for one path or ID: `eTag`, `cTag`, size, folder/file discriminator, checkout state |

### 4.3 Rationale for the read tool

Every write tool still needs a resolved target — `drive_id` + `item_id` or a path Graph can consume. This server does not resolve human-readable descriptions ("the Budget folder on the Finance site") into that target; the caller is expected to supply it already resolved. **This assumption is partially verified — see Phase 0 spike in §11: a live test against a OneDrive-for-Business file confirmed the connector's read output does surface `driveId`/`itemId`, though it could not resolve a folder's `itemId` directly — path-based addressing (`/drives/{drive-id}/root:/{path}:`) is the fallback for folders. Still unconfirmed for SharePoint team-site libraries, where `siteId`'s composite format is actually exercised; not blocking OneDrive-only work.**

`get_item` serves two safety roles once a target is already identified: it is the eTag source for `update_file_content`, and it is the verification primitive before `delete_item` fires, including surfacing checkout state (§8, §12 Q4) before a write is attempted. It is a narrow metadata read, not a browsing tool — file content retrieval and folder enumeration both stay out of scope (§3).

### 4.4 Tool detail

#### `create_file`
```
drive_id?: string       // omit to target the signed-in user's OneDrive
path: string            // drive-relative, e.g. "Shared Documents/notes/2026-07.md"
content: string         // UTF-8 text
conflict_behavior?: "fail" | "rename" | "replace"   // default "fail"
```
Returns: item ID, drive ID, path, size, eTag, web URL.

Parents are **not** auto-created — a typo'd path should fail loudly rather than silently materialize a folder tree. Callers wanting `mkdir -p` call `create_folder` first.

#### `create_folder`
```
drive_id?: string
path: string            // full path of the folder to create
```
Implements `mkdir -p`: walks path segments from the drive root, creating each missing segment. Idempotent — an existing full path returns success with the existing item, not an error.

Graph has no native recursive folder create. Baseline approach: iterative per-segment `POST /drives/{drive-id}/root:/{parentPath}:/children` with `@microsoft.graph.conflictBehavior: fail`, treating HTTP 409 as "already exists, continue." Alternative approaches that rely on path-based upload auto-creating intermediates are **not assumed to work** and must be validated in Phase 0 (§12, Q3).

#### `update_file_content`
```
drive_id?: string
path_or_id: string
content: string
if_match: string        // REQUIRED — eTag/cTag from a prior get_item
```
`if_match` is mandatory, not optional. Whole-file replacement without optimistic concurrency is a silent-data-loss machine: the model reads at T0, someone else writes at T1 (Office web, a sync client, a phone), and the model's write at T2 destroys that change with no trace. **This risk is materially higher on SharePoint than OneDrive** — shared libraries have real co-authors, and eTag churn is constant. Graph returns 412 on mismatch; the tool surfaces that as a distinct, actionable error instructing the caller to re-read.

#### `rename_item` / `move_item`
Both are `PATCH /drives/{drive-id}/items/{id}` — rename sets `name`, move sets `parentReference.id`. Kept as separate tools despite the shared implementation because intent, confirmation posture, and failure modes differ, and a combined tool invites the model to do both when it meant one. A shared internal `PatchItemAsync` handles the call.

`move_item` validates that source and destination resolve to the same `driveId` and fails clearly otherwise. Cross-library and cross-site moves are out of scope.

#### `delete_item`
```
drive_id?: string
path_or_id: string
expected_name: string   // must match the resolved item's name
recursive?: boolean     // default false; required true to delete a non-empty folder
```
Requires explicit in-chat user confirmation before execution — the same posture as `delete_event` in `outlook-writeback`. `expected_name` guards against path/ID drift between the model's read and its write. `recursive` defaults to false so deleting a folder the model *believes* is empty fails rather than silently taking a subtree.

Deletes route to the recycle bin, which lowers blast radius considerably. Note the recycle bin differs by drive type (§8).

**Note on what actually shipped:** the signature above is v1's original design and doesn't match the implemented tool. `expected_name` alone only guards against drift — it doesn't stop the model from skipping the confirmation turn entirely, which is what "the same posture as `delete_event`" actually requires: a real two-call round trip. The implemented `delete_item` adds a fifth parameter, `confirmation_token: string?`, not listed above. The first call previews and returns a token; the second call must echo it back before the delete fires. The mechanism (a stateless, HMAC-signed token binding the two calls together) is `ConfirmationTokenService`, shared with `outlook-writeback`'s `delete_event` — see `drive-writeback/CLAUDE.md`'s "Key points for a future implementer" section for the full design, not repeated here.

---

## 5. Addressing model

With SharePoint originally in scope (see `docs/active/PRD-sharepoint-support.md`), addressing was the single largest design change from rev 1 — the resulting `drive_id`-based model already accommodates SharePoint without needing separate tooling, once that PRD's hardening lands.

**Every tool takes an optional `drive_id`.** Omitted, it targets the signed-in user's OneDrive (`/me/drive`). Supplied, it targets `/drives/{drive-id}`. This keeps the common OneDrive case terse while making SharePoint a first-class target with no separate tool surface.

Within a drive, tools accept **drive-relative paths** as the primary addressing mode, with item IDs accepted anywhere a path is accepted. Paths are what the user speaks; IDs are stable across renames.

### Resolution chain

This server does not perform the resolution itself (§3, §4.3) — the caller supplies an already-resolved `drive_id` and path/item ID. A typical call looks like:

```
create_file(drive_id: "b!xY...", path: "Shared Documents/notes/2026-07.md")
```

Where the `drive_id` came from is out of scope here — expected to come from the caller (e.g. the official M365 connector), pending the Phase 0 verification in §11. Drive and item IDs are opaque Graph identifiers (not human-constructable) and must be passed through exactly as received, never edited or reconstructed.

### Alternative considered: unified locator strings

A single-string scheme (`sp://finance/Shared Documents/budget.xlsx`) would be terser for the model, but it requires the server to maintain a site-name→ID mapping and re-resolve on every call, and it fails ambiguously when two sites share a display name. Explicit `drive_id` is more verbose but unambiguous and Graph-native. **Recommendation: explicit `drive_id`.**

### Path handling requirements

- Percent-encode path segments; reject characters OneDrive/SharePoint disallow (`" * : < > ? / \ |`) before hitting Graph.
- Leading/trailing slashes normalized; `""`, `"/"`, `"root"` all resolve to the drive root.
- No `..` traversal — reject rather than resolve.
- SharePoint-only path rules (stricter URL length limit, blocked file extensions at the library level) are tracked in `docs/active/PRD-sharepoint-support.md`.

---

## 6. Authentication & authorization

Inherits the architecture in `REPO-CONVENTIONS.md` and `outlook-writeback/PRD.md`, with **no deviation**:

- **One App Registration per MCP server** — strict repo invariant. This server gets its own Server app and Connector app. No reuse of the `outlook-writeback` registrations.
- **Server app** holds delegated Graph scopes; refresh token seeded once interactively, stored in Key Vault, silently redeemed per call.
- **Connector app** gates inbound caller access via Authorization Code + PKCE, validated by Azure Easy Auth (App Service Authentication v2).
- **OBO deliberately excluded** — same reasoning as `outlook-writeback`.
- ~~`/admin/reconnect-graph` endpoint within the same Function app for refresh-token re-seeding.~~ **Not built — superseded.** No in-Function-App admin endpoint exists. Re-seeding instead re-runs the same one-time `DriveWriteback.Bootstrap` console tool used for initial seeding (see `DEPLOYMENT.md`'s "One-time bootstrap" section) — same outcome this bullet intended, simpler mechanism, one fewer authenticated surface exposed on the deployed Function App.
- **Entra display name:** "Drive Writeback MCP" — no client name, consistent with the convention.

The tenant is a **managed domain** (cloud authentication, not federated to on-prem ADFS). This matters in one specific way: the one-time interactive seeding flow and the `DriveWriteback.Bootstrap` re-seed both run as plain Entra web sign-ins with no federation hop, WS-Fed redirect, or on-prem STS availability dependency. The seeding design carries over from `outlook-writeback` unchanged.

**Deployed and live.** The Connector app ("Drive Writeback MCP Connector") and Easy Auth are live in production — see `DEPLOYMENT.md`'s "Multi-client OAuth via Easy Auth + Entra ID" section. This server has since also completed the same custom-domain cutover `outlook-writeback` went through (`DEPLOYMENT.md`'s "Custom domain" section): the deployed Function App now serves OAuth through a custom hostname, and `<server>-func.azurewebsites.net` no longer supports a fresh authorize as a result. Cost-management hardening — a Cost Management budget alert on `<server>-rg`, mirroring `outlook-writeback`'s — has **not** yet been added for this server; don't assume it's in place.

### Graph delegated permissions

| Scope | Grants | v1? | Admin consent? |
|---|---|---|---|
| `Files.ReadWrite.All` | All files the signed-in user can access — own OneDrive, shared items, **and SharePoint document libraries** | **Yes** | Required |
| `Sites.Read.All` | Read site metadata; needed to *resolve* sites and enumerate libraries | **Yes** | Required |
| `Sites.ReadWrite.All` | Read/write all site content including lists and pages | **No** | Required |
| `Files.ReadWrite` | Own OneDrive only | Superseded | No |

**The key decision here: `Files.ReadWrite.All` + `Sites.Read.All`, not `Sites.ReadWrite.All`.**

Document library files are `driveItem`s, and `Files.ReadWrite.All` is sufficient to write them. `Sites.Read.All` is needed only for the *resolution* step — `GET /sites/...` and `GET /sites/{id}/drives` require a Sites scope, but only a read one. `Sites.ReadWrite.All` additionally grants write access to lists, site pages, and site configuration, none of which this server touches. Taking it would grant broad authority for zero functional gain.

**`Sites.Selected` — considered and rejected.** It is the correct answer for blast radius: an admin grants the app access to named sites only, and nothing else is reachable regardless of what the model does. But it is an **application** permission requiring client-credentials auth, which would mean running SharePoint on app-only auth while OneDrive stays delegated — a hybrid model that breaks the single-refresh-token design and doubles the auth surface. Revisit only if the path allow-list (§9) proves insufficient in practice.

### Consent and the open-source release

Both `.All` scopes require **tenant admin consent** — a real issue for the planned OSS release, where a deployer who is not a tenant admin cannot consent. Not an issue on this server's actual deployment (owner-as-admin). Full scope covered in `docs/active/PRD-oss-deployment.md`.

---

## 7. Core behaviors & edge cases

### Upload sizing
Graph splits at 4 MB: simple `PUT .../content` at or below, resumable upload session above. v1 supports simple upload only and rejects larger content with a clear error naming the limit.

This is less constraining than it sounds. MCP tool arguments travel as JSON strings through the model's context; a 4 MB file is not something the model can hold or emit. The practical ceiling is far lower. Set an explicit `MaxContentBytes` config (proposed: **1 MB**) rather than letting the Graph limit be the de facto boundary.

### Content encoding
UTF-8 text only in v1. Binary content support (base64 encoding, upload-from-URL) is tracked in `docs/active/PRD-extended-write-features.md`.

### Throttling
SharePoint and OneDrive throttle far more aggressively than Mail — and **SharePoint harder than OneDrive**, with limits applied per-user, per-app, and per-site-collection. Required:
- Honor `Retry-After` on 429 and 503 unconditionally.
- Exponential backoff with jitter, capped retry count.
- Never retry a non-idempotent write on an ambiguous failure without first re-reading state.
- Set a distinctive `User-Agent` — Microsoft deprioritizes traffic without one.

### Conflict behavior
`@microsoft.graph.conflictBehavior` accepts `fail`, `rename`, `replace`. Default `fail` everywhere. `replace` exposed only on `create_file`, only as an explicit caller choice.

### Idempotency
`create_folder` is idempotent by design. `create_file` with `conflict_behavior: fail` is safely retryable. `update_file_content` with `if_match` is safely retryable. `delete_item` on an already-deleted item returns success rather than 404-as-error.

---

## 8. SharePoint-specific behaviors

Not built — full scope (required check-out, required-metadata-column draft-state trap, content approval/versioning, two-stage recycle bin, co-authoring churn) moved to `docs/active/PRD-sharepoint-support.md` §3.

---

## 9. Safety & guardrails

**This server's blast radius is categorically larger than `outlook-writeback`'s, and adding SharePoint enlarges it again.** That server's worst case is a bad unsent draft — fully reversible, zero external effect. This one can overwrite and delete real data, and with `Files.ReadWrite.All` it now reaches **every document library in the tenant the account can access**, not just one user's OneDrive.

Compounding this: a model driving these tools may be acting on content it read from files, email, or the web. Prompt injection that reaches the model reaches these tools. Assume the caller is honest but manipulable.

Required controls:

1. ~~Per-drive path allow-list.~~ **Decided against (§12, Q2): no app-level allow-list.** Write scope is bounded by Graph's own ACL — whatever the signed-in account can access, this server can touch, nothing narrower. This is a deliberate departure from `outlook-writeback`'s posture of always being stricter than the underlying scope. Rationale: the account's actual Graph permissions already define the real blast radius; a hardcoded allow-list only helps against the model *misresolving* a path to somewhere technically-permitted-but-unintended, and that risk is judged not worth maintaining a second access-control layer for. **This changes what "worst case" means for this server** — see the reframed opening paragraph above: the worst case is no longer bounded by a curated sandbox, it is bounded by the account's actual tenant-wide access. The remaining controls below (confirmation, `if_match`, size cap, audit log) are not defense-in-depth on top of an allow-list anymore — they are the only defense, and should be read as such.
2. **No sharing, ever.** `createLink`, `invite`, and permission mutation are not exposed. Exfiltration via an anonymous sharing link is the highest-severity realistic failure mode and it is closed by omission. This matters more on SharePoint, where anonymous links may be enabled tenant-wide.
3. **Confirmation on delete.** Explicit in-chat confirmation, plus `expected_name`, plus `recursive` opt-in.
4. **Mandatory `if_match` on content replacement.** See §4.4.
5. **Size cap.** `MaxContentBytes`, enforced before the Graph call.
6. **Recycle bin, never permanent delete.**
7. **Structured audit log.** Every write emits: tool, drive ID, site ID, resolved path, item ID, eTag before/after, SharePoint version before/after, caller OID from the Easy Auth token, outcome. Written to Application Insights. This is the recovery path when something goes wrong at 11pm. **Implemented as a minimal version, not the full design above.** `DriveWriteService` emits one `ILogger` line per call — tool name, resolved path or ID, outcome, and eTag where relevant (`if_match` doubles as the eTag-before value on `update_file_content`; `etag-after` is logged on create/update) — which Application Insights captures automatically for a deployed Function App; no separate audit-log plumbing was built. Not yet captured: drive ID (every log line is silent on which drive was targeted), site ID (moot today — OneDrive-only, no SharePoint sites in scope), and caller OID from the Easy Auth token. The last was deliberately deferred pending boundary-7b's deployment (`DriveWriteService.cs`'s own doc comment says as much) and remains unimplemented even though boundary-7b is now live (see §6) — a real gap, not a stale note. `DriveItemDeletionService`'s preview and confirm steps emit no log lines of their own; only the final `DriveWriteService.DeleteItemAsync` call is logged.
8. **Dry-run mode.** Server-level config flag that validates and resolves but does not mutate. Invaluable during Phase 1 and for testing against the live tenant before trusting writes.

---

## 10. Hosting & runtime

Unchanged from `outlook-writeback`:

- C# / .NET 10 LTS, isolated worker model
- Azure Functions, Flex Consumption plan on Linux
- Azure Key Vault for the refresh token
- MSAL for .NET
- Application Insights

**Native AOT — decided against.** `outlook-writeback`'s `.csproj` doesn't enable `PublishAot` either — no `PublishAot`/`RuntimeIdentifier`/`InvariantGlobalization` settings anywhere in that project tree — so there's no tension with taking the full `Microsoft.Graph` SDK here, matching the only working precedent this repo has. The Phase 0 "confirm Native AOT publish path" spike (§11) is removed as moot — there's no publish path to confirm once AOT isn't attempted.

### Dependency notes (session-start check)

- **`ModelContextProtocol` (C# SDK):** current stable **1.4.1**; **2.0.0-preview.1** published. **Recommend hanging back on 1.4.x.** A 2.0 major on a preview tag, on an SDK that moved 0.4 → 0.6 → 1.x within roughly a year, is exactly the profile to let settle. Revisit once 2.0.0 GA has had a patch or two.
- **`ModelContextProtocol.AspNetCore`:** track in lockstep with the core package.
- **Microsoft.Graph SDK — decided: take the full SDK, matching `outlook-writeback`.** `outlook-writeback` already carries `Microsoft.Graph` 6.2.0 (checked 2026-08-02 — still current, no upgrade pending). The SDK's reflection-heavy behavior is only disqualifying if AOT is in play, and it isn't (see above), so there's no remaining reason to hand-roll `HttpClient` against Graph REST — especially given SharePoint site/drive resolution is more verbose to hand-roll than mail operations were.
- **MSAL for .NET:** no action; matches `outlook-writeback`.

*Raised, not applied.*

---

## 11. Phasing

**Phase 0 — Validation spikes**
- Confirm `mkdir -p` approach (§4.4) against both a OneDrive and a SharePoint library. **Resolved for OneDrive, confirmed live.** Root cause: colon-path addressing (`Items["root"].ItemWithPath(path)`) of a folder immediately after creating it is not reliable — Graph's path-resolution index can lag behind the item actually existing, and a `/children` POST against an unresolved colon-path silently lands at the drive root rather than erroring. This is what caused earlier runs' `a`/`b`/`c` folders to land at the drive root instead of nesting, confirmed both by a diagnostic test logging each created item's actual `ParentReference.Path`/`Id` and by a direct OneDrive-web check. **Fix, now passing:** `CreateFolderPathAsync` chains by the item `id` each create returns rather than re-deriving a path string from segment names — id-based addressing needs no path resolution, so it isn't exposed to this at all. **Carries into Phase 1's `create_folder`/`get_item` design: prefer id-based addressing over colon-path addressing for anything just created in the same call chain.** (Two earlier diagnoses on this same failure — a "propagation lag on GET" theory, then a "colon-path GET itself is unreliable" theory — were both wrong, each built on top of the real bug without having isolated it, and both were briefly written into this doc as confirmed findings before being retracted. Noted here, not scrubbed from git history, as a reminder to verify against a passing run before writing "resolved.") SharePoint confirmation tracked in `docs/active/PRD-sharepoint-support.md` §6.
- Confirm eTag vs cTag semantics for `If-Match` on `/content` — these differ, and picking wrong yields either false 412s or no protection at all. **Resolved for OneDrive: both are honored.** A deliberately wrong tag reliably 412'd (confirming the protection is real), and a content replacement succeeded using either the item's `eTag` or its `cTag` as `If-Match`, tested independently on two freshly-created files so one attempt couldn't invalidate the other's tag. Either is safe to use for `update_file_content`'s mandatory `if_match` on OneDrive. SharePoint verification tracked in `docs/active/PRD-sharepoint-support.md` §6.
- Confirm site ID and item ID formats so the path-vs-ID discriminator in §5 is sound. **Partially resolved:** a OneDrive item id was observed as `0176NADPLJLSYNNMETHFBJ7IQMXSGTZ4MA` (34 chars, alphanumeric, no `/`) — confirms the discriminator §5 needs (an id never contains `/`, so it can't be confused with a drive-relative path) holds for OneDrive. Site ID format is still untested — tracked in `docs/active/PRD-sharepoint-support.md` §6.
- Confirm `Files.ReadWrite.All` + `Sites.Read.All` is genuinely sufficient for all six use cases against a real library — i.e. that no write path demands `Sites.ReadWrite.All`. **This is the highest-value spike in Phase 0**; if it fails, §6 changes materially. **Resolved for OneDrive:** all three spikes (mkdir-p, eTag/cTag, item-ID) now pass end-to-end under only these two scopes, zero `403`s anywhere. SharePoint verification tracked in `docs/active/PRD-sharepoint-support.md` §6.
- Reproduce the required-metadata-column draft-state trap deliberately, so the detection logic is written against observed behavior. Tracked in `docs/active/PRD-sharepoint-support.md` §6 — not blocking Phase 1, which is OneDrive-only for its initial safe-create tools.
- ~~Confirm Native AOT publish path with the chosen Graph client.~~ **Moot — decided against AOT (§10).** `outlook-writeback` never actually used Native AOT despite this doc previously assuming otherwise; drive-writeback now explicitly matches that (full `Microsoft.Graph` SDK, no `PublishAot`), so there's no publish path left to confirm.
- **Confirm the official M365 connector's read output surfaces `driveId`/`siteId`/`itemId`** in a form passable to this server. This server assumes discovery/selection happens upstream (§3) and never browses — if the connector's output doesn't carry usable Graph IDs, that assumption breaks and a narrow `resolve_path` tool becomes necessary before Phase 1 can proceed. **Partially resolved:** live test against a OneDrive-for-Business file confirmed `driveId`/`itemId` are returned correctly. The connector could not resolve a folder's `itemId` directly (only the file's) — path-based addressing (`/drives/{drive-id}/root:/{path}:`) is the fallback for folders, so §4.3's "caller supplies a resolved ID" assumption should read "resolved ID or resolved path." SharePoint team-site verification (where `siteId`'s composite-triple format actually gets exercised) tracked in `docs/active/PRD-sharepoint-support.md` §6 — not blocking OneDrive-only usage.
- **Confirm "Assignment required" on the Connector app's Enterprise Application works on the tenant's actual Entra license tier.** The Q8 action item (restrict Connector app sign-in to a single user, closing the any-tenant-user-can-authenticate gap) depends on this. **Resolved:** tenant is Entra ID Free. Individual user assignment (not group-based) works on Free tier without a P1/P2 upgrade — confirmed compatible.

**Phase 1 — Safe creates**
`get_item`, `create_folder`, `create_file`. **Implemented and deployed** — see `drive-writeback/CLAUDE.md` Status for detail. Two open questions surfaced during implementation, both **now resolved against a live tenant**:
- Whether Graph's `@microsoft.graph.conflictBehavior` is honored as a raw query parameter on the simple-upload content PUT the way it is on `POST .../children`. **Resolved: yes** — a deliberately-conflicting PUT with `?@microsoft.graph.conflictBehavior=fail` appended 409'd as expected (`DriveGraphClientE2ETests.ContentPut_conflictBehavior_query_parameter_is_an_open_question`). The first live run of this test surfaced a separate, now-fixed test bug: the 409 arrived as a bare `Microsoft.Kiota.Abstractions.ApiException` rather than the richer `ODataError` subtype, because this raw, manually-constructed request has no 409 error factory registered — the test's catch clause needed broadening to the base `ApiException` type. `create_file`'s `conflict_behavior` handling never depended on this answer either way — it's enforced entirely client-side (existence pre-check, then unconditional PUT) — so this was a recorded gap in tenant-behavior knowledge, not a blocker, and stays that way now that it's closed.
- Whether a content PUT to a path whose parent is missing auto-vivifies the parent tree or 404s. **Resolved: it auto-vivifies** (`DriveWriteServiceE2ETests.ContentPut_to_a_missing_parent_is_an_open_question`). Moot for this server's own behavior either way — `CreateFileAsync` guards against a missing parent before ever reaching Graph — but now a confirmed fact rather than an open one.

**Phase 2 — Mutation surface**
`update_file_content`, `rename_item`, `move_item`, `delete_item`. **Implemented and deployed** — all confirmed working end to end from Claude Desktop, including the eTag-quoting fix (see `drive-writeback/CLAUDE.md`). 12 E2E tests (the 3 original Phase 0 spikes plus 9 added across Phase 1/2) passed against a live tenant at the time this phase shipped, including the new Phase 2 round trips: `rename_item`, `move_item`, delete-idempotency, the `update_file_content` dry-run/real round trip, and the full `delete_item` preview-then-confirm confirmation flow through `DriveItemDeletionService`. A 13th E2E test was added afterward, confirming the SharePoint-rejection guard's allow-list — see `drive-writeback/CLAUDE.md`'s Status section for that later addition; it isn't part of this phase's own scope.

**Phase 3 — SharePoint hardening**
Not started. Full scope moved to `docs/active/PRD-sharepoint-support.md`. Until it lands, the server actively rejects SharePoint document library drives (`ResolveDriveIdAsync` in `DriveGraphClient.cs`, `SharePointDriveNotSupportedException`) rather than operate against them with none of this hardening in place — a deliberate scope decision, not silent neglect. See `drive-writeback/CLAUDE.md`'s Status section for current state.

**Phase 4 — Deferred**
Not started. Full scope moved to `docs/active/PRD-extended-write-features.md`.

---

## 12. Open questions

1. ~~What is the actual driving workflow?~~ **Decided:** text-only v1 confirmed adequate — binary content (PDF/XLSX/images) is not needed given `MaxContentBytes` (1 MB) is not expected to bind under normal use. `MaxContentBytes` stays at 1 MB. Write-target scope is "anything in the tenant the user has access to" (§12 Q2) rather than a fixed workflow like `outlook-writeback`'s budget reconciliation.
2. ~~Which sites and libraries?~~ **Superseded and decided (§9):** scope is "anything in the tenant the user has access to." No app-level allow-list — access control is Graph's own ACL. The static per-drive allow-list guardrail is removed from §9.
3. ~~`mkdir -p` implementation~~ **Decided and confirmed for OneDrive (§11):** chain by item `id`, not by re-derived colon-path strings — colon-path addressing of a folder immediately after creating it is unreliable (Graph's path-resolution index can lag, and the failure mode is a silent landing at drive root, not a clean error). SharePoint still unverified.
4. **Check-out handling:** **Decided: detect-and-fail.** `get_item` surfaces checkout state so the model can check before attempting a write; a failed write due to checkout returns a distinct, explicit error (not a generic Graph fault) stating the item is checked out, so the model can relay that to the user rather than attempting to move/update it.
5. ~~eTag or cTag for `If-Match`, and does it differ on SharePoint?~~ **Resolved for OneDrive: both work (§11).** Either the item's `eTag` or its `cTag` is honored by `If-Match` on `/content` — `update_file_content` can accept either without a false 412. SharePoint semantics still unverified.
6. ~~Same Function App or separate per server?~~ **Decided: separate Function App per server.** Easy Auth (App Service Authentication v2) is configured at the Function App level, not per-route — there's no way to point different paths within one Function App at different identity providers. Sharing a Function App would force `outlook-writeback` and `drive-writeback` to share one Connector app, which breaks the one-Entra-app-per-server invariant. No competing reason to share (each server already gets its own subdomain — e.g. `outlook-writeback.<tenant-domain>`, `drive-writeback.<tenant-domain>` — not a shared one). **`REPO-CONVENTIONS.md` should state explicitly:** "One Function App per MCP server, matching the one-Entra-app-per-server invariant. Each server gets its own subdomain and its own Flex Consumption plan."
7. ~~Should `list_children` and `list_sites` paginate?~~ **Moot.** Both tools are out of scope — this server does not browse (§3, §4). Discovery is the caller's responsibility.
8. ~~Does the OSS graceful-degradation requirement belong in v1 or v2?~~ **Decided: split.** Full detail moved to `docs/active/PRD-oss-deployment.md`. **Related, not yet in the doc:** by default, Entra allows any tenant user to sign in to an app registration once it exists, unless sign-in is restricted. Since all Graph calls run under the single seeded Server-app refresh token (not per-caller delegation), an unrestricted Connector app would let any tenant user drive this server with the owner's own Graph access. **Action:** set "Assignment required" = Yes on the Connector app's Enterprise Application object and assign only the intended user — already confirmed working on this tenant's Entra ID Free license tier (§11 Phase 0). Applies to `outlook-writeback` too — should also be captured in `REPO-CONVENTIONS.md`.
9. Two-stage `commit`/publish tool for the SharePoint required-column/content-approval draft-state trap — moved to `docs/active/PRD-sharepoint-support.md` §5.

---

## 13. Success criteria

- All six use cases executable end-to-end from Claude Desktop and claude.ai against the user's OneDrive on the live tenant. **Confirmed live** — Phases 1–2 deployed, all six write tools plus `get_item` verified working from Claude Desktop. SharePoint library support is a separate criterion, tracked in `docs/active/PRD-sharepoint-support.md` §7.
- No writes possible beyond what the signed-in account's own Graph permissions allow (§9) — enforcement is Graph's ACL, not an app-level list; verified by test that a write to an inaccessible drive/site fails with Graph's own 403/404, not a silent success.
- No path to permanent data loss without an explicit user confirmation turn.
- Required-column and check-out draft states detected and reported, never silently reported as success — SharePoint-specific, tracked in `docs/active/PRD-sharepoint-support.md` §7.
- Every mutation reconstructable from the audit log, including SharePoint version numbers where applicable — tracked in `docs/active/PRD-sharepoint-support.md` §7.
- Full `Microsoft.Graph` SDK, Native AOT dropped — see §10.

---

## 14. References

- Microsoft Graph — driveItem resource: https://learn.microsoft.com/en-us/graph/api/resources/driveitem
- Microsoft Graph — drive resource: https://learn.microsoft.com/en-us/graph/api/resources/drive
- Microsoft Graph — site resource and addressing: https://learn.microsoft.com/en-us/graph/api/resources/site
- Microsoft Graph — list a site's drives: https://learn.microsoft.com/en-us/graph/api/drive-list
- Microsoft Graph — upload or replace driveItem content: https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
- Microsoft Graph — create a folder: https://learn.microsoft.com/en-us/graph/api/driveitem-post-children
- Microsoft Graph — move a driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-move
- Microsoft Graph — delete a driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-delete
- Microsoft Graph — throttling guidance: https://learn.microsoft.com/en-us/graph/throttling
- Microsoft Graph — permissions reference: https://learn.microsoft.com/en-us/graph/permissions-reference
- Microsoft Graph — Sites.Selected overview: https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread
- MCP C# SDK: https://github.com/modelcontextprotocol/csharp-sdk
- MCP specification: https://modelcontextprotocol.io/specification
- Repo-internal: `REPO-CONVENTIONS.md`, `outlook-writeback/PRD.md`
