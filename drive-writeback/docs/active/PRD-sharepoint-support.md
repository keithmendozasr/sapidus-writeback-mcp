# PRD — SharePoint Document Library Support for `drive-writeback`

**Owner:** Keith Mendoza
**Status:** Not started
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, claude.ai
**Repo location:** `sapidus-writeback-mcp/drive-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)
**Extracted from:** `docs/active/PRD-drive-write.md` ("v1" below) — v1's original goal (§2) was for all six write tools to work against both OneDrive and SharePoint document libraries, but what actually shipped (Phases 1–2) is OneDrive-only; the running server rejects any SharePoint document library drive outright (`SharePointDriveNotSupportedException` in `DriveGraphClient.ResolveDriveIdAsync`) rather than operate against one without this PRD's hardening in place. This doc collects every SharePoint-specific piece of scope v1 already specified but never built, so it can be tracked and eventually implemented as its own unit of work rather than staying scattered across a PRD that's otherwise describing shipped behavior. Every section of v1 not called out here is unchanged and unaffected by this PRD — see v1 for the tool surface, addressing model, and guardrails this builds on.

---

## 1. Problem statement

v1's tool surface (`create_file`, `create_folder`, `update_file_content`, `rename_item`, `move_item`, `delete_item`, `get_item`) is drive-type-agnostic at the Graph API level — `driveItem` is the same resource whether the backing drive is a personal OneDrive, a OneDrive for Business, or a SharePoint document library. But several SharePoint-only library behaviors will silently produce wrong-looking results if the tools don't detect and handle them explicitly, and v1 deliberately shipped without that handling rather than risk a data-loss or invisible-write failure mode on a library nobody had tested against yet. This PRD is that handling.

## 2. Goal

Extend all six write tools plus `get_item` to work correctly against SharePoint document libraries, not just OneDrive — completing v1 §2's original goal. Concretely: `DriveGraphClient.ResolveDriveIdAsync`'s allow-list (currently `"business"`/`"personal"` only) admits `"documentLibrary"` once the behaviors below are handled, and `SharePointDriveNotSupportedException` goes away.

## 3. SharePoint-specific behaviors to handle

These are the failure modes that do not exist on OneDrive and that will silently produce wrong-looking results if unhandled. Each needs explicit detection and a named error, not a pass-through Graph fault.

**Required check-out.** Libraries can require check-out before edit. `update_file_content` against such a library fails, or succeeds into a state nobody else can see. `get_item` must return checkout state (it already has a "Checkout state: not tracked for OneDrive in this phase" placeholder field ready to carry this). Decision needed: detect-and-fail (matching v1's `get_item`-based check-out decision for the case where it's already known to be checked out), or expose `checkout`/`checkin` tools and manage the cycle.

**Required metadata columns.** A library with required columns will accept an upload and leave the file **checked out to you in a draft state, invisible to everyone else**. This is the single most confusing SharePoint gotcha — the write "succeeds," the API returns 201, and the file effectively does not exist for other users. The server must detect this post-write and warn explicitly.

**Content approval and minor versions.** Similar draft-state trap: the file is written but pending approval and not visible at the published version. Whether a two-stage `commit`/publish tool is needed to close this gap is open — see §5 below. If built, it should mirror `delete_item`'s two-stage confirmation posture (v1 §4.4), not auto-publish a hidden draft.

**Versioning.** Libraries version by default. Every `update_file_content` creates a new version rather than overwriting history. This is a **safety win** — a genuine undo path OneDrive personal does not reliably provide — and should be noted in the audit log (version number before/after).

**Two-stage recycle bin.** Deleted items go to the site recycle bin, then the site collection recycle bin, with a combined retention window (93 days by default). Recovery is possible but the path differs from OneDrive's. Document it; do not automate it.

**Co-authoring churn.** Real concurrent editors mean 412s from `if_match` will be routine rather than exceptional. The error message must make the re-read-and-retry loop obvious to the model (v1 already does this for OneDrive; verify the same message reads correctly at SharePoint's higher churn rate).

## 4. Addressing model additions

v1 §5's addressing model (`drive_id` + drive-relative path, item IDs accepted anywhere a path is) is already SharePoint-shaped — no new tool surface needed. Two SharePoint-only path rules still need enforcing:

- **SharePoint URL length limits are stricter than OneDrive's** (~400 characters for the full decoded URL). Validate against the tighter limit for SharePoint drives rather than surfacing an opaque Graph 400.
- SharePoint blocks certain file extensions at the library level. Detect the resulting error and surface it as a named condition.

Composite `siteId` format (`contoso.sharepoint.com,<site-collection-guid>,<site-guid>`) is exercised when resolving a team-site library's `drive_id` upstream (via the official M365 connector) — this server itself never constructs one, only receives it, per v1 §5's resolution-chain model.

## 5. Open questions

- **Does a two-stage `commit`/publish tool belong in this PRD's scope, for the required-column/content-approval draft-state trap (§3 above)?** Undecided. If built, mirror `delete_item`'s two-stage confirmation posture (v1 §4.4), not auto-publish a hidden draft.
- Check-out handling: detect-and-fail vs. full `checkout`/`checkin` tools — see §3 above.

## 6. Phase 0 spikes still needed

v1's Phase 0 validated three behaviors (`mkdir -p` via id-chaining, eTag/cTag semantics for `If-Match`, item-ID format) against OneDrive only. Each needs its own SharePoint confirmation before this PRD's implementation can be trusted:

- **`mkdir -p` approach.** Untested on SharePoint. The OneDrive fix (chain by item `id`, not re-derived colon-path strings — see v1 §11) should carry over, but SharePoint's path-resolution index may behave differently.
- **eTag vs. cTag for `If-Match` on `/content`.** Resolved for OneDrive (either tag works) — still open for SharePoint, where versioning may change the semantics.
- **Site ID and item ID formats.** OneDrive item IDs are confirmed 34-char alphanumeric with no `/`. Site ID's composite-triple format is still unverified against a live tenant.
- **`Files.ReadWrite.All` + `Sites.Read.All` sufficiency.** Confirmed for OneDrive (zero 403s across all three spikes). Still open for SharePoint — no team site has been available in the tenant to exercise `Sites.Read.All`'s actual resolution boundary.
- **Reproduce the required-metadata-column draft-state trap deliberately**, so §3's detection logic is written against observed behavior rather than assumption.

**Blocking dependency: no SharePoint team site exists in the tenant yet.** None of the above can be spiked until one is available — this is infrastructure, not code, and gates the start of implementation work on this PRD.

## 7. Success criteria

- All six write tools plus `get_item` executable end-to-end from Claude Desktop against at least one real SharePoint document library on the live tenant.
- Required-column and check-out draft states detected and reported, never silently reported as success.
- Every mutation reconstructable from the audit log, including SharePoint version numbers before/after.
- `SharePointDriveNotSupportedException` removed once the above is verified — `"documentLibrary"` joins the `ResolveDriveIdAsync` allow-list.

## 8. References

- Microsoft Graph — driveItem resource: https://learn.microsoft.com/en-us/graph/api/resources/driveitem
- Microsoft Graph — site resource and addressing: https://learn.microsoft.com/en-us/graph/api/resources/site
- Microsoft Graph — list a site's drives: https://learn.microsoft.com/en-us/graph/api/drive-list
- Microsoft Graph — Sites.Selected overview: https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azuread
- Repo-internal: `../REPO-CONVENTIONS.md`, `docs/active/PRD-drive-write.md`
