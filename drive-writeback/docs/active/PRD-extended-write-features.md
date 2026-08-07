# PRD — Extended Write Features for `drive-writeback`

**Owner:** Keith Mendoza
**Status:** Not started
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, claude.ai
**Repo location:** `sapidus-writeback-mcp/drive-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)
**Extracted from:** `docs/active/PRD-drive-write.md` ("v1" below), §11 Phase 4 and the matching §3 Non-goals bullets. This is a backlog of individually-small features v1 deliberately excluded rather than one coherent capability epic — kept together as one doc for now; split further only once a specific item below is actually picked up for real design work. Every section of v1 not called out here is unchanged and unaffected by this PRD.

---

## 1. Problem statement

v1 shipped a deliberately narrow write surface: single-shot uploads up to 1 MB, UTF-8 text content only, same-drive moves only, no copy operation, no recycle-bin restore. Each restriction was a reasonable v1 cut, not a permanent one — this doc tracks what closing each gap would take, so the reasoning isn't lost and each can be picked up independently.

## 2. Candidate features

### 2.1 `copy_item`

Async, long-running Graph operation (`202 Accepted` + a monitor URL to poll) — categorically different from every other tool in this server, which are all synchronous single-call operations. Needs its own polling/status-reporting design before it fits the existing tool shape. Not yet designed.

### 2.2 Binary file content

v1 is UTF-8 text only. Supporting images, PDFs, and Office documents needs a `content_encoding: "utf8" | "base64"` parameter (the natural extension point — v1's `content` parameter and `MaxContentBytes` cap are already encoding-agnostic at the transport level). Base64 through the model's context window is expensive per byte, though, which is why v1 flagged upload-from-URL (§2.3 below) as the likely better path for genuinely large binary content rather than just adding a `base64` encoding option to the existing `content` parameter.

### 2.3 Upload-from-URL

Not in v1 at all. Would let the model hand the server a URL to fetch and write, rather than passing content inline — sidesteps the context-window cost binary content raises (§2.2) entirely for anything already reachable by URL. Not yet designed; would need its own fetch/size-limit/content-type validation story before landing.

### 2.4 Resumable/chunked uploads

Graph's own answer for content over 4 MB (v1 §7 caps `MaxContentBytes` at 1 MB, well under Graph's own 4 MB simple-upload ceiling, specifically to avoid needing this). Only becomes necessary if `MaxContentBytes` itself needs to grow past 4 MB — no current driver for that.

### 2.5 Restore-from-recycle-bin

v1's `delete_item` moves to the recycle bin, never permanently deletes (v1 §3, §9) — the safety property is already there. What's missing is the read/restore half: a tool to list or restore a previously-deleted item. Out of scope for v1 because this server doesn't do discovery/browsing generally (v1 §3) — restore-by-known-item-id might not have that problem, worth a closer look if picked up.

### 2.6 Cross-drive moves

Graph's `PATCH` move operation cannot move an item between drives — it requires copy + delete, meaning this depends on §2.1 (`copy_item`) landing first. `move_item` today validates source and destination resolve to the same `driveId` and fails clearly otherwise (v1 §4.4) — that guard stays until this is built.

## 3. Non-goals reaffirmed from v1

- Partial/diff-based content edits — Graph only supports whole-content replacement; not something client-side chunking can fix.
- SharePoint lists, list items, site pages, or metadata column values — document library files/folders only, regardless of which of the above features land.

## 4. Success criteria

Deliberately none set at the PRD level — each feature in §2 is independent enough that success criteria belong with that feature's own design work once it's picked up, not as one combined bar for this backlog doc.

## 5. References

- Microsoft Graph — upload or replace driveItem content: https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
- Microsoft Graph — move a driveItem: https://learn.microsoft.com/en-us/graph/api/driveitem-move
- Repo-internal: `../REPO-CONVENTIONS.md`, `docs/active/PRD-drive-write.md`
