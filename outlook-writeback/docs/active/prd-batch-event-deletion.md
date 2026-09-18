# PRD: Batch Calendar Event Deletion

**Owner:** Keith Mendoza
**Status:** Draft — design approved, implementation not started
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, Claude Code CLI
**Repo location:** `sapidus-writeback-mcp/outlook-writeback/` (parent PRD: `docs/archive/prd-create-update.md`)
**Tracking issue:** [#19](https://github.com/keithmendozasr/sapidus-writeback-mcp/issues/19)

---

## 1. Problem Statement

`delete_event` only accepts one `eventId` per call. Deleting several events (e.g. "cancel all four dentist reminders" or "remove everything from the old team's calendar sync") currently means one full preview-then-confirm round trip per event, each minting and echoing back its own confirmation token. Issue #19 asks for a single confirmation round trip that covers a whole list, with the added requirement that the confirming call be allowed to drop some of the originally-previewed events (the user reviews the preview and says "delete all of these except that one").

## 2. Goals

- Let `delete_event` preview and delete a list of events in one two-step, confirmation-gated round trip, using a single confirmation token for the whole batch — not one token per event.
- Let the confirming call name a *subset* of the previewed events (some dropped, none added) and delete only those.
- Keep the two-step, nothing-destructive-on-first-call posture that today's single-event `delete_event` already has.
- One failing event (already deleted, Graph error, etc.) must not block the rest of the batch from being attempted.
- Fold in migrating `outlook-writeback` off its private `DeleteConfirmationTokenService` onto `shared/Sapidus.Writeback.Shared`'s `ConfirmationTokenService` (a deferred cleanup already flagged in `drive-writeback/CLAUDE.md`), since the new batch token logic needs to be added to one of these two classes and this is the natural point to stop maintaining both.

## 3. Non-Goals

- No true Graph `$batch` HTTP batching — deletes are attempted sequentially against `DELETE /me/events/{id}`, same as today, just looped. This is an interactive, single-user tool; throughput isn't a driver.
- No way to *add* an event ID at confirm time that wasn't in the original preview — the confirming call's IDs must be a subset of what was previewed, never a superset. An ID outside the previewed set fails closed for that ID (see §5).
- No atomicity guarantee across the batch — a partial success (some deleted, some failed) is an expected, reported outcome, not an error condition for the whole call.
- No change to `drive-writeback`'s `delete_item` in this PRD — only the shared token primitive gains batch capability; wiring it into `delete_item` is a future option, not part of this work.

## 4. Functional Requirements

### `delete_event`

- **Input change:** `eventId` (string) becomes `eventIds` (array of strings, 1–25 items). A single-event delete is just a 1-element array; this is a breaking schema change, acceptable pre-1.0 per this repo's `bump-minor-pre-major` convention.
- **Cap:** requests with more than 25 IDs are rejected before any Graph call, with an error naming the limit. Chosen to bound blast radius and keep the preview response readable, not driven by any Graph-imposed limit.
- **Step 1 (preview, `confirmationToken` omitted):** fetch each event by ID (best-effort — a lookup failure for one ID doesn't stop the others from being previewed), return each found event's subject/time, list any IDs that couldn't be found, and issue **one** confirmation token covering the full set of IDs that were requested.
- **Step 2 (confirm, `confirmationToken` supplied):** caller also supplies `eventIds` — any non-empty subset of the IDs from step 1's request. Any ID not in that original set fails validation for that ID specifically (see §5) rather than invalidating the whole call. Each ID that passes validation is deleted independently; the response is a per-event summary, e.g. "4 of 5 deleted. Failed: `AAMk...ZZZ` (already deleted)."
- **Behavior on a stale/expired/tampered token:** same as today's single-event case — the response says the token is invalid/expired and instructs the caller to preview again with `delete_event` and no `confirmationToken`.

## 5. Design Notes

**Single self-contained batch token, not per-event tokens.** Step 1 signs the full, canonicalized (sorted, deduplicated) set of previewed event IDs into one token, using the same stateless HMAC-signed shape `delete_event` already has today (`{payload}.{signature}`, payload = ids + expiry, 5-minute TTL) — just with a set in the resource-id slot instead of a single ID. The token is reversible/decodable (base64, not encryption) but not forgeable or alterable without the signing key; this is an unchanged property from today's single-ID token, not a new exposure — Graph event IDs aren't secrets.

Step 2 validates by:
1. Verifying the signature and expiry, exactly as today.
2. Decoding the original full ID set embedded in the token payload.
3. Checking that every ID in step 2's requested subset is a member of that decoded set. IDs outside it are rejected individually; IDs inside it that the caller didn't resend are simply not acted on (dropping an event from the confirm call is not an error — it's the "user asked to exclude this one" case from issue #19).

The ID set is joined with a delimiter guaranteed not to collide with a Graph event ID (a control character, not a printable delimiter like `,` — Graph IDs aren't validated against a fixed character set closely enough to assume commas are safe).

**New capability on the confirmation-token primitive:** alongside today's `Issue(resourceId)` / `Validate(resourceId, token)`, add `IssueBatch(IEnumerable<string> resourceIds)` / `ValidateSubset(IEnumerable<string> requestedIds, string token)`. Both are pure functions of their inputs plus the signing key — no new state, no change to the TTL model.

**Consolidation onto `shared/Sapidus.Writeback.Shared.Confirmation.ConfirmationTokenService`.** `outlook-writeback`'s own `DeleteConfirmationTokenService` (`OutlookWriteback.Graph/Confirmation/`) is a private, pre-shared-extraction duplicate of the same code that now lives in `shared/` for `drive-writeback`'s `delete_item` to use. Rather than adding `IssueBatch`/`ValidateSubset` to both copies (or only the local one, deepening the drift), this PRD folds in the already-flagged cleanup: delete `OutlookWriteback.Graph/Confirmation/DeleteConfirmationTokenService.cs`, switch `EventDeletionService` and `Program.cs`'s DI wiring to the shared `ConfirmationTokenService`, and add the two new batch methods there. This makes batch-delete capability available to `drive-writeback`'s `delete_item` later without re-deriving the token mechanics, though wiring it into `delete_item` itself is out of scope here (§3).

**`EventDeletionService` becomes batch-shaped.** `RequestDeletionAsync`/`ConfirmDeletionAsync` (currently single-`eventId` methods) become `RequestDeletionAsync(IEnumerable<string> eventIds)` / `ConfirmDeletionAsync(IEnumerable<string> eventIds, string confirmationToken)`, each returning a per-ID result collection (found/not-found for preview; deleted/failed-with-reason for confirm) instead of a single record. `DeleteEventTool` formats that collection into the tool's response text.

## 6. Testing

- Unit coverage for `ConfirmationTokenService.IssueBatch`/`ValidateSubset`: valid round trip, subset confirms successfully, an ID outside the original set fails for just that ID, tampered signature, expired token, malformed token — mirroring the existing single-ID test shape in `Sapidus.Writeback.Shared.Tests`.
- Unit/integration coverage for `EventDeletionService`'s batch preview/confirm, including partial-failure paths (one Graph delete throws, the rest still complete) and the over-25-IDs rejection.
- `dotnet test --filter "Category!=E2E"` green, matching every prior PRD's completion bar in this folder.

## 7. Open Questions

None outstanding — all forks in this design (tool shape, token design, batch cap, partial-failure handling, and the shared-service consolidation) were resolved in design discussion before this PRD was written.

## 8. Milestone

Single phase — this extends one already-shipped tool and its supporting confirmation-token service; it is not a new server or new tool surface. Done when the implementation lands, `EventDeletionService`'s local `DeleteConfirmationTokenService` is fully removed in favor of the shared service, and `dotnet test --filter "Category!=E2E"` is green.
