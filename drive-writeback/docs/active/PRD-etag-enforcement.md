# PRD — eTag Enforcement Hardening for `drive-writeback`

**Owner:** Keith Mendoza
**Status:** Implemented — pending deployment
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, claude.ai
**Repo location:** `sapidus-writeback-mcp/drive-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)
**Extends:** `docs/archive/PRD-drive-write.md` — a hardening pass on that PRD's already-shipped `get_item` and `update_file_content` tools (Phases 1–2), not a new phase of its own.

---

## 1. Problem statement

Resolves [issue #17](https://github.com/keithmendozasr/sapidus-writeback-mcp/issues/17).

`update_file_content` requires `if_match`, sourced from `get_item`, and Graph enforces it server-side as optimistic concurrency control. This works correctly only if the `if_match` value was captured before the edit was made. But the natural-seeming call sequence — read content, edit, `get_item` right before writing, then `update_file_content` — defeats the protection entirely: `get_item` called immediately before the write always returns the file's *current* server state, so `if_match` always matches itself. An outside edit that happened between the original content read and that final `get_item` call is silently overwritten, with no error.

Concretely, this sequence is unsafe today:

1. Read file content (e.g. via the M365 connector's read tool) — no metadata captured.
2. *(outside edit happens here — undetected)*
3. `get_item` right before writing — returns the *new*, post-edit etag.
4. `update_file_content` with that etag — succeeds, clobbering the outside edit.

**Root cause, in two parts:**

- `get_item`'s current docstring (`Functions/GetItemTool.cs`) frames it as "a pre-write safety check," which invites calling it right before the write — the one place it provides no protection.
- The server's `host.json` MCP instructions (`extensions.mcp.instructions`) actively teach the anti-pattern: on an `if_match` mismatch, they tell the model to "call `get_item` again for the current if_match value and retry" — exactly the same-call self-comparison this issue is about, presented as the recommended recovery path.

Neither tool has any way to know or enforce *when* in the workflow it was actually called.

**The gap that specifically needs closing** (clarified during PRD review): even once `get_item` is called at the *start* of an edit workflow rather than right before the write, there is still a window — potentially minutes long, if the model reads content, then does other work before circling back to write it — between the original content read and that early `get_item` call. If the file is modified externally during that window, `get_item` today has no way to notice; it just hands back whatever etag is currently on the server, and `if_match` will happily match it, because `if_match` only ever protects the narrower window between `get_item` and the write itself. `get_item` needs to be able to say, at the moment it's called: "there is no point handing you this etag — the file already changed since you read it."

---

## 2. Goals

1. Rewrite this server's MCP-level instructions (`host.json`'s `extensions.mcp.instructions`) and `get_item`'s own tool description so that they establish — and no longer contradict — the correct workflow: capture the file's last-known-modified state (or, failing that, the wall-clock time) at the moment content is read, call `get_item` concurrently with that read rather than immediately before the write, and carry the resulting `if_match` value through to `update_file_content` unchanged.
2. `update_file_content`'s `if_match` parameter description explicitly warns against re-fetching a fresh `if_match` value immediately before the call, since doing so defeats the concurrency check it exists to provide.
3. `get_item` **requires** a `not_modified_since` timestamp — the tool call fails validation if it's omitted, so the model can no longer call `get_item` without committing to a comparison point. `get_item` compares the resolved item's Graph `lastModifiedDateTime` against the supplied value; if the item was modified *after* that timestamp, `get_item` fails with a distinct, actionable error instead of returning an etag, and the caller is expected to re-read the file's content and resolve the conflict rather than proceed with a stale edit.

---

## 3. Non-goals

- **Changing `if_match`'s own mechanism.** Graph's `If-Match` enforcement on the content `PUT` is untouched; this PRD only changes when and how the model is guided to capture and carry that value, plus adds the independent `not_modified_since` check ahead of it.
- **Cryptographic/token-based enforcement of the read → `get_item` → write sequence.** Issue #17's own "v2 hardening" note raises a signed token `get_item` could mint (reusing `Sapidus.Writeback.Shared`'s `ConfirmationTokenService`, the same primitive `delete_item` uses to force its two-call round trip) that `update_file_content` would have to echo back. Considered here and deliberately not adopted for this pass — see §5.3. A future PRD could revisit it if the trust-based design below proves insufficient in practice.
- **Changing `get_item`'s use by `delete_item`, `rename_item`, or `move_item`.** None of those flows has a prior "content read" to anchor a `not_modified_since` value to. They're unaffected because they never go through the `get_item` MCP tool at all — they call `DriveGraphClient`'s `GetItemByIdAsync`/`GetItemByPathAsync`/`GetItemAsync` directly as an internal implementation detail (see §4.1's tool-boundary-vs-client-method distinction), and that underlying method keeps the parameter optional/nullable. The requirement in Goal 3 is enforced only at the `get_item` MCP tool's own schema.
- **Applying this to `outlook-writeback`.** Issue #17's own Scope section flags this as "worth checking," not decided — `outlook-writeback` has no `get_item`/`if_match` pattern today, so there's nothing directly analogous to fix there yet. Out of scope for this PRD.
- **A combined read+`get_item` tool.** Issue #17's v2 note also floats this; not pursued here for the same proportionality reason as the token option (§5.3).
- **SharePoint-specific `lastModifiedDateTime`/versioning semantics.** SharePoint document library drives are already rejected outright at the `ResolveDriveIdAsync` layer (`docs/archive/PRD-drive-write.md` §11) — this PRD's design only needs to hold for OneDrive.

---

## 4. Design

### 4.1 `not_modified_since` on `get_item`

**Signature change:**

```
get_item(
    path_or_id: string,             // unchanged
    drive_id?: string,              // unchanged
    not_modified_since: string,     // NEW — ISO 8601 timestamp, REQUIRED
)
```

**Behavior:** the requirement lives at the tool boundary, not the Graph client. `Functions/GetItemTool.cs`'s `McpToolProperty` for `not_modified_since` is marked required — an MCP `tools/call` for `get_item` that omits it fails fast with a missing-required-input error, before any Graph call is made. `DriveGraphClient.GetItemByIdAsync`/`GetItemByPathAsync` (and the `GetItemAsync` discriminator wrapper that calls them) keep the parameter **optional/nullable** at the method signature level, so `delete_item`/`rename_item`/`move_item`'s internal calls into the same client methods (§3) are unaffected and continue to pass no value. When a value is supplied — always, from the `get_item` tool; never, from those internal call sites — the existing Graph `GetAsync` call resolves the item as today, and if the item's `LastModifiedDateTime` is strictly after the supplied timestamp, the call throws a new, distinct exception — `ItemModifiedSinceReadException` or similar, mirroring the existing `SharePointDriveNotSupportedException` pattern already established in `DriveGraphClient.cs` — instead of returning the item. `GetItemTool.RunAsync` surfaces this as a clear, actionable error message (not a generic Graph fault): the file changed since the caller's read, re-read its content and resolve any conflict before retrying. If the item was not modified after the supplied timestamp, `get_item` behaves exactly as today, no change to its response shape.

**Where the comparison happens:** entirely client-side, in `DriveGraphClient`, after the normal Graph `GetAsync`. Graph itself has no server-side equivalent of an HTTP `If-Unmodified-Since` conditional `GET` for `driveItem` — this is not a Graph query parameter, the same way `create_file`'s `conflict_behavior` is enforced client-side rather than via a Graph capability (`docs/archive/PRD-drive-write.md` §7).

### 4.2 Timestamp source

Because `get_item` now requires a value on every call (§2 Goal 3), the model must resolve one of the following two sources every time it calls this tool, not just when it happens to remember to. Two sources are available to the caller, in preference order:

1. **The file's own last-known-modified timestamp**, as surfaced by whatever tool the host used to read the file's content (e.g. the M365 connector's read tool, if its output includes modification metadata). This is the same clock Graph itself stamps `lastModifiedDateTime` with, so the comparison in §4.1 is exact — no skew.
2. **The host's own wall-clock time at the moment the content read occurred**, used only when the read step's output doesn't surface the file's own metadata. This crosses two independent clocks (the host's and Graph's), and is a materially weaker check: a few seconds of drift in either direction can either mask a real conflict (host clock fast) or cause a spurious "modified since" failure and re-read loop (host clock slow). The PRD's instructions text (§4.4) must state this tradeoff plainly rather than presenting both sources as equivalent.

**Verification item — confirmed during PRD review (2026-09-03):** the M365 connector's `sharepoint_search` tool returns `lastModifiedDateTime` as one of its result properties, so source 1 above is genuinely available in practice for the OneDrive-file-read case this PRD cares about, not merely a best-case fallback. Source 2 (wall-clock time) remains available for any read path that doesn't surface this field, but is no longer assumed to be the common case.

### 4.3 Trust model, and why a token wasn't chosen

`not_modified_since` is a value the model is trusted to carry forward honestly from its own read step — nothing stops the model from omitting it, or from passing a freshly-computed "now" instead of its true original read-time value, in which case the check degenerates to no protection at all. This is a deliberate, not accidental, design choice: it mirrors the trust level this server already places in `if_match` itself (nothing cryptographically stops the model from re-fetching `if_match` right before the write either — that's the whole problem this PRD is about) and in `delete_item`'s `expected_name` (a drift guard, not a proof of intent). Issue #17's v2 note proposes a stronger alternative — a signed token `get_item` mints that `update_file_content` must echo, reusing `ConfirmationTokenService`'s HMAC pattern from `delete_item` — but that adds real machinery (server-side state or a second signed value threaded through two tool calls) to close a gap that, at this proportionality level, is no worse than the gap this server already accepts elsewhere. Revisit only if trust-based `not_modified_since` proves insufficient in practice (§8, open questions).

Making the parameter required (§2 Goal 3) ensures a value is always *supplied*; it does not ensure the value is *truthful*. A model could satisfy the schema with a trivial always-passing value (e.g. an epoch timestamp) or a trivial value chosen to dodge the check (e.g. the current wall-clock time at the moment `get_item` happens to be called, which is exactly the anti-pattern this PRD exists to close) without violating the required-input constraint. "Required" closes the omission gap — the model can no longer forget or skip the parameter entirely — but restates, rather than removes, the trust boundary described above.

### 4.4 Reconciliation with issue #17's own rejection note

Issue #17's "Proposed fix" section states: *"a `lastModifiedDateTime` staleness check inside `get_item` was considered and rejected — it has the same 'always compares server against itself' flaw as a same-call etag check … timestamp is at best a weaker early sanity check (second-resolution) usable before `get_item` is even called."*

This PRD's `not_modified_since` design is the check that note describes as rejected, and needs to be read against it explicitly rather than silently adopted:

- The rejected version the issue describes is a **same-call** check — `get_item` deriving both sides of the comparison from its own single Graph response, which is trivially always "fresh" (comparing the server's state against itself, worded differently).
- This PRD's version compares against a value **supplied externally by the caller from a point in time strictly before `get_item` is invoked** — the host's own content-read timestamp, captured independently of this Graph call. That is not self-comparison: it can genuinely fail, and does so exactly when the file changed in the window between the host's read and this `get_item` call.
- What this closes, that `if_match` structurally cannot: `if_match` only protects the window between `get_item` and the write. It has no visibility into whether the file had *already* drifted from what the host believes it read, in the window between the original content read and a `get_item` call that might happen much later (the "read, sit idle for five minutes, file changes underneath, then call `get_item`" scenario this PRD is written to close).
- What this PRD concedes, matching the issue's own characterization: as a check, it is weaker than `if_match` — it's a second-resolution, client-supplied-value comparison, not a Graph-enforced atomic guarantee, and its usefulness depends on the caller supplying a truthful timestamp (§4.3). It is additive to `if_match`, not a replacement for it.

### 4.5 Doc strings and instructions to rewrite

**`host.json`'s `extensions.mcp.instructions`** — current text (excerpt) reads:

> "...get_item fetches metadata (size, folder/file, web URL, and an if_match value) for a drive-relative path or item ID - it is a pre-write safety check, not a browsing tool..."

and, on the `update_file_content` failure path:

> "...call `get_item` again for the current if_match value and retry."

Both need to change. Replacement establishes: capture the file's last-modified timestamp (or wall-clock time, per §4.2) when content is read; call `get_item` concurrently with that read, passing the now-required `not_modified_since`; carry the returned `if_match` unchanged through to `update_file_content`; on an `if_match` mismatch or a `not_modified_since` rejection, re-read the file's content and start the sequence over — never re-fetch `if_match` alone and retry.

**`Functions/GetItemTool.cs`** — the `McpToolTrigger` description carries the same "pre-write safety check" phrasing and needs the same rewrite; add the `not_modified_since` `McpToolProperty` marked required, with a description covering both timestamp sources (§4.2), the skew caveat for the wall-clock fallback, and a statement that `get_item` is scoped to edit workflows only — pure inspection with no prior content read is not a supported use (§5); the success-path response string is unaffected, but a new response/error path for the `not_modified_since` rejection needs its own clear wording, as does the new missing-required-input validation error.

**`Functions/UpdateFileContentTool.cs`** — the `if_match` `McpToolProperty` description gets an explicit added sentence warning against re-fetching a fresh `if_match` value immediately before this call, naming that as the specific anti-pattern that defeats the concurrency check.

### 4.6 Audit logging on `get_item`

**Today, `get_item` (`Functions/GetItemTool.cs`) logs nothing** — no `ILogger` usage at all, confirmed by inspection. This is a gap for this PRD specifically: without it, there's no way to later evaluate whether the trust-based design (§4.3) is actually holding up under real usage, which is exactly what open questions §8 items 4 and 6 need in order to move from "worth revisiting" to an actual decision.

`GetItemTool` gets one structured `LogInformation` call per invocation, added at the tool boundary (mirroring `DriveWriteService`'s existing per-mutation audit log — PRD-drive-write.md §9 item 7 — since `get_item` has no analogous read-side service layer to put it in instead). Every call logs: the resolved `path_or_id`/`driveId`, the caller-supplied `not_modified_since`, the resolved Graph `LastModifiedDateTime`, and the outcome (`pass` or `rejected`). No file content or other sensitive payload is logged, consistent with the root `CLAUDE.md`'s "tool invocation metadata only" logging rule. This flows to Application Insights the same way `DriveWriteService`'s existing logs do (`Program.cs`'s `AddApplicationInsightsTelemetryWorkerService`/`ConfigureFunctionsApplicationInsights`).

This doesn't change this PRD's trust model (§4.3) — nothing here validates that a supplied `not_modified_since` is truthful at call time. What it buys is an evidence trail: once this ships and runs against real usage, Application Insights can be queried for things like how often the `rejected` outcome actually fires, or whether logged `not_modified_since` values cluster suspiciously close to the logged call timestamp (a sign of the wall-clock-right-now anti-pattern surviving despite the parameter being required). That's the concrete mechanism §8 items 4 and 6 now point to instead of staying open with no way to ever close them.

---

## 5. Behavior & edge cases

- **Comparison semantics:** strict "after" only. `LastModifiedDateTime == not_modified_since` passes (no modification detected); only `LastModifiedDateTime > not_modified_since` fails. Equality is the common case when the timestamp source is the file's own metadata (§4.2 source 1) and nothing changed.
- **Precision:** Graph's `lastModifiedDateTime` is second-resolution (matching the issue's own note). Sub-second races are not distinguishable and are not a goal of this check.
- **Missing parameter on the `get_item` tool itself:** fails fast with a missing-required-input validation error, before any Graph call is made — the model cannot call `get_item` without committing to a comparison timestamp.
- **`delete_item`/`rename_item`/`move_item`'s internal calls:** unaffected — they call `DriveGraphClient`'s methods directly rather than the `get_item` MCP tool, and those methods keep the parameter optional/nullable (§3, §4.1).
- **Pure inspection calls with no prior read:** resolved during PRD review — `get_item` is edit-workflow-only, consistent with its existing "not a browsing tool" framing (§1). `drive-writeback` isn't here to serve a host's need to monitor a file being updated; a model with no intent to write, and thus no natural anchor timestamp, is not a supported caller of this tool. No sentinel/escape-hatch value is defined for this case (see former open question 5, §8).
- **Malformed timestamp:** an unparsable `not_modified_since` value fails fast with a clear input-validation error, not a silent skip of the check.
- **Wall-clock fallback skew:** documented in §4.2 as a known-weaker mode, not treated as equivalent to the file-metadata source. No attempt is made in this design to detect or compensate for clock skew (e.g. no grace-period tolerance) — a strict comparison is used deliberately, consistent with this server's existing fail-closed posture (e.g. `ResolveDriveIdAsync`'s SharePoint rejection).
- **Interaction with dry-run mode:** none — `get_item` is a read tool, unaffected by `DRIVE_WRITEBACK_DRY_RUN`.

---

## 6. Testing plan

All new logic is a client-side comparison added to `DriveGraphClient`, testable entirely offline against `StubHttpMessageHandler` (the existing test-project stub), consistent with this server's offline-first tiering (`dotnet test --filter "Category!=E2E"`, see `drive-writeback/CLAUDE.md`):

- Unit/integration coverage for the three-way boundary: item modified before, exactly at, and after `not_modified_since` — only the "after" case fails.
- `not_modified_since` omitted from a `get_item` tool call → missing-required-input validation error, verified at the `GetItemTool` level, before any Graph call.
- `DriveGraphClient.GetItemByIdAsync`/`GetItemByPathAsync` called directly with no value (mirroring `delete_item`/`rename_item`/`move_item`'s internal usage) → unchanged existing-test behavior at the client-method level (regression guard).
- Malformed `not_modified_since` input → validation error, not silently ignored.
- `GetItemTool` response text for both the success path and the new rejection path.
- §4.6's audit log entry fires with the expected fields (path/id, `driveId`, supplied `not_modified_since`, resolved `LastModifiedDateTime`, outcome) on both the pass and rejected paths — verified with `RecordingLogger<T>` (`DriveWriteback.Graph.Tests/TestSupport/RecordingLogger.cs`), the same test helper already used to assert `DriveWriteService`'s per-mutation log entries.

No new E2E test is strictly required for this pass — the comparison logic doesn't depend on any live-tenant-only Graph behavior — but one could optionally confirm that a real Graph `lastModifiedDateTime` value updates promptly enough after a content edit for the check to be reliable in practice.

---

## 7. Success criteria

- `host.json`'s instructions and `GetItemTool.cs`'s trigger description no longer describe `get_item` as a call to make "right before writing," and no longer instruct re-fetching `if_match` alone as the recovery path on a mismatch.
- `UpdateFileContentTool.cs`'s `if_match` description explicitly warns against re-fetching immediately before the call.
- `get_item`'s tool schema marks `not_modified_since` required, and an omitted value fails validation before any Graph call — verified by offline test.
- `get_item` rejects a read whose target was modified after a supplied `not_modified_since`, with a distinct, actionable error — verified by offline test.
- `delete_item`/`rename_item`/`move_item` are unaffected (no `not_modified_since` usage, no behavior change).
- `get_item` logs a structured audit entry (§4.6) on every call, flowing to Application Insights the same way `DriveWriteService`'s existing mutation logs do — verified by offline test.
- `dotnet test --filter "Category!=E2E"` green with the new coverage added.
- Issue #17 closed.

---

## 8. Open questions

1. ~~Does the M365 connector's read tool output actually surface a file's `lastModifiedDateTime`?~~ **Resolved (2026-09-03):** yes — the M365 connector's `sharepoint_search` tool returns `lastModifiedDateTime` as one of its result properties. Source 1 in §4.2 is confirmed available in practice; see §4.2's updated verification note.
2. ~~Final parameter name.~~ **Resolved:** keep `not_modified_since`.
3. ~~Should this same pattern eventually apply to `outlook-writeback`?~~ **Resolved:** no — different concern, not pursued as a follow-on from this PRD. §3's non-goal stands as originally written.
4. If the trust-based design (§4.3) proves insufficient in observed use (models fabricating `not_modified_since`), revisit the token-based alternative issue #17 itself raises. Still open, but no longer unanswerable in principle: §4.6's new audit logging records the supplied `not_modified_since`, the resolved `LastModifiedDateTime`, and the pass/rejected outcome for every `get_item` call, giving this a real evidence source in Application Insights once the server has run against production usage, rather than staying a judgment call with nothing to check it against.
5. ~~What should a model pass for `not_modified_since` when calling `get_item` purely for inspection, with no prior content read to anchor to?~~ **Resolved:** `get_item` is edit-workflow-only — pure inspection is not a supported use case; `drive-writeback` isn't here to serve a host's need to monitor a file being updated. No sentinel/escape-hatch value is defined. See §5's updated edge case and §4.5's doc-string guidance.
6. Does requiring the parameter on every call meaningfully reduce how often a model calls `get_item` right before writing with no real anchor timestamp, given nothing stops it from passing the current wall-clock time as a trivially-satisfying value (§4.3)? Still open, same answer as item 4 — §4.6's logging is the mechanism that would let this be revisited with real usage data (e.g. supplied `not_modified_since` values sitting suspiciously close to the logged call timestamp) instead of guesswork.

---

## 9. References

- [Issue #17](https://github.com/keithmendozasr/sapidus-writeback-mcp/issues/17) — this PRD's originating issue.
- Microsoft Graph — driveItem resource (`lastModifiedDateTime`): https://learn.microsoft.com/en-us/graph/api/resources/driveitem
- Microsoft Graph — upload or replace driveItem content (`If-Match`): https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
- Repo-internal: `docs/archive/PRD-drive-write.md` (§4.4, §7, §11 — `if_match`/`ConfirmationTokenService`/connector-output spike this PRD builds on), `drive-writeback/CLAUDE.md`.
