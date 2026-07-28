# PRD: Multi-Recipient Support for Outlook Draft Tools

**Owner:** Keith Mendoza
**Status:** Completed
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, Claude Code CLI
**Repo location:** `sapidus-writeback-mcp/outlook-writeback/` (one server folder in the `sapidus-writeback-mcp` monorepo — see `../REPO-CONVENTIONS.md` for the cross-server rules this PRD inherits)
**Builds on:** `docs/archive/prd-create-update.md` ("v1" below) — this PRD is additive to and supersedes the recipient-related portions of v1 §6 (`create_draft`/`update_draft`'s input shape). It also supersedes two adjacent pieces of v1 §6: both `update_event`'s and `create_event`'s `attendees` wire type (and, for `update_event`, its update semantics), since fixing the underlying null-vs-empty ambiguity for email recipients means fixing the same string-carrying-a-list representation problem for the calendar side too — see §1 and §6 below. Every other section of v1 is unchanged; those sections are noted below with a one-line pointer back rather than restated.

---

## 1. Problem Statement

v1 §6 described `create_draft`'s input as "to address(es)" — plural-tolerant wording from the start — but what actually got built is a single required `toAddress: string`, and `Cc`/`Bcc` were never mentioned anywhere in v1 at all, not even as a non-goal. In practice this doesn't hold up: real emails routinely go to more than one person, and often need a Cc or Bcc, so this is a felt, everyday limitation rather than a hypothetical gap. Multi-`To` is best understood as scope v1 already intended but never fully built; `Cc`/`Bcc` is genuinely new scope this PRD introduces.

The calendar side of this server already solved the adjacent multi-value problem, but with a workaround rather than a first-class solution: `create_event`/`update_event`'s `attendees` parameter is a single comma-separated string, not a native list — despite the underlying Azure Functions MCP extension (`Microsoft.Azure.Functions.Worker.Extensions.Mcp` 1.4.0, already referenced by this project per v1) supporting array-typed tool parameters natively; see §11 for how this was confirmed. This PRD does not extend the comma-separated workaround to email. It replaces the workaround entirely, for both email and calendar: `to`, `cc`, `bcc`, and `attendees` all become a nullable JSON array of address strings (`string[]?`) at the MCP tool boundary, not a string the server re-splits internally.

The same three-state ambiguity recurs one level down inside today's string-splitting workaround: `update_event`'s `attendees` can't distinguish "leave the attendee list alone" from "clear it," because a blank string collapses to the same value that means "no change." Moving to a real array type resolves this natively instead of faking it with a parsing convention — the three states become `null` (omitted), an empty array (clear), and a populated array (replace). This PRD's recipient-list work and that calendar gap are the same underlying representation problem — using a string to carry what is structurally a list — reached from two different tools, so fixing one without the other would leave the fix asymmetric for no good reason. This PRD closes both.

## 2. Goals

- Let `create_draft` and `update_draft` optionally accept `To` recipient, plus optional `Cc` and `Bcc` recipients, each expressed as a native JSON array of address strings (`string[]?`) at the MCP tool boundary — the same array-of-`attendees` fix this PRD also makes to `create_event` and `update_event`, replacing the comma-separated-string convention v1 used for `attendees`.
- **All three recipient fields are optional on create.** `create_draft`'s `to` drops its current `isRequired: true` schema constraint, matching `cc`/`bcc` — a draft may be created with zero recipients across all three lists (e.g., a placeholder to address later), expressed as an empty array or the field omitted entirely.
- **A uniform three-state update convention, shared by every list-valued field this server has, expressed natively through the array type rather than through string parsing.** On `update_draft` (`to`/`cc`/`bcc`) and `update_event` (`attendees`) alike: omitting the field (the JSON key is absent, so the Functions binding sees C# `null`) leaves the existing list unchanged; passing an explicit empty array (`[]`) clears the list entirely; passing an array of one or more addresses replaces the list entirely. This is a genuine behavior *and* wire-type change for `update_event`'s `attendees` (today it's a comma-separated string where blank is indistinguishable from omitted), adopted deliberately so both tools share one convention instead of diverging on it.
- Keep the existing Graph scopes (`Mail.ReadWrite`, `Calendars.ReadWrite`) — no new permission is needed, since `CcRecipients`/`BccRecipients` are already properties of the same `Message` resource `ToRecipients` lives on today, and the `attendees` change is a representation fix, not a new field.
- **A small shared helper for what still needs consolidating now that splitting itself is gone.** Even with a real array type, two things are still worth centralizing rather than duplicating across `create_draft`/`update_draft` (`to`/`cc`/`bcc`) and `update_event` (`attendees`): trimming whitespace around each entry a caller could still hand-construct inside an otherwise-valid array, and — for the email tools specifically — detecting an address that appears in more than one of `to`/`cc`/`bcc`. Deliberately **not** included: silently dropping blank entries. A blank string inside an otherwise-populated array must be rejected, not dropped — dropping it could turn a populated array into an empty one behind the caller's back, which is the same "blank silently becomes clear" bug this PRD exists to close, just moved one level down into array elements instead of a whole comma-separated string.
- **Deduplication across To/Cc/Bcc.** If the same address appears in more than one list, the caller must be asked to de-duplicate the recipients — detected by the shared helper above, not silently resolved by picking one list to keep the address in.

## 3. Non-Goals

- **Email-address format validation.** Nothing validates addresses today — not `toAddress`, not `attendees` — and this change doesn't add any. The shared helper only trims entries and rejects structurally blank ones (§6); it never inspects whether a non-blank entry actually looks like an email address. Graph itself is the thing that rejects a malformed address, at the API boundary. This is called out explicitly so a future reader doesn't assume the new helper does more than it does.
- **An attendee-type-style distinction for Cc/Bcc.** The calendar side has `Attendee.Type` (Required/Optional/Resource) as a single field with a type tag; email recipients don't work that way — To vs Cc vs Bcc is expressed as three separate Graph properties (`ToRecipients`, `CcRecipients`, `BccRecipients`), not a per-recipient type flag, and this PRD doesn't change that shape.
- **Multiple `Attendee.Type` values per recipient, or any other calendar-attendee feature beyond the clear/leave-alone/replace fix.** The `attendees` change in this PRD is scoped narrowly to closing the null-vs-empty gap already present in v1; it isn't a broader revisit of calendar-tool behavior.

## 4. Users

Unchanged from `docs/archive/prd-create-update.md` §4 — still the same single primary/only user.

## 5. Use Cases

Extends v1 §5's UC1: after reconciliation, Claude drafts a short, text-only email based on chat context — in practice this routinely needs more than one `To` recipient, and occasionally a `Cc` (e.g. looping in a manager) or a `Bcc`. That's the concrete, weekly-frequency scenario this PRD targets.

Also extends v1 §5's UC3 (update an existing calendar event): rescheduling or editing details sometimes means removing an attendee entirely rather than replacing the list with a smaller one you have to re-type — today that's not expressible via `update_event`. Nothing else in v1's use case table changes.

## 6. Functional Requirements — MCP Tools

### `create_draft`

- **Input:** `to` (renamed from `toAddress`, now genuinely optional and array-typed — drops the `isRequired: true` schema constraint `toAddress` has today), `subject`, `body`, optional `isHtml` flag (unchanged from v1), optional `cc` (array of strings, new), optional `bcc` (array of strings, new). All three recipient fields are a nullable JSON array of zero or more addresses — not a comma-separated string.
- **No minimum-recipient constraint.** `to`/`cc`/`bcc` may all be omitted or an empty array — a draft with zero recipients across all three lists is a legitimate result (e.g., a placeholder drafted before the recipient is decided). Graph itself accepts a `Message` with empty/absent recipient collections; nothing here needs to reject that.
- Everything else about `create_draft` (Graph call, scope, attachment rejection) is unchanged from v1 §6.

### `update_draft`

- **Input:** `to` (renamed from `toAddress`, optional, array of strings), `subject`, `body`, `isHtml` (unchanged), optional `cc` (new), optional `bcc` (new).
- **Update semantics — three states per list field**, shared with `update_event`'s `attendees` (see below): for each of `to`, `cc`, `bcc`, the array-typed parameter maps to exactly one of:
  - **omitted** (the parameter key isn't present in the call, so the Functions binding sees C# `null`) → leave the existing list unchanged;
  - **present as an explicit empty array** (`[]`) → clear the list entirely — Graph receives an explicit empty recipient collection for that field;
  - **present as an array of one or more addresses** → replace the list entirely with exactly those addresses (not merged with what's already there).

  This resolves what was Open Question 1 in earlier drafts of this PRD (§11): blank and omitted are no longer indistinguishable, and there's no string-parsing convention standing in for the distinction — the array type carries it directly.
- **Correctness fix required in `OutlookGraphClient.UpdateDraftAsync`.** Its current guard (`OutlookGraphClient.cs:82-83`) is `if (toAddress is null && subject is null && bodyText is null) throw new ArgumentException(...)`. As written, a call that only changes `cc` or `bcc` (leaving `to`, `subject`, and `body` all unset) would incorrectly throw. This guard must be extended to also check `cc`/`bcc`, so "at least one field is changing" correctly accounts for all five updatable fields. Note that under the new three-state semantics, a call that only clears `cc` (an explicitly-empty-array `cc` alongside everything else omitted) correctly does **not** throw once this fix lands, since an empty array is a non-null value, not `null`.

### `create_event`

- **Input shape change.** `attendees` changes from a single optional comma-separated `string` to a nullable array of strings (`string[]?`), matching `update_event`'s `attendees` and `create_draft`'s `to`/`cc`/`bcc` — no wire-type inconsistency remains across this server's list-valued fields.
- **No three-state "unchanged" semantics here** — `create_event` has no prior attendee list to preserve, so there's nothing for an omitted value to leave alone. Only two outcomes matter on create: omitted or empty ⇒ no attendees invited; a populated array ⇒ invite exactly those attendees. This mirrors why `create_draft`'s `to`/`cc`/`bcc` also have no "unchanged" state (§6 `create_draft` above) — the concept only applies where there's something to be left alone.
- **Omitted and an explicit empty array produce the same practical result but not an identical request payload.** `RecipientList.Normalize(null)` returns `null`, so `BuildEvent` never sets `Attendees` and the property is absent from the Graph request entirely; `RecipientList.Normalize([])` returns `[]`, so `BuildEvent` sets `Attendees` to an empty (non-null) list, which serializes as an explicit `"attendees":[]`. Both create an event with no attendees — the divergence is in what's sent over the wire, not in the outcome.

### `update_event`

- **Input shape change.** `attendees` changes from a single optional comma-separated `string` to a nullable array of strings (`string[]?`) — this is a real schema change to an already-shipped parameter, not just a semantics fix. See §13 Known Limitations for the compatibility note.
- **Update semantics — same three-state convention as `update_draft`** above: omitted `attendees` (`null`) leaves the event's attendee list unchanged; an explicit empty array (`[]`) clears it entirely; a non-empty array replaces it entirely. This is a **behavior change** from v1 — today, blank and omitted are both silent no-ops.
- No change to `delete_event` or any other calendar tool.

### Shared recipient-list helper (new)

- A single public helper — `OutlookWriteback.Graph.RecipientList` — replaces what would otherwise be duplicated defensive handling across the three array-typed parameters that need it (`create_draft`/`update_draft`'s `to`/`cc`/`bcc`, `update_event`'s `attendees`). Unlike the string-splitting helper an earlier draft of this PRD proposed, there is no parsing step here — the MCP tool binding is responsible for handing the Functions layer a real `string[]?` (or rejecting malformed input as a schema-validation error before the tool method is even invoked).
- `RecipientList.Normalize(IReadOnlyList<string>? raw) -> IReadOnlyList<string>?` — trims whitespace around each entry; passes `null` and `[]` through unchanged, since those two states are exactly the "unchanged" / "clear" signal the caller sent and must not be conflated with each other. **Throws `ArgumentException` if any entry is blank/whitespace-only after trimming** — a blank array element is rejected outright rather than silently dropped, precisely so a populated-but-partially-blank array can never quietly collapse into the "clear" state (`[]`). This is a structural check (is this element usable at all), not the email-format validation §3 excludes.
- `RecipientList.FindDuplicates(IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc) -> IReadOnlyList<string>` — implements §2's dedup-detection goal: returns any address present in more than one of the three lists, so `create_draft`/`update_draft` can reject the call and ask the caller to de-duplicate rather than silently keeping the address in only one list. Not used by `create_event`/`update_event` — `attendees` has no sibling lists to collide with.
- It must be **public**, not `internal` with `InternalsVisibleTo` (the pattern `BuildDraftMessage` uses today) — the `Functions` project, not just the test project, needs to call it directly.
- Lands in `OutlookWriteback.Graph`, alongside the rest of the Graph-facing logic, unit-tested in `OutlookWriteback.Graph.Tests` the same way `BuildDraftMessage`/`BuildEvent` already are. This closes the coverage gap for `create_event`'s and `update_event`'s `attendees` and all of `create_draft`/`update_draft`'s recipient handling — `create_event`'s previously inline, untested `Split` call is gone entirely, replaced by the same shared, tested helper the other three tools use.

### Graph client changes

- `BuildDraftMessage`/`BuildUpdateDraftMessage` (`OutlookGraphClient.cs:131-152`) change from a single `string toAddress` to `IEnumerable<string>? toAddresses` — nullable on **both** create and update now, since `to` is optional on create too — plus new `IEnumerable<string>? ccAddresses` and `IEnumerable<string>? bccAddresses` parameters.
  - `BuildDraftMessage` (create) always assigns all three recipient collections, coalescing a `null` input to an empty list first (`toAddresses ?? []`, etc.) — there's no prior state to preserve on create, so `null` and `[]` are handled identically.
  - `BuildUpdateDraftMessage` (update) keeps its existing `if (x is not null) message.X = ...` conditional-assignment pattern unchanged — that pattern already implements the three-state semantics correctly now that the Functions layer hands it a real `string[]?` instead of a value squeezed through comma-parsing. No new branching logic is needed here.
- A new `BuildRecipients(IEnumerable<string> addresses) -> List<Recipient>` helper — directly analogous to the existing `BuildAttendees` (`OutlookGraphClient.cs:181-186`) — replaces the current hardcoded one-element `ToRecipients = [new Recipient { EmailAddress = new EmailAddress { Address = toAddress } }]` literal, and is reused to populate `ToRecipients`, `CcRecipients`, and `BccRecipients`. Given an empty `addresses` enumerable it correctly returns an empty `List<Recipient>`, which is what makes the "clear" state work — no special-casing needed there either.
- `CreateDraftAsync`/`UpdateDraftAsync` (`OutlookGraphClient.cs:61-89`) gain `ccAddresses`/`bccAddresses` parameters (optional, defaulting to `null`), positioned after `isHtml` and before `cancellationToken`, mirroring how `CreateEventAsync`/`UpdateEventAsync` already place `attendeeAddresses`.
- **`BuildUpdateEvent` and `UpdateEventAsync` need no code changes at all.** `BuildUpdateEvent`'s existing `if (attendeeAddresses is not null) calendarEvent.Attendees = BuildAttendees(attendeeAddresses);` (`OutlookGraphClient.cs:213-214`) and `UpdateEventAsync`'s existing all-null guard (`OutlookGraphClient.cs:119-120`) were already structurally correct for the three-state convention — the only reason `update_event` couldn't clear attendees before was that `Functions/UpdateEventTool.cs`'s tool schema exposed `attendees` as a comma-separated string and split it inline, collapsing a blank string to `null` before it ever reached `OutlookGraphClient`. Once `attendees`'s MCP schema itself becomes a nullable array (§6 above), that collapsing bug has no path to occur — the Functions binding delivers `null`/`[]`/`[...]` directly, with no string in between to misparse.

## 7. Authentication & Authorization

Unchanged from `docs/archive/prd-create-update.md` §7a/§7b — no new Graph scope is needed (`Cc`/`Bcc` are already covered by the existing `Mail.ReadWrite` scope on the `Message` resource), and neither auth boundary (7a: MCP server → Graph, 7b: Claude → MCP server) is affected by this change.

## 8. Architecture Overview

Unchanged from `docs/archive/prd-create-update.md` §8.

## 9. Client Setup (per surface)

Unchanged from `docs/archive/prd-create-update.md` §9.

## 10. Non-Functional Requirements

Unchanged from `docs/archive/prd-create-update.md` §10, with one clarification: the existing "never log mailbox content" logging rule already covers recipient addresses, and that continues to apply to the new `cc`/`bcc` fields exactly as it applies to `to` today — no new logging is introduced by this change.

## 11. Open Questions

1. ~~How does one clear an existing `Cc` or `Bcc` on `update_draft`?~~ **Resolved.** §6 adopts a three-state convention (omitted = unchanged, empty array = clear, populated array = replace) expressed natively through an array-typed input, applied uniformly to `update_draft`'s `to`/`cc`/`bcc` and `update_event`'s `attendees` alike.
2. ~~Does the Azure Functions MCP extension (`Microsoft.Azure.Functions.Worker.Extensions.Mcp`, isolated worker model) support binding a native JSON array as an MCP tool's input parameter (e.g., `string[]?` on the C# method signature), or only primitive-typed parameters?~~ **Resolved.** Confirmed directly against the exact package version this project already references (1.4.0, `outlook-writeback/bin/Debug/net10.0`): its internal `TypeExtensions.MapToToolPropertyType`, invoked via reflection against `string[]`, `IEnumerable<string>`, and `List<string>`, maps each to `{ TypeName: "string", IsArray: true }` — i.e., a JSON Schema `{"type": "array", "items": {"type": "string"}}` property. Array-typed tool parameters are supported today; v1's `attendees` being a comma-separated string was a design choice in that PRD, not a platform limitation. No spike or fallback design is needed before implementation.

## 12. Milestones

- **Phase 1 — Multi-recipient support:** (Numbered from 1, not continuing v1's Phase 0–4 sequence — this is its own document with its own milestone list, not a later phase of the completed v1 PRD.)
  - Rename `toAddress` → `to` on both `create_draft` and `update_draft`, retype it from `string` to `string[]?`, drop its `isRequired: true` schema constraint so all three of `to`/`cc`/`bcc` are optional on create, and add optional `cc`/`bcc` (also `string[]?`) to both tools, per §6.
  - Retype `attendees` from `string` to `string[]?` in the MCP tool schema for **both** `create_event` and `update_event`, per §6 — `Functions/CreateEventTool.cs` drops its inline `Split` call in favor of `RecipientList.Normalize`, the same as `UpdateEventTool.cs`.
  - Add the shared `RecipientList` helper to `OutlookWriteback.Graph`, with unit tests covering: `Normalize` passing `null` through unchanged, passing `[]` through unchanged, trimming whitespace around entries, and throwing `ArgumentException` when a populated array contains a blank/whitespace-only entry; and `FindDuplicates` covering no-overlap, a single address duplicated across two lists, and an address duplicated across all three.
  - Retrofit `Functions/UpdateEventTool.cs` — its `attendees` binding changes from a comma-separated string parameter to `string[]?`, removing its inline `Split` call entirely (§6's Graph-client-changes note: no replacement parsing is needed, since the array arrives pre-split). Add a test proving an empty-array `attendees` now clears the list, where blank/omitted used to both be no-ops.
  - Update `OutlookWriteback.Graph.Tests/OutlookGraphClientPayloadTests.cs`'s existing `.Single()`-based assertions (lines 13-27 and 78-132) to list-based assertions, following the same pattern already used for attendees in `BuildEvent_maps_attendees_when_provided` (same file, lines 60-75).
  - Add new test cases for multi-address `to`, and for `cc`/`bcc` present/absent/changed, on both `BuildDraftMessage` and `BuildUpdateDraftMessage`, plus a zero-recipient `BuildDraftMessage` case (all three lists empty/absent).
  - Extend the `UpdateDraftAsync` "at least one field" guard (`OutlookGraphClient.cs:82-83`) to include `cc`/`bcc`, with tests proving: a `cc`-only (or `bcc`-only) update no longer throws, and a call passing only an empty-array `cc` (clearing it, everything else omitted) also doesn't throw.
  - Add a matching test for `BuildUpdateEvent`/`update_event` proving an empty-array `attendees` input clears `calendarEvent.Attendees` to an empty (not null) list, and that `UpdateEventAsync_throws_when_no_fields_are_provided` still passes unchanged (it already covers the true-omission case).
  - **Added beyond the original milestone list, at Keith's request after implementation.** `RecipientList.Normalize`'s blank-entry rejection and `create_draft`/`update_draft`'s duplicate-recipient rejection previously had no coverage above the `OutlookWriteback.Graph` layer — nothing exercised the actual `Functions/CreateDraftTool.cs`/`UpdateDraftTool.cs`/`UpdateEventTool.cs` classes where that validation is wired in and where the exceptions get thrown. Closed by adding a `ProjectReference` from `OutlookWriteback.Graph.Tests` to `OutlookWriteback.csproj` (the Functions project) and three new test files - `Functions/CreateDraftToolTests.cs`, `Functions/UpdateDraftToolTests.cs`, `Functions/UpdateEventToolTests.cs` - instantiating the tool classes directly (the `[Function]`/`[McpToolTrigger]` attributes are just metadata; the classes are plain C# and don't need the Functions host to invoke). Covers: blank-entry and duplicate-recipient rejection at the tool layer, whitespace trimming actually reaching the Graph request body, and the empty-array-clears-the-list behavior for `cc` and `attendees` through the tool layer specifically (not just at the `OutlookGraphClient` layer already covered above).
  - **Added during PR review, after `create_event`'s `attendees` was retyped to close the wire-type inconsistency.** A fourth test file, `Functions/CreateEventToolTests.cs`, covers the same blank-entry rejection and whitespace-trimming behavior for `create_event` that the other three tool-test files cover for their own tools, plus a no-attendees baseline case. A new `OutlookGraphClientPayloadTests.cs` case, `BuildEvent_sets_attendees_to_an_empty_not_null_list_when_an_empty_array_is_provided`, covers the create path reaching `BuildEvent` with an explicit `[]` for the first time (previously unreachable, since `CreateEventTool` always collapsed blank/absent input to `null` before this change).

## 13. Known Limitations

Unchanged from `docs/archive/prd-create-update.md` §13 (the Claude-side OAuth refresh gap, #228, is unrelated to this change).

**New limitation introduced by this PRD — `update_event`'s `attendees` breaks, not just changes meaning, for existing callers.** Today, `attendees` is a comma-separated string on `update_event`, and a blank one is a silent no-op. After this PRD ships, `attendees` is array-typed, so any existing caller (chat context, saved prompts, muscle memory) still passing a comma-separated string will fail MCP schema validation outright rather than silently misbehaving — the call errors instead of quietly clearing or no-op'ing. That's arguably a safer failure mode than v1's original silent-no-op ambiguity, but it's still a hard break with no deprecation window: the fix requires the reinterpretation and the retype together. Worth a one-line callout in the tool's MCP description (`Functions/UpdateEventTool.cs`) so a calling model doesn't assume the old string shape still works. Little is actually lost here in practice, since a blank comma-separated `attendees` was already a no-op — an existing caller passing a *populated* comma-separated string, though, does lose that call outright, same as below.

**New limitation — `create_event`'s `attendees` breaks the same way, but with more at stake.** `create_event`'s comma-separated `attendees` was fully functional (unlike `update_event`'s, where only the no-op case was affected) — a caller passing `"alice@example.com,bob@example.com"` today successfully invites both attendees. After this PRD ships, that same call fails MCP schema validation outright, since `attendees` is now `string[]?`. Existing callers relying on the comma-separated form lose working functionality, not just an ambiguous no-op — worth the same one-line callout in `Functions/CreateEventTool.cs`'s MCP description.

**Discovered during implementation — validation error messages don't reach the caller.** `RecipientList.Normalize`'s blank-entry rejection and `create_draft`/`update_draft`'s duplicate-recipient rejection both throw `ArgumentException` with a specific, actionable message. Confirmed empirically against a local `func start` instance (raw `tools/call` requests with a blank `to` entry and with the same address in both `to` and `cc`): the Azure Functions MCP extension does not propagate that message through the JSON-RPC response — the caller sees only `{"isError": true, "content": [{"text": "An error occurred invoking 'create_draft'."}]}`, with the exception's actual message discarded somewhere in the isolated-worker-to-host boundary. This is not a regression introduced by this PRD: the same probe against `update_draft`'s pre-existing (v1) "at least one field" guard produces an identical generic message, so the opacity is a pre-existing platform characteristic of this Functions MCP extension version, not something new validation logic broke. Mitigation adopted: each affected tool's MCP description now tells the calling model that a generic failure likely means a blank or duplicated recipient entry, so it has a chance to self-correct without seeing the specific reason. A real fix (if one exists) would require investigating the extension's exception-propagation path — out of scope here.

## 14. Future Consideration: Auth Pattern for Additional MCP Servers

Unchanged from `docs/archive/prd-create-update.md` §14 — the per-server Entra app invariant is not affected by this change.
