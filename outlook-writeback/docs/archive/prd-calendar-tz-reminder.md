# PRD: Calendar Event Timezone + Reminder Support

**Owner:** Keith Mendoza
**Status:** Completed
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, Claude Code CLI
**Repo location:** `sapidus-writeback-mcp/outlook-writeback/` (parent PRD: `docs/archive/prd-create-update.md`)

---

## 1. Problem Statement

`create_event` and `update_event` accept `start`/`end` as ISO 8601 strings with a UTC offset, which correctly pins the event's absolute instant — but `OutlookGraphClient.ToGraphDateTime` always sends Graph `TimeZone = "UTC"` regardless of that offset. The instant is right; the *displayed* time zone in Outlook is not, unless the mailbox's default zone happens to be UTC. Separately, neither tool exposes Graph's `isReminderOn`/`reminderMinutesBeforeStart` fields at all, so every created or updated event is stuck with whatever reminder state the mailbox's implicit default (or the event's prior state) leaves it at.

## 2. Goals

- Require an explicit display timezone on `create_event`, so new events are created showing the correct local time from the start.
- Keep the timezone optional on `update_event` (most updates aren't about the timezone), but validate it and reject it if passed without `start`/`end` (a timezone with no time change is a no-op that would otherwise fail silently).
- Let both tools optionally set a reminder (minutes-before-start).
- Keep `update_event` calls that don't pass a timezone byte-for-byte backward compatible with today's wire format.

## 3. Non-Goals

- No way to explicitly disable an existing reminder — only set/change one to a given minute value.
- No independent per-side (start vs. end) timezone — one `timeZone` value applies to whichever of start/end is being changed.
- No timezone validation beyond what `TimeZoneInfo.FindSystemTimeZoneById` already provides (accepts both IANA and Windows time zone names cross-platform since .NET 6).

## 4. Functional Requirements

### `create_event`
- **New input:** `timeZone` (required) — IANA (e.g. `America/New_York`) or Windows (e.g. `Eastern Standard Time`) time zone identifier. Controls how `start`/`end` are displayed on the calendar.
- **New input:** `reminderMinutes` (optional) — minutes before start to show a reminder (0 = at start time). Omit to leave reminders at the mailbox/Graph default.

### `update_event`
- **New input:** `timeZone` (optional) — same format as above. Only takes effect together with `start` and/or `end`; passing it alone is rejected with a clear error rather than silently doing nothing. If only one of `start`/`end` changes, only that side's timezone label updates — the other keeps its previous value (a known, accepted display-label edge case, not treated as a data-integrity bug).
- **New input:** `reminderMinutes` (optional) — same semantics as create; omitting it leaves the event's existing reminder state untouched (Graph PATCH only touches fields present in the payload).

## 5. Design Notes

**`ToGraphDateTime` dual-mode conversion.** To keep `update_event` callers who never pass a timezone completely unaffected, `ToGraphDateTime` branches on whether a timezone was supplied:

- No timezone → today's exact behavior: `TimeZone = "UTC"`, `DateTime = value.UtcDateTime.ToString("o")` (carries a trailing `Z`).
- Timezone supplied → `TimeZoneInfo.FindSystemTimeZoneById` validates the id (letting `TimeZoneNotFoundException` bubble up unwrapped as the input-validation error), then `TimeZoneInfo.ConvertTime(value, timeZone)` resolves the correct local wall-clock time for that specific instant (so DST is handled per-instant, not via a fixed offset), formatted with no trailing `Z`/offset — the format Graph expects when a named zone is given.

This means passing `timeZone: "UTC"` explicitly (always the case for `create_event`, since it's required there) produces a *different* wire format (no `Z`) than the implicit-UTC legacy path used by timezone-less `update_event` calls. Both are valid per Graph's `dateTimeTimeZone` contract; the asymmetry is intentional and documented in code so a future reader doesn't "fix" it into consistency and reintroduce a wire-format regression on existing `update_event` callers.

## 6. Milestone

Single phase — this is a parameter-level addition to two already-shipped tools, not a new server or new tool surface. Done when the implementation lands and `dotnet test --filter "Category!=E2E"` is green.
