# PRD: Recurring Events in `create_event` and `update_event`

**Owner:** Keith Mendoza
**Status:** Implemented — pending deployment
**Target platform:** Azure
**Language:** C#
**Target clients:** Claude Desktop, Claude Cowork, Claude Code CLI
**Repo location:** `sapidus-writeback-mcp/outlook-writeback/` (parent PRDs: `docs/archive/prd-create-update.md`, `docs/archive/prd-calendar-tz-reminder.md`)
**Issue:** [#16](https://github.com/keithmendozasr/sapidus-writeback-mcp/issues/16)

---

## 1. Problem Statement

`create_event` creates only a single, non-repeating event, and `update_event` can't touch recurrence. A recurring appointment (e.g. weekly PT) has to be added by hand in Outlook afterward, or created as N separate events that can't be edited or cancelled as one series; an existing series' schedule can't be adjusted through the tools either.

## 2. Goals

- Let `create_event` create a recurring series in one call: daily, weekly, monthly, or yearly; an interval; days of the week; and an end condition (occurrence count, an until-date, or none).
- Support both absolute monthly/yearly patterns ("the 15th of every month") and relative ones ("second Tuesday of every month", "first Friday in September every year"). Graph supports both (`absoluteMonthly`/`relativeMonthly`/`absoluteYearly`/`relativeYearly` pattern types).
- Let `update_event` replace an existing event's recurrence with the same parameters (e.g. change a weekly series to every other week, or add an end date).
- Keep every existing `create_event` and `update_event` call byte-for-byte unchanged: all new parameters are optional and trailing.

## 3. Non-Goals

- **No way to remove recurrence** (turn a series back into a single event). Graph's update-event sample includes `"recurrence": null`, but its effect is not documented and the Graph .NET SDK omits null properties, so it can't be shipped untested. Follow-up once it can be checked against live Graph (delete + recreate is the workaround).
- `update_event` edits the series master's recurrence only; it doesn't edit individual occurrences or exceptions.
- No per-occurrence exceptions, and no "every weekday" shorthand (use weekly with Mon–Fri).
- Relative monthly/yearly patterns take exactly one weekday (Graph's "first matching day" semantics for several days is confusing); weekly takes any number.
- No deletion/preview changes: `delete_event` on a series occurrence is not touched here.

## 4. Functional Requirements

New optional `create_event` inputs, flat scalars (the MCP extension binds primitives, `int?` and `string[]?` reliably; a nested object is unverified):

| Parameter | Meaning |
|---|---|
| `recurrenceType` | `daily`, `weekly`, `monthly`, `yearly`. Required if any other `recurrence*` input is set. `start`/`end` describe the first occurrence. |
| `recurrenceInterval` | Every N units; default 1. |
| `recurrenceDaysOfWeek` | Day names. Required for weekly (start date must fall on one of them); exactly one for relative monthly/yearly; rejected otherwise. |
| `recurrenceWeekIndex` | `first`/`second`/`third`/`fourth`/`last`. Switches monthly/yearly to relative. |
| `recurrenceMonth` | 1–12. Yearly + `recurrenceWeekIndex` only; defaults to the start date's month. |
| `recurrenceCount` | End after N occurrences. |
| `recurrenceUntil` | End on/before `yyyy-MM-dd` in the event's `timeZone`. |

Neither `recurrenceCount` nor `recurrenceUntil` means no end date; both is an error.

`update_event` takes the same seven inputs. Setting `recurrenceType` there **replaces** the event's recurrence and requires `start` and `timeZone` (Graph requires the range's `startDate` to equal the event's start, and the first-occurrence validation needs the date and zone); pass the series' current first-occurrence start (readable via the M365 connector) when the schedule isn't moving. `end` stays optional. Calling it on a non-recurring event is expected to turn it into a series but is unverified. With no `recurrence*` inputs the PATCH payload is unchanged. Updating a series master that has separately edited occurrences sends extra notification emails and Outlook may reject with `ErrorOccurrenceCrossingBoundary` (Graph docs).

## 5. Design Notes

- `OutlookWriteback.Graph/RecurrenceSpec.cs` is a Graph-agnostic, validated description built by `RecurrenceSpec.Create` (shape checks that don't need the start date) and `ValidateAgainstStart` (checks the first occurrence is a member of the series: weekly weekday in the list; relative patterns' start is that Nth weekday, and in that month for yearly; `until` not before start). Failing loudly beats Graph silently shifting the first occurrence.
- `OutlookGraphClient.BuildRecurrence` maps the spec to `PatternedRecurrence`, setting only the pattern properties the type needs (Graph errors on extras). Weekly sets `firstDayOfWeek: sunday` (Graph default; the docs list it as required for weekly).
- The range's `startDate` is the start's calendar date **in `timeZone`** (not the UTC date) and `recurrenceTimeZone` is `timeZone`; an end date is interpreted in the same zone. `ToGraphDateTime`'s dual-mode behavior is untouched.
- Absolute monthly on day 29–31: Graph decides how short months are handled; not overridden here.
- Success text becomes `Recurring event series created. ID: ...` for a series on create (the ID is the series master's); `update_event` keeps `Event updated. ID: ...`.
- **Input errors are returned as text, not thrown.** A thrown exception reaches the calling agent only as a generic failure (the message stays in the server log), so `create_event`/`update_event` catch input problems (`ArgumentException`) and return `No event was created. The input was rejected: <what to fix>` (`No changes were made.` for update), as `delete_event` already does. This covers `start`/`end` (strict ISO 8601 with an explicit offset: `Functions/ToolInput.cs` — an offset-less or free-form value such as `10am` is rejected rather than silently read in the server's zone), unknown `timeZone`, `reminderMinutes`, blank attendees, and every `recurrence*` rule. `recurrenceUntil` must be exactly `yyyy-MM-dd`.
- Both tools share one set of parameter descriptions (`Functions/RecurrenceParameterDocs.cs`) and one parsing entry point (`RecurrenceSpec.CreateOptional`) so their schemas can't drift.

## 6. Milestone

Single phase. Done when the implementation lands and `dotnet test --filter "Category!=E2E"` is green (it is). Real-Graph confirmation (first-occurrence rules, day-31 monthly behavior) is still open — the E2E tier doesn't cover recurrence yet — so after deployment, create a short series via the tool and read it back with the M365 connector. Moves to `docs/archive/` as `Completed` once deployed.
