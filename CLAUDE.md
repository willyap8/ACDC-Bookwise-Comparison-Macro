# CLAUDE.md — Bookwise Daily Comparison Macro

## Project context

Part of the **ACDC Data Migration Project**: migrating clinical scheduling data from legacy **Bookwise** and **iPM** systems into **Oracle Millennium EMR** across three hospital sites — Box Hill, Maroondah, Yarra Ranges.

This repo is one of three sibling tools in the macro suite. Each is managed in its own GitHub repo:
- `Bookwise Daily Comparison Macro` (this repo)
- `EMRComparisonMacro` (separate GitHub repo)
- `iPM Daily Comparison Macro` (separate GitHub repo)

Deployed as an Excel Add-In (`.xlam`), launched via a Quick Access Toolbar (QAT) button, used by **non-technical clinical and booking staff** in an M365/SharePoint environment.

## What this macro does

Compares a **MASTER Bookwise sheet** against **daily Bookwise exports**, flagging discrepancies for review.

Current version: **v3.1**.

## Non-negotiable rules — data integrity

These override any other instruction or convenience shortcut:

- Never modify source files (the daily export). Source files are opened **read-only**.
- Hard stops fire *before* any changes to the MASTER begin if pre-conditions aren't met (see Hard Stops below) — never partially process then fail.
- Prefer local copies over running directly against live SharePoint-hosted files (sync conflict risk).
- Never parse a **text date** with `CDate` — its day/month interpretation is locale-dependent and silently flips `dd/mm` ↔ `mm/dd`. Text dates must be parsed explicitly in Australian day-first order (see `NormaliseDate`). `CDate` is only acceptable on unambiguous inputs: numeric date serials, and time-of-day text (which has no day/month ambiguity, e.g. the `TimeValue`/`CDate` fallback in `NormaliseTime`).
- Dates written to review/output sheets use the `CopyDateAsText` pattern: pre-format the destination cell as `@` (text), then write the canonical date string. This prevents US-locale serial-number reinterpretation.

## Add-in architecture rules

- Always reference `ActiveSheet.Parent`, **never** `ThisWorkbook` — the macro runs from the add-in, not the target workbook.
- Sub names must stay stable across point releases. The QAT button binds to a specific sub name; renaming it silently breaks the entry point for every installed user. The current entry point is `BookwiseReconcileDailyReportV3`.
- Re-runs must be idempotent: column-insertion logic reuses/renames legacy columns rather than duplicating; review sheets are cleared and fully rebuilt each run, not appended to.

## v3.1-specific behaviour

- Robust date/time normalisation via dedicated helper routines: `NormaliseDate`, `NormaliseTime`.
- Dates written using the `CopyDateAsText` pattern (see above).
- **Hard stop:** a duplicate Bookwise **Book No.** halts the macro entirely — this is not a flag-for-review condition, it stops the run before any MASTER changes happen.
- **Resolved in v3.1:** the Bookwise double-click time-format corruption no longer produces false "Modified" flags. `NormaliseTime` normalises text, serial, and genuine-date-valued times to `hh:mm` before comparison, so a time whose stored format was altered by a double-click compares equal to its unchanged counterpart.

## Working conventions

- **Plan before building.** Resolve requirements and open questions before writing code. For any non-trivial change, state the plan first.
- **Ask before assuming.** If a request is ambiguous, ask a short clarifying question rather than guessing — especially anything touching hard-stop conditions or MASTER write logic.
- **Inspect real files before proposing fixes.** Use Python/openpyxl on real de-identified sample files to confirm actual storage types, formats, and column positions before proposing a fix.
- **Change incrementally.** Make one logical change per commit/PR, tested and confirmed working before the next. Wait for a precise failure description before diagnosing.

## Version control (GitHub + Claude Code)

- This repo is on GitHub and worked on via Claude Code. Version control is live — this is the workflow, not an interim measure.
- **VBA is version-controlled as the plain-text `.bas` export** (`BookwiseComparisonMacro.bas`) — no native binary `.xlam` committed to git. There are currently no class modules; if any `.cls` modules are ever added, export them as plain text too.
- After every change to the macro, keep the committed `.bas` in sync with the actual `.xlam` used in production.
- Named releases are tracked as Git tags (`v1`, `v3`, `v3.1`, …); the latest version is the top of the default branch. See the Version History table in `README.md`.

## Open items

- Log any noted-but-not-yet-actioned enhancements here as they come up, so they aren't lost between sessions. *(None currently open — the double-click time-format fix was resolved in v3.1; see above.)*
