# Changelog

All notable changes to this project are documented here.

## [1.0.0] - 2026-11-02

### Added
- Sync of Basecamp to-dos assigned to you (`/my/assignments.json` + completed) into a Google Sheet.
- Account-wide @mention detection (comments and descriptions) using the cross-project
  recordings feed and your `attachable_sgid`, scanned incrementally with a checkpoint.
- Re-mention behaviour: a fresh mention on a task already in the sheet moves it to the bottom,
  keeps a manual "Working"/"Pending", and resets a manual "Done".
- Comment-anchored links (`...#__recording_<id>`) so one click lands on the mention.
- Due-date colour ladder, manual **My Status** dropdown, and "DONE By Assigner" (forest green)
  when Basecamp itself marks the to-do complete.
- **Activity Time** column (date + time of assignment or latest mention).
- Reconciliation sweep that re-checks rows no new comment would otherwise refresh.
- Auto-jump to the row with the latest Activity Time on open/refresh.
- 5-minute and hourly time-driven triggers, script lock against overlapping runs, **Sync Log** sheet.
- Custom menu: Sync Now, Refresh Colours, Jump to Latest Activity, setup/diagnostic helpers.

### Fixed
- Jump-to-latest did nothing because Sheets shows Activity Time without AM/PM; the parser now accepts both.
- "Jump to Latest Activity" menu item pointed at a private (`_`) function and could not run; added a public wrapper.
- Anchored comment links were overwritten by plain task URLs on the next sync/reconcile; they are now preserved.
- New rows were always created light blue regardless of due date.
- Rows that entered via a comment mention never noticed later completion in Basecamp.
- Syncs failed when run from a time-driven trigger because `SpreadsheetApp.getUi()` throws there.
