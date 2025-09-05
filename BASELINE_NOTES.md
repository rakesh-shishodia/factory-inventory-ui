# Baseline Notes — Inventory Management System

This document records the state of the project after reverting to commit **f2d8639** and cleaning up the Google Sheets configuration.

## Code Baseline
- Repo hard-reset to commit `f2d8639`: *Implement min-rule sync with allow_raise flag for Ecwid stock updates*.
- No additional code changes since that commit.

## Google Sheets Configuration
- **Products** sheet:
  - Data validation applied to enforce numeric values (e.g. current_stock ≥ 0).
  - Headers protected and first row frozen.
- **SyncQueue** sheet:
  - Headers: `ts | sku | target_stock | allow_raise | status | last_error`.
  - Removed validation from `allow_raise` column (to avoid auto-filled FALSE values).
  - Manually deleted ghost rows that caused appendRow() to write at the bottom.
  - Headers protected and first row frozen.
- **Employees** sheet:
  - Contains a single column `name` listing employee names.
  - This list is used to populate the Employee Name dropdown in the mobile UI.
  - To add/remove employees, simply edit this sheet (delete row to remove, add row to add).
  - Changes are cached for 24h in the UI but will auto-refresh after that period.

## Status
System is confirmed working:
- Transactions add rows to **Transactions** sheet.
- Stock updates correctly in **Products**.
- Queue rows are appended correctly to **SyncQueue** with `status=queued`.

## Next Planned Features
- Add archiving/rotation for SyncQueue (move done rows to a log sheet monthly).
- Improve logging and reporting.
- UI polish and mobile optimizations.