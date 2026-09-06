# Workbook generator

`build_workbook.py` builds the "Banff Pride Partnerships" workbook (.xlsx, Google Sheets compatible) from a local `seed.py` and `july.json`.
Those two inputs hold partner contact details and live commitments, so they are **not** committed. Re-running the generator is only for schema changes; day-to-day edits happen in the live Google Sheet.

Formulas use only functions shared by Excel 365 and Google Sheets (FILTER, SORT, INDEX, COUNTIFS, SUMIFS, IFERROR). Views read record tabs; nothing is typed into a view.
