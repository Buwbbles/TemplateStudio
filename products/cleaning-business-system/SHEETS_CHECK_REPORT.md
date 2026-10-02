# Cleaning Business System v1 - Google Sheets check

Date: 2026-10-02
Live file: https://docs.google.com/spreadsheets/d/1ALTVh9eINKnQER5FB-tC9HeqsW_xAhQVnlJeFZ4xV5w/edit
Built from: build_cleaning_system.py -> gsheets_plan_sample.json (same source as CleaningBusinessSystem_v1.xlsx)
Reference: ref_libreoffice_sample.json (LibreOffice recalculation, 53/53 checks passed)

## Result: PASS on formulas, NOT yet shippable (no styling/dropdowns)

All formulas loaded and calculated in real Google Sheets with zero error cells in the ranges read.
Every total checked matches the LibreOffice numbers exactly:

| Check | LibreOffice | Google Sheets |
|---|---|---|
| Total billed / Collected / Outstanding / Overdue count | 4179.05 / 3122.95 / 1056.10 / 5 | 4179.05 / 3122.95 / 1056.1 / 5 |
| P&L year: jobs / revenue / tax / expenses / profit / cash | 27 / 3942.50 / 236.55 / 1298.69 / 2643.81 / 3122.95 | identical |
| P&L September row | 8 / 1182.50 / 70.95 / 246.35 / 936.15 / 79.2% / 484.95 | identical |
| Expenses by category total | 1298.69 | 1298.69 |
| Clients completed / billed / balance (8 clients) | all 8 rows | identical |
| Invoice list INV-0001..INV-0028 (cancelled job skipped) | yes | yes |
| Last job row -> INV-0028, Unpaid, line 27 | yes | yes |

## Bug found and fixed in the Sheets copy
- Invoices "Days overdue" displayed as a date (e.g. 2/24/1900 instead of 56) because Sheets inherits a date format from TODAY()-date. Fixed by setting J5:J504 to number format "0". The Excel build script must set this format explicitly too (it is exactly the "date formulas didn't work" complaint buyers leave).

## Done after the check (same day)
- Styling: teal input headers, gray automatic headers, light gray automatic cells, yellow Settings cells, column widths, frozen header rows.
- Dropdowns: Jobs client/service/status, Clients usual service, Expenses category/paid-with, Payments invoice #/paid-with. Real-date checks on Jobs, Expenses and Payment dates.
- Warning-only protection on all formula columns and the P&L tab (4 protected ranges).
- Totals re-read after formatting: unchanged.

## Not done yet (needed before listing)
4. Buyer "make a copy" link from Drive.
5. Round trip: download this Sheet as .xlsx and recalc in LibreOffice.
6. Preview images that match the file.
Excel desktop test skipped for now by Commander's call.
