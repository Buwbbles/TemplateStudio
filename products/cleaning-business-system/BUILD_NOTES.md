# Cleaning Business System v1 — Build Notes (Builder)

Brief: Brief 1, Cleaning Business System, $14.99 (Excel-first pass; Google Sheets version comes after Google is connected).

## STATUS: SOURCE COMPLETE, NOT YET BUILT OR VERIFIED

This run was **unattended with no shell access**, and the shared folder `/home/snow/TemplateStudio/` is **not granted** to this agent's file tools ("an autonomous run cannot bless a new folder"). So:

- `CleaningBusinessSystem_v1.xlsx` does **not exist yet**. The build script below generates it.
- **Nothing in this file has been run.** No LibreOffice recalc, no error scan, no add-a-client test has happened.
- All files currently live in the Builder's private workspace at `products/cleaning-business-system/`.

To finish (watched session, ~2 minutes):

```bash
cd <folder containing these files>
python3 build_cleaning_system.py          # writes CleaningBusinessSystem_v1.xlsx
python3 verify_cleaning_system.py         # recalcs 4 variants in LibreOffice, writes VERIFY_REPORT.txt, exit 0 = pass
mkdir -p /home/snow/TemplateStudio/products/cleaning-business-system
cp build_cleaning_system.py verify_cleaning_system.py BUILD_NOTES.md CleaningBusinessSystem_v1.xlsx VERIFY_REPORT.txt \
   /home/snow/TemplateStudio/products/cleaning-business-system/
```

Both scripts use only the Python standard library (no openpyxl needed). If either throws, that's a real bug — paste the traceback to Builder.

## Files

| File | What it is |
|---|---|
| `build_cleaning_system.py` | Generates the .xlsx directly (zip + SpreadsheetML). Holds all sample data, the formulas, validations, styles. `--empty` builds a cleared copy. |
| `verify_cleaning_system.py` | Lints formulas, scans text, builds 4 variants, recalculates each with headless LibreOffice, checks for error values and compares key numbers to a Python re-computation. Writes `VERIFY_REPORT.txt` and `recalc_check_sample_LibreOffice.xlsx`. |
| `BUILD_NOTES.md` | This file. |
| `CleaningBusinessSystem_v1.xlsx` | *Produced by the build script — not yet produced.* |

## Tab map

| Tab | Who types | What's on it |
|---|---|---|
| Start Here | nobody | Colour key, 5 setup steps, 60-second first-job walkthrough, exact ranges to clear sample data, tips. |
| Settings | buyer (yellow cells) | A5:B13 business info: name B5, owner B6, phone B7, email B8, **tax rate B9**, **currency symbol B10**, **payment terms B11**, **P&L year B12**, **invoice prefix B13**. Services D5:F19 (name, standard rate, basis). Expense categories H5:H19. Payment methods J5:J12. |
| Clients | A–H input, I–K auto | Name (unique, enforced), phone, email, service address, usual service (dropdown), custom rate (optional), frequency (dropdown), notes. Auto: completed jobs, total billed, balance owed. 200 rows (5–204). |
| Jobs | A–H input, I–P auto, Q hidden helper | Date, Client (dropdown from Clients), Service (dropdown, blank = client's usual), Qty/hours (blank = 1), Extras, Rate override, Status (Scheduled/Completed/Cancelled), Notes. Auto: address, service used, rate, subtotal, tax, total, invoice #, payment status. 500 rows (5–504). |
| Invoices | A–J auto (K hidden helper); M–P input, Q auto | Left: invoice register — one row per Completed job, compact, in job order. Right: Payments Log (date, invoice # dropdown, paid with, amount, client auto). KPIs in row 2. 500 invoices / 1000 payments. Tab is named "Invoices" (a "/" isn't allowed in tab names); its title reads "Invoices & Payments". |
| Expenses | A–F input | Date, category (dropdown), description, amount, paid with, notes. Total in B2. 500 rows. |
| Monthly P&L | nobody | Rows 5–16 = 12 months of the Settings year: jobs completed, revenue (before tax), sales tax collected, expenses, net profit, margin, cash collected. Row 17 year totals. Rows 20–36: expenses by category × month, with totals. |

Colour code: teal header = type here; dark-gray header + light-gray cells = automatic.

## Key formulas (row 5 shown; every row is the same pattern)

Lookups all use `MATCH($B5, Clients!$A$5:$A$204, 0)` — written below as `m`.

**Jobs**
- I Address: `IF($B5="","",IFERROR(INDEX(Clients!$D$5:$D$204,m)&"",""))` — `&""` turns an empty address into blank text instead of 0.
- J Service used: `IF($B5="","",IF($C5<>"",$C5,IFERROR(INDEX(Clients!$E$5:$E$204,m)&"","")))`
- K Rate — priority: Rate override → client's custom rate (only when the job is the client's usual service, i.e. Service blank or equal to it) → Settings standard rate for the service used; 0 if nothing matches:
  `IF($B5="","",IF($F5<>"",$F5,IFERROR(IF(AND(INDEX(Clients!F,m)&""<>"",OR($C5="",$C5=INDEX(Clients!E,m))),INDEX(Clients!F,m),INDEX(Settings!$E$5:$E$19,MATCH($J5,Settings!$D$5:$D$19,0))),0)))`
- L Subtotal: `IF($B5="","",$K5*IF($D5="",1,$D5)+$E5)`
- M Tax: `IF($B5="","",ROUND($L5*Settings!$B$9,2))` (rounded per job so invoice totals are exact cents)
- N Total: `IF($B5="","",$L5+$M5)`
- Q helper (hidden) — running count of invoiceable jobs: `IF(AND($A5<>"",$B5<>"",$G5="Completed"),COUNTIFS($A$5:$A5,"<>",$B$5:$B5,"<>",$G$5:$G5,"Completed"),"")`
- O Invoice #: `IF($Q5="","",Settings!$B$13&TEXT(ROWS($A$5:$A5),"0000"))` — number = the job's line (INV-0007 = 7th job row). Stable: changing another job's status never renumbers this one. Gaps appear for Scheduled/Cancelled lines.
- P Payment status: `INDEX(Invoices!$I$5:$I$504, MATCH($O5, Invoices!$A$5:$A$504, 0))` guarded by IF/IFERROR.

**Invoices (register)**
- K helper (hidden): `IFERROR(MATCH(ROWS($K$5:$K5),Jobs!$Q$5:$Q$504,0),"")` — "which job row is invoice #k". This is how a compact list is built without FILTER.
- A/B/C/D/E: `IF($K5="","",INDEX(Jobs!<col>$5:<col>$504,$K5))` for invoice #, date, client, service used, total.
- F Paid: `IF($A5="","",SUMIFS($P$5:$P$1004,$N$5:$N$1004,$A5))` — sums every payment logged against the invoice (part payments OK).
- G Balance: `E-F`. H Due date: `$B5+Settings!$B$11` (real date + whole days — no text dates anywhere).
- I Status: `IF($G5<=0,"Paid",IF(TODAY()>$H5,"Overdue",IF($F5>0,"Partial","Unpaid")))`
- J Days overdue: `IF(AND($G5>0,TODAY()>$H5),TODAY()-$H5,"")` — blank when not overdue.
- Q Client on payment row: lookup of the picked invoice # in A; shows "Invoice not found" if the number isn't in the register.

**Monthly P&L** (A5 = `DATE(Settings!$B$12,1,1)` … A16 = month 12)
- Window used in every column: `range,">="&$A5, range,"<="&EOMONTH($A5,0)`
- B Jobs: `COUNTIFS(Jobs!$G:$G-range,"Completed", window on Jobs!A)`
- C Revenue: `SUMIFS(Jobs!$L$5:$L$504, status "Completed", window)` (before tax — tax is not income)
- D Tax: same on Jobs!M. E Expenses: `SUMIFS(Expenses!$D$5:$D$504, window on Expenses!A)`
- F Net: `C-E`. G Margin: `IF($C5=0,"",$F5/$C5)` (blank, never #DIV/0!). H Cash collected: SUMIFS on payment amount by payment date.
- Category grid: `IF($A21="","",SUMIFS(Expenses!D, Expenses!B,$A21, window built from DATE(year,k,1)))`.

**Clients** I/J/K: COUNTIFS of Completed jobs; SUMIFS of Invoices!E and Invoices!G by client.

**Functions used (complete list):** IF, IFERROR, AND, OR, INDEX, MATCH, SUMIFS, COUNTIFS, COUNTIF, SUM, ROUND, ROWS, TEXT, DATE, EOMONTH, TODAY. All exist in Excel 2007+ and Google Sheets. No XLOOKUP, LET, LAMBDA, FILTER/UNIQUE/SORT/SEQUENCE, IFS, MAXIFS, TEXTJOIN, QUERY, ARRAYFORMULA, macros, named ranges or tables. The verify script enforces this whitelist and rejects any `_xlfn.` prefix.

## Enter-once behaviour (fixes for the competitor complaints)

- "Info over and over again": client details typed once on Clients; Jobs pulls address/service/price; Invoices pulls from Jobs; P&L and client balances roll up automatically. Payment entry is a 4-cell row picked from a dropdown.
- "Couldn't work out how to edit": teal = type, gray = automatic; every input column has a hover tip; Start Here has a 60-second walkthrough.
- "Can't get anything in Excel": built as a native .xlsx with only classic functions; no Sheets-only features.
- Date complaints: date columns only accept real dates (data validation 1/1/2000–12/31/2099), date maths is plain serial arithmetic and EOMONTH windows.

Data validation: client dropdown, service dropdowns, status/frequency lists, category & payment-method lists from Settings, invoice-number dropdown on payments, date and ≥0 number checks, unique client names (`COUNTIF(...)=1`), tax 0–100%, terms 0–365, year 2000–2099. Conditional formats: Overdue red, Paid green, Partial amber (Invoices I and Jobs P). Header rows frozen. No sheet protection (protection makes "select and Delete" fail in Excel, which would break the clear-sample instructions).

## Sample data

8 fictional clients (example.com emails, 555 numbers), 30 jobs Jul–Sep 2026 (27 Completed, 1 Cancelled, 2 Scheduled), 10 expenses, 22 payments (5 invoices left unpaid incl. one old overdue one, 1 part-paid). Exercises: custom client rate, hourly service with qty, extras, explicit different service, rate override, cancellation.

Hand-computed expected results (the verify script checks these against LibreOffice; not yet confirmed):

| Month 2026 | Jobs | Revenue | Expenses | Net |
|---|---|---|---|---|
| Jul | 10 | 1,435.00 | 483.14 | 951.86 |
| Aug | 9 | 1,325.00 | 569.20 | 755.80 |
| Sep | 8 | 1,182.50 | 246.35 | 936.15 |

Clearing sample data (as printed on Start Here): Clients A5:H204, Jobs A5:H504, Expenses A5:F504, Invoices M5:P1004 — Delete key only. Settings keep their sample values; the buyer overwrites them.

## What verify_cleaning_system.py checks

1. Formula whitelist + no `_xlfn`; all 7 tabs present in order; no AI wording in any cell, formula or document property.
2. For each variant — **sample**, **empty** (sample cleared exactly as Start Here says), **empty_plus1** (cleared + new client "Lakeside Bakery" + one Completed job 9/25/2026), **sample_plus1** — recalculated with
   `soffice -env:UserInstallation=file://<tmp profile> --headless --norestore --nologo --convert-to "xlsx:Calc MS Excel 2007 XML" --outdir <out> <file>`
   (temp profile sets OOXMLRecalcMode=0 = always recalc on load):
   - >10,000 formula cells came back with values; zero #REF!/#VALUE!/#NAME?/#DIV/0!/#N/A/#NUM!/Err:xxx;
   - every job's rate/total and real-date check; unused job/invoice/client rows are blank (not 0, not errors);
   - every invoice's number, client, amount, paid, balance, status and days overdue vs Python (using today's date);
   - P&L jobs/revenue/expenses/cash per month; margin blank when no revenue; category grid total = expense total;
   - client billed/balance roll-ups.
3. New-client checks: job pulls the address; INV-0001 appears for Lakeside Bakery on a cleared file; Sep revenue = 120.00; on the sample file Lakeside is the last invoice and Sep revenue rises by exactly 120.00.

Limitation of the add-a-client test: it regenerates the file with the extra rows rather than typing into an open file — same cells and formulas, but it doesn't exercise dropdown clicks.

## Verified / assumed

- **verified:** nothing executed. I only wrote the source and re-read it (fixed two Python issues by inspection: backslashes inside f-string expressions that break on Python < 3.12, and a `.replace()` that bound to a tuple).
- **assumed (not checked):** that both scripts run without errors; that LibreOffice recalculates with no error values; that the expected numbers above match; that Excel 2016+/365 opens the file with no repair prompt and recalculates on load (`fullCalcOnLoad="1"`; the file ships **without cached values**, so previews such as Protected View or a file-browser thumbnail may show blank gray columns until Excel calculates); Google Sheets import behaviour (dropdowns from other tabs, conditional formats, frozen rows); Mac Excel; column widths and row heights look right on screen.

## Known tradeoffs / open items

- Invoice numbers follow job lines, so buyers shouldn't sort the Jobs tab or delete job rows (Start Here says so; Cancelled is the alternative). Sorting would renumber invoices and detach logged payments.
- Revenue is accrual (by job date); cash basis is shown alongside as Cash collected.
- Client names with `*` or `?` act as wildcards in SUMIFS/COUNTIFS (Start Here warns).
- Changing a service name on Settings doesn't rename it on existing clients/jobs.
- Optional next step after verification: ship the LibreOffice-recalculated copy (has cached values) instead of the raw build, **only** after checking it kept validations, conditional formats, hidden helper columns and frozen panes in real Excel.
- Checker needs real Excel and Google Sheets once connected; neither is available yet.
