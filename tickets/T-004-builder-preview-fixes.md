# T-004 — BUILDER: fix Cleaning Business System previews

Owner: BUILDER (engineer) · Reviewer: CHECKER (Gatekeeper) · Budget cap: $0.60 · Model: Sonnet
Inputs: your previews at products/cleaning-business-system/previews/preview-1..5.html;
rendered PNGs + contact sheet at /home/snow/TemplateStudio/products/cleaning-business-system/listing-images/;
the live Google Sheet 1ALTVh9eINKnQER5FB-tC9HeqsW_xAhQVnlJeFZ4xV5w (source of truth for every number).

## Must fix (blocks listing)
1. Preview 1: remove the "+ Excel" badge. We sell Google Sheets only until a real Excel test passes.
2. Preview 2: Priya Raman row shows rate $220 and total $259.70; other rows are rate x 1.06. Copy the
   exact values for every row from the live sheet. No invented numbers anywhere.
3. Preview 3: footnote mentions a "Payments Log on the right" that is not shown. Show it or delete the line.
4. Every number on every slide must match the sheet's sample data exactly.

## Should fix
5. Fill the empty space: P1 gap between CTA and tab pills; P2/P3 card bottoms half empty; P4 right half empty;
   P5 orphaned 7th tile. Bigger, readable mini-tables (no wrapped cells).
6. P1 footer tab list omits Settings; P1 lacks the "1 of 5" tag.
7. US spelling: "Totalled" -> "Totaled". Fix the lone wrapped word "Owed.".

## Output
Updated preview-1..5.html (2000x1600) in the same folder, plus a CHANGES.md listing each fix.
Stop condition: all 7 items done, or budget hit. Do not change the spreadsheet.
