# Cleaning Business System — Etsy Preview Fixes (T-004)

Source of truth: Google Sheet `1ALTVh9eINKnQER5FB-tC9HeqsW_xAhQVnlJeFZ4xV5w` ("Cleaning Business System v1 - Sheets test build"), read live via the Sheets connector. No spreadsheet cells were changed — read-only.

## MUST FIX

**1. Preview 1 — removed the "+ Excel" badge.**
We only sell the Google Sheets version right now. Deleted the second `.badge` div; kept a single "Google Sheets" badge.

**2. Preview 2 — Priya Raman rate/total row.**
Checked against `Jobs!A8:Q8` (7/8/2026, Priya Raman, Deep Clean): Rate `$220`, Extras `$25`, Subtotal `$245`, Tax `$14.70`, Total `$259.70`, Invoice `INV-0004`. **The original numbers were already an exact match to the sheet** — the row only *looks* inconsistent with the "rate × 1.06" rows because of the $25 extra add-on (fridge clean), which the mockup didn't explain. Rather than invent different numbers, I added a footnote: "This visit added a $25 extra (inside-fridge clean), which is why the total is more than rate × the usual 6% tax — still an exact match to the Jobs tab," marked the row with a `*`, and added one more real row (`Jobs!A22`, Priya Raman, Standard Clean, $120/$127.20, INV-0018 — a second Priya Raman job at her *normal* rate with no extras) so a reader can compare the two Priya rows directly and see the $25 difference is the only variable.

**3. Preview 3 — "Payments Log on the right" footnote.**
The image doesn't show a Payments Log panel, so the old line ("Log each payment on the right (Payments Log)") pointed at something not visible. Reworded to: "Log each payment in the Payments Log columns on the same tab, and the balance, status and days overdue update on their own" — accurate to the real sheet (Invoices columns M–Q: Payment date / Invoice # / Paid with / Amount / Client), with no claim about on-screen position.

**4. Every number checked against the live sheet.** Full verification table below — all were already correct except item 2's framing; no invented numbers were introduced anywhere.

## SHOULD FIX

**5. Filled empty space with real data (no decoration-only content):**
- P1: added a 3-row "features" strip (Start Here walkthrough, pre-loaded sample data, color-coded cells) in the gap between the benefit banner and the tab-pill row.
- P2: added 2 more real client rows (Danielle Brooks, Riverside Yoga Studio — `Clients!A11`, `Clients!A12`) and 1 more real job row (`Jobs!A22`, 8/18/2026 Priya Raman Standard Clean, INV-0018), enlarged both mini-tables, bumped font sizes, and resized panels to match the new row counts (no empty bottom strip).
- P3: added 2 more real invoice rows (`Invoices!A24` INV-0021 Marcus & Elena Ortiz, `Invoices!A27` INV-0024 Riverside Yoga Studio) and grew the panel to fit them.
- P4: added a third stat card, "Profit margin (YTD) — 67.1%" from `Monthly P&L!G17` (Year total row), filling the empty right side next to the two existing stat cards.
- P5: the 7th tile (Monthly P&L) now spans the full grid width (`grid-column:1/-1`) with a dark accent style instead of sitting alone in a half-empty row.

**6. P1 footer + page tag.**
Footer now lists all 7 real tabs by name (Start Here, Settings, Clients, Jobs, Invoices, Expenses, Monthly P&L) — Settings was missing before. Added a "1 of 5" tag in the top bar, matching the page-counter style already used on previews 2–5.

**7. Spelling + wrapped word.**
- P5 card 7: "totalled" → "totaled" (US spelling).
- P1 mini-table header: "Balance owed ($)" → "Owed ($)" plus `white-space:nowrap` on table headers and `table-layout:fixed` with explicit column widths, so "owed" can no longer wrap onto its own line in the 660px mock table.

## Verification table (every number shown vs. the sheet)

| Slide | Value shown | Sheet cell |
|---|---|---|
| P1/P2 | Hannah Whitfield: 5 jobs / $583.00 billed / $116.60 owed | `Clients!A5,I5,J5,K5` |
| P1/P2 | Marcus & Elena Ortiz: 5 jobs / $477.00 / $0.00 | `Clients!A6,I6,J6,K6` |
| P1 | Brightpath Dental Office: 4 jobs / $0.00 owed | `Clients!A7,I7,K7` |
| P1/P2 | Priya Raman: 3 jobs / $620.10 / $233.20 | `Clients!A8,I8,J8,K8` |
| P2 | Greenleaf Property Mgmt: 3 jobs / $996.40 / $478.40 | `Clients!A10,I10,J10,K10` |
| P2 | Danielle Brooks: 2 jobs / $360.40 / $0.00 | `Clients!A11,I11,J11,K11` |
| P2 | Riverside Yoga Studio: 3 jobs / $318.00 / $0.00 | `Clients!A12,I12,J12,K12` |
| P2 | Jobs rows (rate/total/invoice): INV-0001 110/116.60, INV-0002 45/143.10, INV-0004 220/259.70, INV-0008 300/318.00, INV-0018 120/127.20, INV-0027 95/100.70 | `Jobs!A5` (INV-0001), `A6` (INV-0002), `A8` (INV-0004), `A12` (INV-0008), `A22` (INV-0018), `A31` (INV-0027) — Rate/Total/Invoice# columns |
| P3 | Total billed $4,179.05 / Collected $3,122.95 / Outstanding $1,056.10 / Overdue 5 | `Invoices!B2,D2,F2,H2` |
| P3 | Invoice rows INV-0009, 0017, 0025, 0026, 0027, 0010, 0020, 0021, 0024 (amount/balance/due/status/days) | `Invoices!A13` (INV-0009), `A21` (INV-0017), `A28` (INV-0025), `A29` (INV-0026), `A30` (INV-0027), `A14` (INV-0010), `A23` (INV-0020), `A24` (INV-0021), `A27` (INV-0024) |
| P4 | Jul 2026: 10 jobs / $1,435.00 / $86.10 tax / $483.14 exp / $951.86 net | `Monthly P&L!A11:F11` |
| P4 | Aug 2026: 9 / $1,325.00 / $79.50 / $569.20 / $755.80 | `Monthly P&L!A12:F12` |
| P4 | Sep 2026: 8 / $1,182.50 / $70.95 / $246.35 / $936.15 | `Monthly P&L!A13:F13` |
| P4 | Year total: 27 / $3,942.50 / $236.55 / $1,298.69 / $2,643.81 / 67.1% margin | `Monthly P&L!A17:G17` |

## Not verified
I have no shell access in this run, so I could not render the HTML to PNG and eyeball the final layout pixel-by-pixel — the CSS positioning/sizing above is computed by hand from the existing stylesheet's own box model, not confirmed by a screenshot. LUCILLE renders the PNGs next; flag back to me if any panel overflows or text clips at 2000×1600 and I'll adjust.

Budget used: 1 Sheets metadata call + 6 range reads + 5 file rewrites + this note. Well under the $0.60 cap.
