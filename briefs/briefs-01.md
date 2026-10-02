# Validator Briefs 01
Source scan: `strategy/niche-scan-01.md` · Raw evidence: `strategy/raw/scan-notes-01.md` · Captured around early Oct 2026 from live Etsy pages.

**Validator rule:** a brief passes only with (a) 3+ competitors showing visible sales, (b) their prices set against our $12 floor, and (c) a 1–3★ complaint we can fix, all linked.
**Result:** 3 PASS · 2 HOLD.
"Shop sales" means the shop's lifetime "N Sales" figure on its `/reviews` page. Etsy doesn't show per-listing sales on the pages I read.

---

## Brief 1: Cleaning Business System (Excel + Google Sheets): **PASS**

**Buyer:** owner-operator or small-team residential cleaner who tracks clients, jobs, payments and expenses.

**Competitors with visible sales**

| Shop | Shop sales | Their price | vs $12 | Evidence |
|---|---|---|---|---|
| SavvyandThriving | 47,310 | $8.76 (was $17.52) | below | [listing 4311063860](https://www.etsy.com/listing/4311063860): 4.6★ from 22 item reviews |
| TheProductivePlans | 57,522 | $6.04 (was $12.09) | below | [shop reviews](https://www.etsy.com/shop/TheProductivePlans/reviews) |
| Evelibrary | 20,472 | $12.35 | at floor | [shop reviews](https://www.etsy.com/shop/Evelibrary/reviews) |
| DocuSimplified | 1,000 | $11.03 | below | [shop reviews](https://www.etsy.com/shop/DocuSimplified/reviews) |
| Dechampionz | 414 (5.0★, 24) | $14.49 | above | [listing 4524482994](https://www.etsy.com/listing/4524482994) |

Others priced above the floor: ZMRBusinessSolutions $17.43, SimpleBizOps $14.99, ARQELYN $15.89, Solendrix $26.42. 9 of 20 first-page cards are at or above $12 ([search](https://www.etsy.com/search?q=cleaning+business+spreadsheet)).

**Complaints we can fix** (all on the SavvyandThriving cleaning bundle, [listing 4311063860](https://www.etsy.com/listing/4311063860), review panel filtered by star rating):
- 1★, Laly, Aug 13 2025: "I can't get anything in Excel and it won't let me do anything."
- 1★, Jessica, Jun 25 2025: "Couldn't work out how to edit even after I downloaded different apps to use. Gave up and haven't used."
- 3★, Becky, Mar 26 2026: "I thought it was going to communicate between better. You have to enter in information over and over again no carry over"

**Our fix (build spec)**
1. Two native files: a real `.xlsx` and a Google Sheets copy link. No Google login needed for the Excel buyer.
2. Enter data once. The Clients tab feeds Jobs through dropdowns, Jobs feeds Invoices/Payments, and everything rolls up to a Monthly P&L. A client's name, address and rate are never retyped.
3. A "Start Here" tab with 5 steps and a 60-second first-job walkthrough, plus sample data that can be cleared in one step.

**Price:** $14.99. That's above the floor and in line with Dechampionz $14.49 and SimpleBizOps $14.99.

**Checker acceptance:** every formula recalculates in both Excel and Sheets; a new client entered once shows up in Jobs, Invoice and P&L; the file opens in Excel with no Google account; the preview images are screenshots of the shipped file.

**Why first:** this is the strongest evidence set, and the same engine gives us lawn care and pressure washing variants.

---

## Brief 2: Home Bakery Costing + Orders (Excel + Google Sheets): **PASS**

**Buyer:** home or cottage baker pricing cakes, cookies and bread for orders and markets.

**Competitors with visible sales**

| Shop | Shop sales | Their price | vs $12 | Evidence |
|---|---|---|---|---|
| HeyMorning | 61,998 | $6.99 (was $13.98), Bestseller | below | [listing 4351747898](https://www.etsy.com/listing/4351747898): 4.8★ from 44 reviews |
| SavvyandThriving | 47,310 | $3.29 | below | [search](https://www.etsy.com/search?q=recipe+cost+calculator+spreadsheet) |
| ProductivePenguinCo | 38,382 | $6.52 (was $13.05) | below | [shop reviews](https://www.etsy.com/shop/ProductivePenguinCo/reviews) |
| BocaBocaTemplates | 13,685 | $8.24 (was $16.47) | below | [listing 1833211437](https://www.etsy.com/listing/1833211437): 3.6★ from 9, 67% recommend |
| SoloFinanceTools | 322 | $17.27, Bestseller (home bakery bundle) | above | [shop reviews](https://www.etsy.com/shop/SoloFinanceTools/reviews) |
| TheModernTemplateSt | 521 | $12.43, "In 6 carts", Sheets only | at floor | [listing 4313137995](https://www.etsy.com/listing/4313137995): 3.9★ from 7 |
| AtlasWander | 2.7k | $17.65 | above | [listing 1811920524](https://www.etsy.com/listing/1811920524): 3.8★ from 4 |

6 of 19 first-page cards are at or above $12, including ClearNestDigitalCo $16.99, OpsEdgeHospitality $17.00 and BusinessInTemplates $20.50.

**Complaints we can fix**
- 1★, Cara, Jul 27 2025 (TheModernTemplateSt, [4313137995](https://www.etsy.com/listing/4313137995)): "It does not allow you to change the amount in your recipe vs whats on the ingredient master sheet."
- 2★, Feb 21 2026 ([BuildBooksTemplates reviews](https://www.etsy.com/shop/BuildBooksTemplates/reviews), page 4; the item title was cut off in my capture, but the review is about recipe costing): "What would be more helpful is a sheet where we can write down a product such as flour, what a 5lb bag cost and then let it calculate what a cup costs. I cant do that with this spread sheet."
- 1★, Aug 22 2026 (BuildBooksTemplates, "Recipe Cost Template Excel & Google Sheets"): "Not even close to what was shown in the post. No color, not graphs.....just three or four very basic pages."
- 1★, Jennifer, Oct 16 2025 (AtlasWander, [1811920524](https://www.etsy.com/listing/1811920524)): "I have been unable to download with ZERO help from seller"

**Our fix (build spec)**
1. Unit conversion at the core. Enter "flour, 5 lb bag, $4.50" once, and recipes can use cups, grams, oz or tbsp, with the cost computed per unit used. This needs a density table for common baking ingredients (flour, sugar, butter, etc.) that the buyer can edit.
2. Recipe quantities are free to change per recipe, separate from the purchase size on the Ingredients sheet. This answers Cara's complaint directly.
3. Price builder: ingredients + packaging + labor/hr + overhead % → suggested price at a target margin. An Orders tab feeds a monthly profit summary.
4. Native Excel and Google Sheets. Preview images are real screenshots of the shipped file.

**Price:** $14.99. Justified by depth (unit conversion plus orders plus P&L) against $3–9 single-sheet calculators. SoloFinanceTools sells a bakery bundle at $17.27 as a Bestseller.

**Risk:** the cheapest competition of the three PASS briefs. This product sells on visible depth in the mockups, not on price.

**Checker acceptance:** "5 lb flour at $4.50" gives the correct cost per cup (shown in the instructions tab with the math); changing a recipe amount never edits the Ingredients sheet; Excel and Sheets results match to the cent.

---

## Brief 3: Hair Stylist / Booth Renter Books (Excel + Google Sheets): **PASS on evidence, PRICE RISK**

**Buyer:** booth-renting or commission stylist tracking services, tips, product sales, booth rent and tax set-aside.

**Competitors with visible sales**

| Shop | Shop sales | Their price | vs $12 | Evidence |
|---|---|---|---|---|
| BlushDesignTemplates | 25,806 | $1.71, "In 20+ carts" | below | [listing 4486490946](https://www.etsy.com/listing/4486490946): 4.5★ from 9 |
| SavvyandThriving | 47,310 | $4.45 | below | [listing 1433048986](https://www.etsy.com/listing/1433048986): 4.8★ from 22 |
| GraceDigitalsph | 899 | $5.43 | below | [listing 1794160990](https://www.etsy.com/listing/1794160990) |
| BoldlyCreatedGifts | 431 | $7.21 | below | [shop reviews](https://www.etsy.com/shop/BoldlyCreatedGifts/reviews) |
| AnarosaStudioCo | 189 (4.6★, 15) | $7.19 | below | [listing 4513374790](https://www.etsy.com/listing/4513374790) |

Only 4 of 16 first-page cards are at or above $12 (ValueHealer $14.99, LedgerForgeCo $14.99, TheMineLab $14.58, RunTheCalc $19.99), and none of those shops' sales were verified.

**Complaints we can fix**
- 1★, Tina, Jul 16 2026 (BlushDesignTemplates, [4486490946](https://www.etsy.com/listing/4486490946)): "I am did not receive the correct downloads for the spreadsheets I ordered."
- 1★, Rebecca, May 12 2025 (SavvyandThriving, [1433048986](https://www.etsy.com/listing/1433048986)): "Can't get what I paid for"

**Our fix:** correct, tested delivery (the Checker confirms every link and file), plus depth the $2–7 sheets lack: a booth rent vs commission comparison, tips logged separately, product-sales margin, and a quarterly estimated-tax set-aside.

**Price:** $12.99 at most. The market is anchored at $2–8.

**Validator note:** this passes the rule, but the complaints are about delivery, not the tool, and the price gap is the widest of the three. It ranks third, after Briefs 1 and 2.

---

## Brief 4: Lawn Care / Landscaping System: **HOLD**

**Competitors with visible sales:** WaveflowTemplates 12,224 sales at $9.60 · BuildBooksTemplates 4,091 at $3.99 ([listing 4349166175](https://www.etsy.com/listing/4349166175)) · Dechampionz 414 at $14.49 · TheSheetBook 164 at $7.85 · KSMadeDigital 142 at $14.25 · MuseAndMakeCo 115 at $14.30.

**Prices:** the best fit of any niche, with **14 of 21** first-page cards at or above $12 (FlightStreetSystems $19.95, LawnCareSystems $18.99, QuickQuoteSheets $19.00, TheReadyMadeStudio $24.99) ([search](https://www.etsy.com/search?q=lawn+care+business+spreadsheet)).

**Missing:** no lawn-specific 1–3★ review yet. The closest is a **4★** at Dechampionz (Sep 13 2026; the item wasn't shown in my capture): "the ease of use isn't so good for me … For the older generation, not so much." That doesn't meet the 1–3★ rule.

**To pass:** one lawn or landscaping 1–3★ review with a link. Planned build: the Brief 1 engine plus route/day scheduling, a recurring mow schedule and seasonal services.

---

## Brief 5: Candle Making Cost Calculator: **HOLD**

**Competitors with visible sales:** SerenatasJourney 38,648 sales at $15.24 · SavvyandThriving 47,310 at $3.67 ([listing 1880416837](https://www.etsy.com/listing/1880416837), 4.9★ from 9) · CraftCalculators 1,988 at $4.55 · WickSetGo 632 at $3.49 · AnnesArtDesignSpace 153 at $8.70 ([listing 1401885213](https://www.etsy.com/listing/1401885213), 4.9★ from 8).

**Prices:** 5 of 19 at or above $12 ([search](https://www.etsy.com/search?q=candle+making+cost+calculator)).

**Missing:** no candle-specific 1–3★ review. The adjacent one is 1★, Jan 7 2026, on HandySheetTemplates' Etsy pricing calculator (shop: 6,306 sales): "I don't appreciate spreadsheets that I have to go and fix the formulas on." It's the same calculator category but a different niche, so it doesn't count.

**To pass:** one candle-specific 1–3★ review with a link. Planned build: the Brief 2 costing engine adapted to wax weight, fragrance load %, wick/vessel cost and batch yield.

---

## Validator summary for the Commander

| Brief | Status | Suggested price | Build order |
|---|---|---|---|
| 1. Cleaning Business System | PASS | $14.99 | Build #1 |
| 2. Home Bakery Costing + Orders | PASS | $14.99 | Build #2 |
| 3. Stylist / Booth Renter Books | PASS, price risk | $12.99 | Build #3 or swap for the lawn variant once it passes |
| 4. Lawn Care System | HOLD | ~$14.99 | Variant of #1 |
| 5. Candle Cost Calculator | HOLD | ~$12.99 | Variant of #2 |

**Verification limits:** shop sales are shop-wide, not per listing. Reviews were read on live Etsy pages. Any line marked "title cut off" or "item not shown" could not be tied to an exact item.
