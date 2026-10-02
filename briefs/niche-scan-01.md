# Niche Scan 01: 10 small-business spreadsheet niches on Etsy
Run by: LUCILLE (Overseer), doing Strategy's first tasking directly
Data: live Etsy pages read through the station browser, captured around early Oct 2026. Raw capture is in `strategy/raw/scan-notes-01.md`.

## How to read this
- **Price**: the price shown on the first page of Etsy search for the query (organic and ad cards). "≥$12 share" is the share of those cards priced at or above our $12 floor. Many sellers show a "sale" price, and I counted the price a buyer actually pays.
- **Sales**: each shop's lifetime **"N Sales"** figure from `etsy.com/shop/<shop>/reviews`. This counts the whole shop, not one listing. Etsy doesn't show per-listing sales on these pages, so where I have per-listing evidence I give the listing's review count instead.
- **Complaint**: a 1–3★ review I read myself on Etsy, either in a listing's review panel (filtered by star rating) or on a shop's review pages. A review from a different niche is labeled "adjacent".
- **Score** (1–5 each, 20 max): Demand (proven sellers) · Price fit vs $12 · Fixable complaint found · Build fit (can we build and test it well in Sheets and Excel).

## Scoreboard

| # | Niche | Search URL | Proven sellers (shop sales) | ≥$12 share | Fixable 1–3★ found? | D | P | C | B | Total | Call |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 1 | Cleaning business system | [search](https://www.etsy.com/search?q=cleaning+business+spreadsheet) | TheProductivePlans 57,522 · SavvyandThriving 47,310 · Evelibrary 20,472 · TrackerPro 19,595 · DocuSimplified 1,000 · Dechampionz 414 | 9/20 | Yes, 3 on one listing | 5 | 3 | 5 | 5 | **18** | **BRIEF: PASS** |
| 2 | Home bakery / recipe costing | [search](https://www.etsy.com/search?q=recipe+cost+calculator+spreadsheet) | HeyMorning 61,998 · SavvyandThriving 47,310 · ProductivePenguinCo 38,382 · PalmAndGrace 14,551 · BocaBoca 13,685 · AtlasWander 2.7k | 6/19 | Yes, 4 across 4 listings | 5 | 2 | 5 | 5 | **17** | **BRIEF: PASS** |
| 3 | Lawn care / landscaping | [search](https://www.etsy.com/search?q=lawn+care+business+spreadsheet) | WaveflowTemplates 12,224 · BuildBooksTemplates 4,091 · Dechampionz 414 · TheSheetBook 164 · KSMadeDigital 142 · MuseAndMakeCo 115 | **14/21** | Not yet (only a 4★ ease-of-use note) | 3 | 5 | 2 | 5 | **15** | **BRIEF: HOLD** (needs a complaint) |
| 4 | Pressure washing business | [search](https://www.etsy.com/search?q=pressure+washing+business+spreadsheet) | HandySheetTemplates 6,306 · SavvyandThriving · HeyMorning · TheProductivePlans · SellerMetric 142 | 12/24 | Adjacent only | 3 | 4 | 2 | 5 | 14 | Fold into the service-trade core (see below) |
| 5 | Hair stylist / booth renter books | [search](https://www.etsy.com/search?q=hair+stylist+income+expense+tracker) | BlushDesignTemplates 25,806 · SavvyandThriving 47,310 · TrackerPro 19,595 · GraceDigitalsph 899 · BoldlyCreatedGifts 431 · AnarosaStudioCo 189 | 4/16 | Yes, 2, both delivery failures | 4 | 1 | 4 | 4 | 13 | **BRIEF: PASS on evidence, price risk** |
| 6 | Candle making cost calculator | [search](https://www.etsy.com/search?q=candle+making+cost+calculator) | SavvyandThriving 47,310 · SerenatasJourney 38,648 · CraftCalculators 1,988 · WickSetGo 632 · AnnesArtDesignSpace 153 | 5/19 | Adjacent only (pricing calculator) | 4 | 2 | 2 | 5 | 13 | **BRIEF: HOLD** (needs a complaint) |
| 7 | Home daycare business | [search](https://www.etsy.com/search?q=home+daycare+business+spreadsheet) | Specialists are tiny: Digiplet 275 · SavageLabs 202 · SellerMetric 142 · FUNdamentalDaycare 27 · CrisBuildsStuff 4 · DaycareLedger 0 | **14/21** | Not searched in depth | 1 | 5 | 1 | 4 | 11 | Park: high prices, unproven sales |
| 8 | Photography client tracker | [search](https://www.etsy.com/search?q=photography+business+client+tracker+spreadsheet) | Mostly generalist shops (TrackerPro, SavvyandThriving, TheProductivePlans, BlushDesign) | 11/24 (many ads) | No | 2 | 3 | 1 | 4 | 10 | Park |
| 9 | Handmade product pricing | [search](https://www.etsy.com/search?q=handmade+pricing+calculator+spreadsheet) | HeyMorning 61,998 (bestseller $3.49) · PalmAndGrace 14,551 · SoloFinanceTools 322 · listing 1188596388 has 355 reviews at $1.91 | **1/8** | Adjacent: "fix the formulas" | 5 | 1 | 3 | 5 | 14 | **Kill for now**: a race to the bottom at $1–5 |
| 10 | Reseller inventory / profit | [search](https://www.etsy.com/search?q=reseller+inventory+profit+tracker+spreadsheet) | Generalists at $1–9 (SpreadsheetMall $2.75 bestseller, TemplateTrack $4.58 bestseller) | 5/24 | No | 3 | 1 | 1 | 4 | 9 | Kill for now |

## What the data says
1. **The money is in big generalist shops selling cheap.** HeyMorning (61,998 sales), TheProductivePlans (57,522), SavvyandThriving (47,310) and ProductivePenguinCo (38,382) show up in almost every niche, usually priced $3–9 on a permanent "sale". With a $12 floor we can't win on price, so we have to win on depth and on files that work.
2. **Most buyer complaints are failures our Checker step already catches:**
   - **The file doesn't open in Excel, or needs a Google login:** "I can't get anything in Excel and it won't let me do anything" (SavvyandThriving, cleaning bundle); "cant use without logging into my google account" (SavvyandThriving checkbook); "I need a template that works on excel. This is a google template" (TrackerPro).
   - **Downloads or links are broken, or the wrong files arrive:** ProductivePenguinCo, TheProductivePlans, TrackerPro, BlushDesignTemplates and AtlasWander all have a 1★ for this.
   - **The preview doesn't match the file:** "Not even close to what was shown in the post. No color, not graphs" (BuildBooksTemplates, recipe cost); "This is clearly an AI ad and what is delivered is a basic spreadsheet" (SellerMetric); "Do not go off the AI photos in the product page" (TheSheetBook).
   - **Formulas are wrong or missing:** "does not auto calculate and the formulas are wrong" (BuildBooksTemplates payroll); "I don't appreciate spreadsheets that I have to go and fix the formulas on" (HandySheetTemplates); "Nothing is automatic, everything is manual" (BuildBooksTemplates).
   - **Data has to be typed in twice:** "You have to enter in information over and over again no carry over" (SavvyandThriving cleaning bundle).
3. **Prices sit above the floor in service trades.** Lawn care (14/21 at ≥$12), pressure washing (12/24) and daycare (14/21) are the only niches where most sellers charge what we need. Lawn and pressure washing also share their whole engine with cleaning: client list → jobs/route → quote/invoice → monthly P&L.

## Recommendation
- **Build one tested "service-trade core"** (clients, jobs, invoices, expenses, P&L, all linked so nothing is entered twice), shipped in both native Excel and Google Sheets. Release it first as the **Cleaning Business System** (Brief 1), where the complaint evidence is strongest. Lawn care and pressure washing variants come next, for much less build effort.
- **Second product: Home Bakery Costing** (Brief 2). It has the biggest demand and the clearest fixable flaw (unit conversion between what you buy and what a recipe uses), but the price fight is harder, so it needs a bundle-depth offer to justify ≥$12.
- **Kill for now:** handmade pricing and reseller inventory. Both are crowded at $1–5.
- **Park:** daycare and photography. Daycare has the best prices, but no specialist seller shows real volume (the biggest has 275 shop sales).

## Gaps I'm flagging
- Etsy doesn't show per-listing sales on the pages I read, so "proven sellers" means shop-wide sales. Item review counts are the closest per-listing proxy I have.
- Etsy's total results count per search wasn't captured, so I have no competition-depth number per niche.
- Lawn care and candle have no niche-specific 1–3★ review yet. They stay on HOLD until the Scout finds one.
