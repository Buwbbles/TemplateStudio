# Rooms and crew — 2026-10-02 (all 7 charter workers on the roster)

All workers: Sonnet 5 (medium). Each needs its OWN desk (WORKSTATION) in its room.
Props: WORKSTATION=compute, DISH=web, INTEL CAB=files, WORKBENCH=terminal, SERVER CART=memory, STUDIO=image analysis.
Google Sheets/Drive/Docs connectors are account-level: no prop needed.

| Room (charter) | Rename from | Agent | Props |
|---|---|---|---|
| 1. Tool Foundry | "Builder" | BUILDER | desk, INTEL CAB, WORKBENCH, SERVER CART |
| 1. Tool Foundry | | PACKAGER | desk, DISH, INTEL CAB, SERVER CART |
| 2. Research & Strategy | "Strategy" | SCOUT | desk, DISH, INTEL CAB, SERVER CART |
| 2. Research & Strategy | | AUDITOR | desk, DISH, INTEL CAB, SERVER CART |
| 3. Quality & Compliance | "Checker" | CHECKER (Gatekeeper) | desk, DISH, INTEL CAB, WORKBENCH, STUDIO, SERVER CART |
| 4. Operations & Customer Desk | new (a HAB room) | OPERATOR | desk, DISH, INTEL CAB, SERVER CART |
| 5. The Ledger | new (a HAB room) | LEDGER CLERK | desk, INTEL CAB, WORKBENCH, SERVER CART |
| HOME | | LUCILLE | unchanged |

Conveyors: none needed; LUCILLE routes every ticket.

## Tickets dispatched 2026-10-02
T-004 BUILDER preview fixes ($0.60) · T-005 PACKAGER listing package ($0.30) · T-006 SCOUT close holds ($0.50)
T-007 LEDGER CLERK ledger sheet ($0.40) · T-008 OPERATOR reply templates ($0.25)
Queued: T-009 CHECKER Foundry Gate on previews + listing (after T-004/T-005) · LUCILLE: render PNGs, create Etsy draft.
AUDITOR idle until first listing is live 30 days or 3+ live listings (nothing to audit yet).

## Status update 2026-10-02 08:22
- T-004 BUILDER: done ($0.88). Previews fixed; PNGs rendered by LUCILLE to products/cleaning-business-system/previews/.
- T-005 PACKAGER: done ($0.27, Opus). listing-package.json + shop-copy.md copied to the product folder. Title 117 chars, 13 tags OK, no Excel, $14.99.
- T-006 SCOUT: first run blocked (fetch tool read the output filename "x.md" as a website; .md is a country domain). Retry running on Sonnet.
- T-007 LEDGER CLERK: blocked (Sheets connector unavailable in its run). LUCILLE built the ledger: sheet 1xxnj6Jl4FI8JI6ISpajNPzzZ_ZOuKjuLommE9MYQ9jM. Fee check: $14.99 -> fees $1.87, net $13.12; $12.99 -> $1.68, net $11.31. Samples cleared. Compute tab seeded: $33.92 month to date, $26.79 of it LUCILLE on Opus.
- T-008 OPERATOR: file written (6 templates) before a 45s dispatch timeout; copied to ops/reply-templates.md. Goes through Gate.
- T-009 CHECKER: Foundry Gate running on previews + listing + shop copy + templates.
Lesson: never put "name.md" in a worker prompt; say "markdown file named X".
