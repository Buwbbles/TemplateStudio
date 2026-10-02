#!/usr/bin/env python3
"""Turn the Cleaning Business System build into a compact Google Sheets load plan.

The Sheets connector can only write values and batchUpdate requests (no .xlsx upload), so:
  * literal cells + the first formula of every same-pattern column run are written as values
    (USER_ENTERED, so "=..." strings become formulas);
  * the rest of each run is filled with a copyPaste (PASTE_FORMULA) request, which shifts
    relative references exactly like dragging a formula down.
Dates are written as serial numbers so Sheets and LibreOffice results compare 1:1.

Usage: python3 gsheets_export.py [sample|empty]  -> gsheets_plan_<variant>.json
"""
import datetime as dt
import json
import os
import re
import sys
from decimal import Decimal

import build_cleaning_system as B

HERE = os.path.dirname(os.path.abspath(__file__))
REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")


def shift(formula, dr):
    """Shift unanchored row numbers by dr (ignores text inside quotes)."""
    out, parts = [], re.split(r'("[^"]*")', formula)
    for p in parts:
        if p.startswith('"'):
            out.append(p)
            continue
        out.append(REF.sub(lambda m: f"{m.group(1)}{m.group(2)}{m.group(3)}"
                           f"{m.group(4) if m.group(3) else int(m.group(4)) + dr}", p))
    return "".join(out)


def cell_value(v):
    if isinstance(v, B.F):
        return "=" + str(v)
    if isinstance(v, dt.date):
        return (v - B.EPOCH).days
    if isinstance(v, Decimal):
        return float(v)
    return v


def plan(variant="sample"):
    if variant == "empty":
        cl, jb, ex, py = [], [], [], []
    else:
        cl, jb, ex, py = B.CLIENTS, B.JOBS, B.EXPENSES, B.sample_payments(B.CLIENTS, B.JOBS)
    sheets = [B.build_start_here(), B.build_settings(), B.build_clients(cl), B.build_jobs(jb),
              B.build_invoices(py), B.build_expenses(ex), B.build_pnl()]
    out = []
    for idx, sh in enumerate(sheets):
        cells = {k: v for k, (v, _s) in sh.cells.items() if v is not None and v != ""}
        max_r = max((r for r, _ in cells), default=1)
        max_c = max((c for _, c in cells), default=1)
        fills, written = [], set()
        # find vertical runs of the same relative formula
        for c in range(1, max_c + 1):
            r = 1
            while r <= max_r:
                v = cells.get((r, c))
                if isinstance(v, B.F):
                    end = r
                    while isinstance(cells.get((end + 1, c)), B.F) and shift(str(v), end + 1 - r) == str(cells[(end + 1, c)]):
                        end += 1
                    if end - r >= 2:
                        fills.append((c, r, end))
                        for rr in range(r + 1, end + 1):
                            written.add((rr, c))
                    r = end + 1
                else:
                    r += 1
        # merge adjacent-column fills with identical row spans into one copyPaste
        fills.sort(key=lambda t: (t[1], t[2], t[0]))
        merged = []
        for c, r0, r1 in fills:
            if merged and merged[-1][2] == r0 and merged[-1][3] == r1 and merged[-1][1] == c - 1:
                merged[-1][1] = c
            else:
                merged.append([c, c, r0, r1])
        # value rows: contiguous column span per row, only cells not covered by fills
        rows = {}
        for (r, c), v in cells.items():
            if (r, c) not in written:
                rows.setdefault(r, {})[c] = cell_value(v)
        blocks = []  # [start_row, start_col, [[...], ...]] grouping consecutive rows with same span
        for r in sorted(rows):
            cs = rows[r]
            c0, c1 = min(cs), max(cs)
            line = [cs.get(c, "") for c in range(c0, c1 + 1)]
            if blocks and blocks[-1]["r1"] == r - 1 and blocks[-1]["c0"] == c0 and blocks[-1]["c1"] == c1:
                blocks[-1]["rows"].append(line)
                blocks[-1]["r1"] = r
            else:
                blocks.append({"r0": r, "r1": r, "c0": c0, "c1": c1, "rows": [line]})
        out.append({"name": sh.name, "index": idx, "max_row": max_r, "max_col": max_c,
                    "freeze_rows": sh.freeze_rows,
                    "blocks": [{"range": f"'{sh.name}'!{B.col_letter(b['c0'])}{b['r0']}:"
                                         f"{B.col_letter(b['c1'])}{b['r1']}", "values": b["rows"]} for b in blocks],
                    "fills": [{"c0": a, "c1": b, "r0": r0, "r1": r1} for a, b, r0, r1 in merged]})
    return out


if __name__ == "__main__":
    variant = sys.argv[1] if len(sys.argv) > 1 else "sample"
    p = plan(variant)
    path = os.path.join(HERE, f"gsheets_plan_{variant}.json")
    with open(path, "w") as fh:
        json.dump(p, fh, separators=(",", ":"))
    for s in p:
        n = sum(len(r) for b in s["blocks"] for r in b["values"])
        print(s["name"], "blocks", len(s["blocks"]), "cells", n, "fills", len(s["fills"]),
              "bytes", len(json.dumps(s["blocks"], separators=(",", ":"))))
