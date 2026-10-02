#!/usr/bin/env python3
"""Verify the Cleaning Business System workbook headless in LibreOffice (standard library only).

What it does:
  1. Lints every formula in the generated file: only whitelisted Excel-2016 + Google Sheets functions.
  2. Scans all text for banned wording (no AI mentions in product files).
  3. Builds 4 variants with build_cleaning_system.build_workbook():
       sample        - the shipped file (sample data)
       empty         - sample inputs cleared exactly as Start Here tells the buyer
       empty_plus1   - cleared, then 1 new client + 1 completed job
       sample_plus1  - sample data + 1 new client + 1 completed job
  4. Recalculates each one with `soffice --headless --convert-to xlsx` using a throwaway profile set to
     "always recalculate on load", then reads the cached values LibreOffice wrote.
  5. Fails on any error value (#REF!, #VALUE!, #NAME?, #DIV/0!, #N/A, Err:xxx) in any formula cell, and checks
     key numbers against a Python re-computation of the same rules.

Usage:  python3 verify_cleaning_system.py            (writes VERIFY_REPORT.txt next to this script)
Exit code 0 = all checks passed.
"""
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_cleaning_system as B  # noqa: E402

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
ERRORS = {"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!"}
ALLOWED_FUNCS = {"IF", "IFERROR", "AND", "OR", "INDEX", "MATCH", "SUMIFS", "COUNTIFS", "COUNTIF", "SUM",
                 "ROUND", "ROWS", "TEXT", "DATE", "EOMONTH", "TODAY"}
BANNED_TEXT = re.compile(r"\bAI\b|artificial intelligence|chatgpt|openai|claude|\bLLM\b|machine learning",
                         re.IGNORECASE)
NEW_CLIENT = ("Lakeside Bakery", "(555) 123-4567", "orders@example.com", "18 Dock St, Springfield",
              "Standard Clean", None, "Weekly", "")
NEW_JOB = (dt.date(2026, 9, 25), "Lakeside Bakery", "", None, None, None, "Completed", "")

results = []


def check(name, ok, detail=""):
    results.append((bool(ok), name, detail))
    print(("PASS " if ok else "FAIL ") + name + (f"  [{detail}]" if detail else ""))
    return ok


def close(a, b, tol=0.005):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


# ---------------------------------------------------------------- xlsx reader
def read_xlsx(path):
    z = zipfile.ZipFile(path)
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(M + "si"):
            ss.append("".join(t.text or "" for t in si.iter(M + "t")))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rmap = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    sheets = {}
    for s in wb.find(M + "sheets"):
        target = rmap[s.get(R_ID)]
        part = target.lstrip("/") if target.startswith("/") else "xl/" + target
        cells = {}
        for c in ET.fromstring(z.read(part)).iter(M + "c"):
            t = c.get("t", "n")
            f, v, inl = c.find(M + "f"), c.find(M + "v"), c.find(M + "is")
            val = v.text if v is not None else None
            if t == "s" and val is not None:
                val = ss[int(val)]
            elif t == "inlineStr":
                val = "".join(x.text or "" for x in inl.iter(M + "t")) if inl is not None else ""
            elif t == "n" and val is not None:
                val = float(val)
            elif t == "str" and val is None:
                val = ""
            cells[c.get("r")] = {"t": t, "v": val, "f": f.text if f is not None else None, "has_f": f is not None}
        sheets[s.get("name")] = cells
    other_text = ""
    for n in z.namelist():
        if n.startswith("docProps/"):
            other_text += z.read(n).decode("utf-8", "replace")
    return sheets, other_text


def val(sheets, sheet, ref):
    c = sheets[sheet].get(ref)
    return None if c is None else c["v"]


def blank(v):
    return v is None or v == ""


# ---------------------------------------------------------------- LibreOffice recalc
def find_soffice():
    for name in ("soffice", "libreoffice", "libreoffice26.2", "/opt/libreoffice26.2/program/soffice",
                 "/usr/lib/libreoffice/program/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        p = shutil.which(name) or (name if os.path.isfile(name) else None)
        if p:
            return p
    return None


def make_profile(root):
    user = os.path.join(root, "user")
    os.makedirs(user, exist_ok=True)
    with open(os.path.join(user, "registrymodifications.xcu"), "w", encoding="utf-8") as fh:
        fh.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                 '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                 'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                 '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
                 '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
                 '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
                 '<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
                 '</oor:items>\n')
    return "file://" + os.path.abspath(root)


def recalc(soffice, profile_url, src, outdir):
    cmd = [soffice, f"-env:UserInstallation={profile_url}", "--headless", "--norestore", "--nologo",
           "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", outdir, src]
    p = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    out = os.path.join(outdir, os.path.basename(src))
    print("  $", " ".join(cmd))
    print("  ->", (p.stdout + p.stderr).strip()[:300])
    if p.returncode != 0 or not os.path.exists(out):
        raise RuntimeError(f"LibreOffice convert failed (rc={p.returncode}) for {src}")
    return out


# ---------------------------------------------------------------- generic scans
def scan_errors(sheets):
    bad, n_formula = [], 0
    for sname, cells in sheets.items():
        for ref, c in cells.items():
            if not c["has_f"]:
                continue
            n_formula += 1
            v = c["v"]
            if c["t"] == "e" or (isinstance(v, str) and (v in ERRORS or v.startswith("Err:"))):
                bad.append(f"{sname}!{ref}={v}")
    return n_formula, bad


def lint_formulas(sheets):
    funcs, bad = set(), []
    for sname, cells in sheets.items():
        for ref, c in cells.items():
            if not c["f"]:
                continue
            if "_xlfn" in c["f"]:
                bad.append(f"{sname}!{ref} uses _xlfn")
            body = re.sub(r'"[^"]*"', '""', c["f"])
            for fn in re.findall(r"([A-Z][A-Z0-9\.]*)\(", body):
                funcs.add(fn)
    return funcs, sorted(funcs - ALLOWED_FUNCS), bad


def scan_text(sheets, other):
    hits = []
    for sname, cells in sheets.items():
        for ref, c in cells.items():
            for s in (c["v"] if isinstance(c["v"], str) else "", c["f"] or ""):
                if s and BANNED_TEXT.search(s):
                    hits.append(f"{sname}!{ref}: {s[:60]}")
    if BANNED_TEXT.search(other):
        hits.append("docProps contains banned wording")
    return hits


# ---------------------------------------------------------------- expectations
def expected(clients, jobs, expenses, payments, today):
    amts = B.job_amounts(clients, jobs)
    terms = B.SETTINGS[6][1]
    prefix = B.SETTINGS[8][1]
    paid_by_inv = {}
    for (_d, inv, _m, a) in payments:
        paid_by_inv[inv] = paid_by_inv.get(inv, Decimal("0")) + Decimal(str(a))
    invoices = []
    for j in amts:
        if j["status"] != "Completed" or j["date"] is None or not j["client"]:
            continue
        inv = f"{prefix}{j['line']:04d}"
        paid = paid_by_inv.get(inv, Decimal("0"))
        bal = j["total"] - paid
        due = j["date"] + dt.timedelta(days=terms)
        status = "Paid" if bal <= 0 else ("Overdue" if today > due else ("Partial" if paid > 0 else "Unpaid"))
        invoices.append(dict(inv=inv, client=j["client"], total=j["total"], paid=paid, bal=bal, status=status,
                             overdue=(today - due).days if (bal > 0 and today > due) else ""))
    year = B.SETTINGS[7][1]
    months = {}
    for m in range(1, 13):
        rev = sum((j["subtotal"] for j in amts if j["status"] == "Completed" and j["date"].year == year
                   and j["date"].month == m), Decimal("0"))
        cnt = sum(1 for j in amts if j["status"] == "Completed" and j["date"].year == year and j["date"].month == m)
        exp = sum((Decimal(str(e[3])) for e in expenses if e[0].year == year and e[0].month == m), Decimal("0"))
        cash = sum((Decimal(str(p[3])) for p in payments if p[0].year == year and p[0].month == m), Decimal("0"))
        months[m] = dict(rev=rev, cnt=cnt, exp=exp, cash=cash)
    return amts, invoices, months


def check_variant(label, sheets, clients, jobs, expenses, payments, today):
    n_f, bad = scan_errors(sheets)
    check(f"{label}: formula cells recalculated", n_f > 10000, f"{n_f} formula cells")
    check(f"{label}: zero error values in formula cells", not bad, "; ".join(bad[:10]))
    amts, invs, months = expected(clients, jobs, expenses, payments, today)

    # Jobs: auto columns per filled row, blanks beyond
    ok_rows, msgs = True, []
    for j in amts:
        r = 4 + j["line"]
        if not close(val(sheets, "Jobs", f"K{r}"), j["rate"]) or not close(val(sheets, "Jobs", f"N{r}"), j["total"]):
            ok_rows = False
            msgs.append(f"row {r}: rate {val(sheets, 'Jobs', f'K{r}')} vs {j['rate']}, "
                        f"total {val(sheets, 'Jobs', f'N{r}')} vs {j['total']}")
        if not isinstance(val(sheets, "Jobs", f"A{r}"), float):
            ok_rows = False
            msgs.append(f"row {r}: date is not a real date value")
    check(f"{label}: Jobs rate/total/date match expected for {len(amts)} rows", ok_rows, "; ".join(msgs[:5]))
    first_blank = 5 + len(amts)
    nonblank = [f"{c}{r}={val(sheets, 'Jobs', f'{c}{r}')}" for r in range(first_blank, B.JB_LAST + 1)
                for c in "IJKLMNOP" if not blank(val(sheets, "Jobs", f"{c}{r}"))]
    check(f"{label}: empty Jobs rows show blank (I:P, rows {first_blank}-{B.JB_LAST})", not nonblank,
          "; ".join(nonblank[:5]))

    # Invoices
    listed = [r for r in range(5, B.JB_LAST + 1) if not blank(val(sheets, "Invoices", f"A{r}"))]
    check(f"{label}: invoice count", len(listed) == len(invs), f"{len(listed)} listed vs {len(invs)} expected")
    msgs = []
    for k, e in enumerate(invs):
        r = 5 + k
        got = {c: val(sheets, "Invoices", f"{c}{r}") for c in "ACEFGIJ"}
        if (got["A"] != e["inv"] or got["C"] != e["client"] or not close(got["E"], e["total"])
                or not close(got["F"], e["paid"]) or not close(got["G"], e["bal"]) or got["I"] != e["status"]
                or (e["overdue"] == "" and not blank(got["J"])) or (e["overdue"] != "" and not close(got["J"], e["overdue"]))):
            msgs.append(f"row {r}: got {got} expected {e}")
    check(f"{label}: every invoice's number/client/amount/paid/balance/status/days overdue", not msgs,
          "; ".join(msgs[:3]))
    nonblank = [f"{c}{r}" for r in range(5 + len(invs), B.JB_LAST + 1) for c in "ABCDEFGHIJ"
                if not blank(val(sheets, "Invoices", f"{c}{r}"))]
    check(f"{label}: empty invoice rows show blank", not nonblank, "; ".join(nonblank[:5]))

    # P&L
    msgs = []
    for m in range(1, 13):
        r = 4 + m
        e = months[m]
        for col, key in (("B", "cnt"), ("C", "rev"), ("E", "exp"), ("H", "cash")):
            if not close(val(sheets, "Monthly P&L", f"{col}{r}"), e[key]):
                msgs.append(f"{col}{r}={val(sheets, 'Monthly P&L', f'{col}{r}')} expected {e[key]}")
        if e["rev"] == 0 and not blank(val(sheets, "Monthly P&L", f"G{r}")):
            msgs.append(f"G{r} should be blank when no revenue")
    check(f"{label}: Monthly P&L jobs/revenue/expenses/cash by month", not msgs, "; ".join(msgs[:5]))
    tot_exp = sum((Decimal(str(e[3])) for e in expenses), Decimal("0"))
    check(f"{label}: category table total = expenses total", close(val(sheets, "Monthly P&L", "N36"), tot_exp),
          f"N36={val(sheets, 'Monthly P&L', 'N36')} vs {tot_exp}")

    # Clients roll-ups
    msgs = []
    for i, c in enumerate(clients):
        r = 5 + i
        billed = sum((v["total"] for v in invs if v["client"] == c[0]), Decimal("0"))
        bal = sum((v["bal"] for v in invs if v["client"] == c[0]), Decimal("0"))
        if not close(val(sheets, "Clients", f"J{r}"), billed) or not close(val(sheets, "Clients", f"K{r}"), bal):
            msgs.append(f"{c[0]}: billed {val(sheets, 'Clients', f'J{r}')} vs {billed}, "
                        f"balance {val(sheets, 'Clients', f'K{r}')} vs {bal}")
    check(f"{label}: Clients billed/balance roll-ups", not msgs, "; ".join(msgs[:3]))
    nonblank = [f"{c}{r}" for r in range(5 + len(clients), B.CL_LAST + 1) for c in "IJK"
                if not blank(val(sheets, "Clients", f"{c}{r}"))]
    check(f"{label}: empty client rows show blank", not nonblank, "; ".join(nonblank[:5]))
    return invs, months


def main():
    today = dt.date.today()
    soffice = find_soffice()
    if not check("LibreOffice found", soffice, soffice or "soffice not on PATH"):
        return finish()
    work = tempfile.mkdtemp(prefix="csb_verify_")
    profile = make_profile(os.path.join(work, "profile"))
    src_dir, out_dir = os.path.join(work, "src"), os.path.join(work, "out")
    os.makedirs(src_dir)
    os.makedirs(out_dir)

    shipped = B.build_workbook(B.DEFAULT_OUT)
    raw, other = read_xlsx(shipped)
    funcs, outside, xlfn = lint_formulas(raw)
    check("lint: only whitelisted functions", not outside and not xlfn,
          f"used={sorted(funcs)} outside={outside} {xlfn[:3]}")
    hits = scan_text(raw, other)
    check("text: no banned wording (AI mentions)", not hits, "; ".join(hits[:5]))
    check("tabs present", list(raw) == ["Start Here", "Settings", "Clients", "Jobs", "Invoices", "Expenses",
                                         "Monthly P&L"], str(list(raw)))

    pays = B.sample_payments(B.CLIENTS, B.JOBS)
    variants = {
        "sample": (B.CLIENTS, B.JOBS, B.EXPENSES, pays),
        "empty": ([], [], [], []),
        "empty_plus1": ([NEW_CLIENT], [NEW_JOB], [], []),
        "sample_plus1": (B.CLIENTS + [NEW_CLIENT], B.JOBS + [NEW_JOB], B.EXPENSES, pays),
    }
    got = {}
    for name, (cl, jb, ex, py) in variants.items():
        src = B.build_workbook(os.path.join(src_dir, f"{name}.xlsx"), clients=cl, jobs=jb, expenses=ex, payments=py)
        print(f"\n== {name}")
        sheets, _ = read_xlsx(recalc(soffice, profile, src, out_dir))
        got[name] = (sheets, check_variant(name, sheets, cl, jb, ex, py, today))

    # explicit add-one-client checks
    s, (invs, months) = got["empty_plus1"]
    check("empty_plus1: new job pulled client's address",
          val(s, "Jobs", "I5") == NEW_CLIENT[3], str(val(s, "Jobs", "I5")))
    check("empty_plus1: new invoice INV-0001 for Lakeside Bakery",
          val(s, "Invoices", "A5") == "INV-0001" and val(s, "Invoices", "C5") == "Lakeside Bakery",
          f"{val(s, 'Invoices', 'A5')} / {val(s, 'Invoices', 'C5')}")
    check("empty_plus1: Sep P&L revenue = 120.00", close(val(s, "Monthly P&L", "C13"), 120),
          str(val(s, "Monthly P&L", "C13")))
    s0, (_i0, m0) = got["sample"]
    s1, (i1, _m1) = got["sample_plus1"]
    check("sample_plus1: Lakeside Bakery is the last invoice",
          bool(i1) and val(s1, "Invoices", f"C{4 + len(i1)}") == "Lakeside Bakery",
          str(val(s1, "Invoices", f"C{4 + len(i1)}")))
    check("sample_plus1: Sep P&L revenue rose by exactly 120.00",
          close(float(val(s1, "Monthly P&L", "C13")) - float(val(s0, "Monthly P&L", "C13")), 120),
          f"{val(s0, 'Monthly P&L', 'C13')} -> {val(s1, 'Monthly P&L', 'C13')}")
    shutil.copy(os.path.join(out_dir, "sample.xlsx"), os.path.join(HERE, "recalc_check_sample_LibreOffice.xlsx"))
    return finish()


def finish():
    passed = sum(1 for ok, *_ in results if ok)
    lines = [f"Verify run {dt.datetime.now().isoformat(timespec='seconds')}  ({passed}/{len(results)} passed)"]
    lines += [("PASS " if ok else "FAIL ") + n + (f"  [{d}]" if d else "") for ok, n, d in results]
    with open(os.path.join(HERE, "VERIFY_REPORT.txt"), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")
    print("\n" + lines[0])
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    sys.exit(main())
