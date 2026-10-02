#!/usr/bin/env python3
"""Build the Cleaning Business System workbook (.xlsx) with the Python standard library only.

Usage:
  python3 build_cleaning_system.py                    # sample-data build -> CleaningBusinessSystem_v1.xlsx
  python3 build_cleaning_system.py --empty OUT.xlsx   # same workbook with the sample inputs cleared

No third-party packages (openpyxl is NOT required). The verify script imports build_workbook()
from this file to make test variants (empty, empty + new client, sample + new client).

Formula policy: only functions that behave the same in Excel 2016+ and Google Sheets:
IF, IFERROR, AND, OR, INDEX, MATCH, SUMIFS, COUNTIFS, COUNTIF, SUM, ROUND, ROWS, TEXT,
DATE, EOMONTH, TODAY. No XLOOKUP/LET/LAMBDA/FILTER/dynamic arrays/IFS/MAXIFS/macros.
"""
import argparse
import datetime as dt
import os
import re
import zipfile
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_OUT = os.path.join(HERE, "CleaningBusinessSystem_v1.xlsx")

# ---------------------------------------------------------------- capacity (first/last data rows)
CL_FIRST, CL_LAST = 5, 204      # 200 clients
JB_FIRST, JB_LAST = 5, 504      # 500 jobs
EX_FIRST, EX_LAST = 5, 504      # 500 expenses
PY_FIRST, PY_LAST = 5, 1004     # 1000 payments
LS_FIRST, LS_LAST = 5, 19       # 15 services / categories on Settings
PM_FIRST, PM_LAST = 5, 12       # 8 payment methods

STATUSES = "Scheduled,Completed,Cancelled"
FREQUENCIES = "One-time,Weekly,Every 2 weeks,Monthly"

# ---------------------------------------------------------------- settings + lists
SETTINGS = [  # (label, value, kind) -> Settings!A5:B13
    ("Business name", "Sparkle & Shine Cleaning Co.", "text"),
    ("Owner / contact name", "Your Name", "text"),
    ("Phone", "(555) 010-2040", "text"),
    ("Email", "hello@example.com", "text"),
    ("Sales tax rate (0% if you don't charge tax)", 0.06, "pct"),
    ("Currency symbol (shown in headers)", "$", "text"),
    ("Payment terms (days until an invoice is due)", 14, "int"),
    ("P&L year", 2026, "int"),
    ("Invoice number prefix", "INV-", "text"),
]
# Absolute references used by formulas (keep in sync with SETTINGS order above)
S_NAME, S_TAX, S_CUR, S_TERMS, S_YEAR, S_PREFIX = (
    "Settings!$B$5", "Settings!$B$9", "Settings!$B$10", "Settings!$B$11", "Settings!$B$12", "Settings!$B$13")

SERVICES = [  # (service, standard rate, basis)
    ("Standard Clean", 120, "per visit"),
    ("Deep Clean", 220, "per visit"),
    ("Move-In / Move-Out", 320, "per visit"),
    ("Recurring Weekly", 95, "per visit"),
    ("Recurring Every 2 Weeks", 110, "per visit"),
    ("Office Clean (hourly)", 45, "per hour - enter hours in Qty"),
    ("Window Cleaning", 85, "per visit"),
    ("Post-Construction", 400, "per visit"),
]
CATEGORIES = ["Cleaning Supplies", "Equipment", "Fuel & Mileage", "Wages & Contractors", "Insurance",
              "Marketing & Advertising", "Phone & Software", "Vehicle Maintenance", "Licenses & Fees", "Other"]
METHODS = ["Cash", "Check", "Card", "Bank Transfer", "Payment App", "Other"]

# ---------------------------------------------------------------- sample data (fictional)
d = dt.date
CLIENTS = [  # name, phone, email, address, default service, custom rate, frequency, notes
    ("Hannah Whitfield", "(555) 201-3344", "hannah.w@example.com", "412 Maple Ridge Dr, Springfield",
     "Recurring Every 2 Weeks", None, "Every 2 weeks", "Gate code 4471. Two friendly dogs."),
    ("Marcus & Elena Ortiz", "(555) 318-7702", "ortiz.home@example.com", "88 Birchwood Ln, Springfield",
     "Recurring Weekly", 90, "Weekly", "Loyalty price. Key in lockbox."),
    ("Brightpath Dental Office", "(555) 440-1180", "frontdesk@example.com",
     "1500 Commerce Pkwy, Suite 210, Springfield", "Office Clean (hourly)", None, "Weekly",
     "After 6 pm only."),
    ("Priya Raman", "(555) 562-9015", "priya.r@example.com", "27 Lakeview Ct, Riverton",
     "Deep Clean", None, "Monthly", "Fragrance-free products only."),
    ("Tom Kessler", "(555) 673-2208", "tkessler@example.com", "9 Old Mill Rd, Riverton",
     "Standard Clean", None, "Monthly", ""),
    ("Greenleaf Property Mgmt", "(555) 784-5530", "units@example.com", "300 Harbor St, Springfield",
     "Move-In / Move-Out", 300, "One-time", "Rental turnovers. Unit number goes in job notes."),
    ("Danielle Brooks", "(555) 895-6127", "dani.brooks@example.com", "1742 Sunset Ave, Springfield",
     "Standard Clean", None, "Every 2 weeks", "Prefers Friday mornings."),
    ("Riverside Yoga Studio", "(555) 906-3471", "hello.riverside@example.com", "55 River Rd, Riverton",
     "Office Clean (hourly)", 40, "Weekly", "Mop studio floors with water only."),
]
JOBS = [  # date, client, service ("" = client's usual), qty, extras, rate override, status, notes
    (d(2026, 7, 2), "Hannah Whitfield", "", None, None, None, "Completed", ""),
    (d(2026, 7, 3), "Brightpath Dental Office", "", 3, None, None, "Completed", ""),
    (d(2026, 7, 6), "Marcus & Elena Ortiz", "", None, None, None, "Completed", ""),
    (d(2026, 7, 8), "Priya Raman", "", None, 25, None, "Completed", "Extra: inside fridge"),
    (d(2026, 7, 10), "Brightpath Dental Office", "", 3, None, None, "Completed", ""),
    (d(2026, 7, 13), "Marcus & Elena Ortiz", "", None, None, None, "Completed", ""),
    (d(2026, 7, 16), "Hannah Whitfield", "", None, None, None, "Completed", ""),
    (d(2026, 7, 20), "Greenleaf Property Mgmt", "", None, None, None, "Completed", "Unit 4B move-out"),
    (d(2026, 7, 24), "Tom Kessler", "", None, None, None, "Completed", ""),
    (d(2026, 7, 28), "Riverside Yoga Studio", "", 2.5, None, None, "Completed", ""),
    (d(2026, 8, 3), "Marcus & Elena Ortiz", "", None, None, None, "Completed", ""),
    (d(2026, 8, 4), "Riverside Yoga Studio", "", 2.5, None, None, "Completed", ""),
    (d(2026, 8, 6), "Danielle Brooks", "Deep Clean", None, None, None, "Completed", "First visit - deep clean"),
    (d(2026, 8, 7), "Brightpath Dental Office", "", 3, None, None, "Completed", ""),
    (d(2026, 8, 10), "Marcus & Elena Ortiz", "", None, None, None, "Completed", ""),
    (d(2026, 8, 13), "Hannah Whitfield", "", None, None, None, "Completed", ""),
    (d(2026, 8, 14), "Greenleaf Property Mgmt", "Move-In / Move-Out", None, 40, None, "Completed",
     "Unit 2A move-in. Extra: carpet spot treatment"),
    (d(2026, 8, 18), "Priya Raman", "Standard Clean", None, None, None, "Completed", "Lighter clean this month"),
    (d(2026, 8, 21), "Tom Kessler", "", None, None, None, "Cancelled", "Client rescheduled"),
    (d(2026, 8, 27), "Danielle Brooks", "", None, None, None, "Completed", ""),
    (d(2026, 9, 1), "Marcus & Elena Ortiz", "", None, None, None, "Completed", ""),
    (d(2026, 9, 3), "Brightpath Dental Office", "", 3.5, None, None, "Completed", ""),
    (d(2026, 9, 8), "Hannah Whitfield", "", None, None, None, "Completed", ""),
    (d(2026, 9, 10), "Riverside Yoga Studio", "", 2.5, None, None, "Completed", ""),
    (d(2026, 9, 12), "Greenleaf Property Mgmt", "", None, None, None, "Completed", "Unit 7C move-out"),
    (d(2026, 9, 15), "Priya Raman", "", None, None, None, "Completed", ""),
    (d(2026, 9, 17), "Tom Kessler", "Window Cleaning", None, None, 95, "Completed", "Second-story windows"),
    (d(2026, 9, 22), "Hannah Whitfield", "", None, None, None, "Completed", ""),
    (d(2026, 9, 29), "Marcus & Elena Ortiz", "", None, None, None, "Scheduled", ""),
    (d(2026, 9, 30), "Danielle Brooks", "", None, None, None, "Scheduled", ""),
]
EXPENSES = [  # date, category, description, amount, paid with, notes
    (d(2026, 7, 1), "Cleaning Supplies", "Bulk cleaners and microfiber cloths", 186.40, "Card", ""),
    (d(2026, 7, 5), "Insurance", "General liability - monthly", 89.00, "Bank Transfer", ""),
    (d(2026, 7, 15), "Marketing & Advertising", "Door hangers (500)", 64.99, "Card", ""),
    (d(2026, 7, 31), "Fuel & Mileage", "Fuel - July", 142.75, "Card", ""),
    (d(2026, 8, 5), "Insurance", "General liability - monthly", 89.00, "Bank Transfer", ""),
    (d(2026, 8, 12), "Equipment", "HEPA canister vacuum", 329.00, "Card", "Keep receipt for taxes"),
    (d(2026, 8, 31), "Fuel & Mileage", "Fuel - August", 151.20, "Card", ""),
    (d(2026, 9, 5), "Insurance", "General liability - monthly", 89.00, "Bank Transfer", ""),
    (d(2026, 9, 9), "Phone & Software", "Scheduling app + business phone line", 45.00, "Card", ""),
    (d(2026, 9, 20), "Cleaning Supplies", "Restock - glass cleaner, mop heads", 112.35, "Card", ""),
]
UNPAID_LINES = {9, 25, 26, 27, 28}   # job line numbers left unpaid in the sample
PARTIAL_LINES = {17: 200.00}        # job line -> partial amount paid


# ---------------------------------------------------------------- python replica of the sheet math
def money(x):
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def job_amounts(clients, jobs, services=SERVICES, tax_rate=SETTINGS[4][1]):
    """Same rules as the Jobs formulas. Returns list of dicts (line is 1-based)."""
    cl = {c[0]: c for c in clients}
    svc = {s[0]: s[1] for s in services}
    out = []
    for i, (date, client, service, qty, extras, override, status, _n) in enumerate(jobs, start=1):
        c = cl.get(client)
        used = service if service else (c[4] if c else "")
        if override is not None:
            rate = override
        elif c and c[5] is not None and (service == "" or service == c[4]):
            rate = c[5]
        else:
            rate = svc.get(used, 0)
        sub = money(Decimal(str(rate)) * Decimal(str(qty if qty is not None else 1)) + Decimal(str(extras or 0)))
        tax = money(sub * Decimal(str(tax_rate)))
        out.append(dict(line=i, date=date, client=client, service=used, rate=rate, subtotal=sub, tax=tax,
                        total=sub + tax, status=status))
    return out


def sample_payments(clients, jobs, prefix="INV-"):
    pays = []
    for j in job_amounts(clients, jobs):
        n = j["line"]
        if j["status"] != "Completed" or n in UNPAID_LINES:
            continue
        amt = Decimal(str(PARTIAL_LINES[n])) if n in PARTIAL_LINES else j["total"]
        pays.append((j["date"] + dt.timedelta(days=2 + n % 9), f"{prefix}{n:04d}", METHODS[n % 4], float(amt)))
    pays.sort(key=lambda p: p[0])
    return pays


# ---------------------------------------------------------------- tiny xlsx writer
class F(str):
    """A formula (without the leading '=')."""


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;"))


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    return int(m.group(2)), col_num(m.group(1))


EPOCH = dt.date(1899, 12, 30)

# style ids (index into CELL_XFS below)
(ST_DEFAULT, ST_TITLE, ST_NOTE, ST_HDR_IN, ST_HDR_AUTO, ST_IN_TEXT, ST_IN_DATE, ST_IN_MONEY, ST_IN_NUM,
 ST_AUTO_TEXT, ST_AUTO_DATE, ST_AUTO_MONEY, ST_AUTO_INT, ST_AUTO_PCT, ST_BOLD, ST_SET_TEXT, ST_SET_PCT,
 ST_SET_MONEY, ST_SET_INT, ST_AUTO_MONTH, ST_TOT_MONEY, ST_TOT_LABEL, ST_PLAIN, ST_SECTION, ST_TOT_INT,
 ST_KPI_MONEY, ST_HELPER, ST_TOT_PCT, ST_KPI_INT) = range(29)

# (numFmtId, fontId, fillId, borderId, alignment-xml or "")
WRAPC = '<alignment horizontal="center" vertical="center" wrapText="1"/>'
CELL_XFS = [
    (0, 0, 0, 0, ""), (0, 1, 0, 0, ""), (0, 3, 0, 0, ""), (0, 2, 2, 1, WRAPC), (0, 2, 3, 1, WRAPC),
    (49, 0, 0, 1, ""), (14, 0, 0, 1, ""), (4, 0, 0, 1, ""), (0, 0, 0, 1, ""),
    (0, 0, 4, 1, ""), (14, 0, 4, 1, ""), (165, 0, 4, 1, ""), (1, 0, 4, 1, ""), (10, 0, 4, 1, ""),
    (0, 4, 0, 0, ""), (49, 0, 5, 1, ""), (10, 0, 5, 1, ""), (4, 0, 5, 1, ""), (1, 0, 5, 1, ""),
    (164, 4, 4, 1, ""), (165, 4, 4, 1, ""), (0, 4, 4, 1, ""), (0, 0, 0, 0, ""), (0, 5, 0, 0, ""),
    (1, 4, 4, 1, ""), (165, 4, 0, 0, ""), (0, 3, 0, 0, ""), (10, 4, 4, 1, ""), (1, 4, 0, 0, ""),
]

STYLES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="mmm yyyy"/>'
    '<numFmt numFmtId="165" formatCode="#,##0.00;[Red]\\-#,##0.00"/></numFmts>'
    '<fonts count="6">'
    '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="16"/><color rgb="FF1F4E5A"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
    '<font><i/><sz val="10"/><color rgb="FF666666"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="12"/><color rgb="FF1F6F78"/><name val="Calibri"/><family val="2"/></font>'
    '</fonts>'
    '<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F6F78"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF5F6B6D"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F4"/><bgColor indexed="64"/></patternFill></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFFF6D5"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    '<border><left style="thin"><color rgb="FFD0D7D9"/></left><right style="thin"><color rgb="FFD0D7D9"/></right>'
    '<top style="thin"><color rgb="FFD0D7D9"/></top><bottom style="thin"><color rgb="FFD0D7D9"/></bottom>'
    '<diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="%d">%s</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="3">'
    '<dxf><font><b/><color rgb="FFB00020"/></font><fill><patternFill><bgColor rgb="FFFDE2E2"/></patternFill></fill></dxf>'
    '<dxf><font><color rgb="FF2E7D32"/></font><fill><patternFill><bgColor rgb="FFE3F4E5"/></patternFill></fill></dxf>'
    '<dxf><font><color rgb="FF9A6700"/></font><fill><patternFill><bgColor rgb="FFFFF3CD"/></patternFill></fill></dxf>'
    '</dxfs><tableStyles count="0"/></styleSheet>'
) % (len(CELL_XFS), "".join(
    '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"%s%s%s%s>%s</xf>' % (
        n, fo, fi, b,
        ' applyNumberFormat="1"' if n else "", ' applyFont="1"' if fo else "",
        ' applyFill="1"' if fi else "", ' applyBorder="1"' if b else "", al)
    if al else
    '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"%s%s%s%s/>' % (
        n, fo, fi, b,
        ' applyNumberFormat="1"' if n else "", ' applyFont="1"' if fo else "",
        ' applyFill="1"' if fi else "", ' applyBorder="1"' if b else "")
    for (n, fo, fi, b, al) in CELL_XFS))
# alignment xfs need applyAlignment
STYLES_XML = STYLES_XML.replace('"><alignment', '" applyAlignment="1"><alignment')


class Sheet:
    def __init__(self, name, tab="FF1F6F78"):
        self.name, self.tab = name, tab
        self.cells, self.widths, self.hidden_cols, self.heights = {}, {}, set(), {}
        self.validations, self.cond = [], []
        self.freeze_rows, self.gridlines = 0, True

    def set(self, ref, value, style=0):
        self.cells[split_ref(ref)] = (value, style)

    def width(self, col, w, hidden=False):
        self.widths[col_num(col)] = w
        if hidden:
            self.hidden_cols.add(col_num(col))

    def list_dv(self, sqref, source, prompt=""):
        self.validations.append(("list", sqref, source, None, prompt, "Pick a value from the dropdown list."))

    def dv(self, kind, sqref, f1, f2=None, prompt="", error="", operator=None):
        self.validations.append((kind, sqref, f1, (f2, operator), prompt, error))

    def _cell_xml(self, r, c, value, style):
        ref = f"{col_letter(c)}{r}"
        s = f' s="{style}"' if style else ""
        if value is None:
            return f'<c r="{ref}"{s}/>'
        if isinstance(value, F):
            return f'<c r="{ref}"{s}><f>{esc(value)}</f></c>'
        if isinstance(value, dt.date):
            return f'<c r="{ref}"{s}><v>{(value - EPOCH).days}</v></c>'
        if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
            return f'<c r="{ref}"{s}><v>{value}</v></c>'
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{esc(value)}</t></is></c>'

    def xml(self, selected=False):
        rows = {}
        for (r, c), (v, s) in self.cells.items():
            rows.setdefault(r, []).append((c, v, s))
        maxr = max(rows) if rows else 1
        maxc = max(c for (_r, c) in self.cells) if self.cells else 1
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
               f'<sheetPr><tabColor rgb="{self.tab}"/></sheetPr>',
               f'<dimension ref="A1:{col_letter(maxc)}{maxr}"/>']
        tab_sel = ' tabSelected="1"' if selected else ""
        grid = "" if self.gridlines else ' showGridLines="0"'
        sv = f'<sheetView workbookViewId="0"{tab_sel}{grid}>'
        if self.freeze_rows:
            top = self.freeze_rows + 1
            sv += (f'<pane ySplit="{self.freeze_rows}" topLeftCell="A{top}" activePane="bottomLeft" state="frozen"/>'
                   f'<selection pane="bottomLeft" activeCell="A{top}" sqref="A{top}"/>')
        else:
            sv += '<selection activeCell="A1" sqref="A1"/>'
        out.append(f'<sheetViews>{sv}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>')
        if self.widths:
            hid = ' hidden="1"'
            out.append("<cols>" + "".join(
                f'<col min="{c}" max="{c}" width="{w}" customWidth="1"{hid if c in self.hidden_cols else ""}/>'
                for c, w in sorted(self.widths.items())) + "</cols>")
        out.append("<sheetData>")
        for r in sorted(rows):
            ht = f' ht="{self.heights[r]}" customHeight="1"' if r in self.heights else ""
            out.append(f'<row r="{r}"{ht}>' + "".join(self._cell_xml(r, c, v, s)
                                                      for c, v, s in sorted(rows[r], key=lambda t: t[0])) + "</row>")
        out.append("</sheetData>")
        if self.cond:
            prio = 1
            for sqref, rules in self.cond:
                parts = []
                for (op, formula, dxf) in rules:
                    parts.append(f'<cfRule type="cellIs" dxfId="{dxf}" priority="{prio}" operator="{op}">'
                                 f'<formula>{esc(formula)}</formula></cfRule>')
                    prio += 1
                out.append(f'<conditionalFormatting sqref="{sqref}">{"".join(parts)}</conditionalFormatting>')
        if self.validations:
            out.append(f'<dataValidations count="{len(self.validations)}">')
            for kind, sqref, f1, extra, prompt, error in self.validations:
                attrs = f'type="{kind}" allowBlank="1" showInputMessage="1" showErrorMessage="1"'
                f2 = None
                if extra:
                    f2, op = extra
                    if op:
                        attrs += f' operator="{op}"'
                if error:
                    attrs += f' errorTitle="Check this entry" error="{esc(error)}"'
                if prompt:
                    attrs += f' promptTitle="Tip" prompt="{esc(prompt)}"'
                body = f"<formula1>{esc(f1)}</formula1>" + (f"<formula2>{esc(f2)}</formula2>" if f2 is not None else "")
                out.append(f'<dataValidation {attrs} sqref="{sqref}">{body}</dataValidation>')
            out.append("</dataValidations>")
        out.append('<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
        out.append('<pageSetup orientation="landscape" fitToHeight="0"/></worksheet>')
        return "".join(out)


def header_row(sh, row, specs):
    """specs: list of (col, text_or_formula, is_input, width)."""
    for col, text, is_input, w in specs:
        sh.set(f"{col}{row}", text, ST_HDR_IN if is_input else ST_HDR_AUTO)
        sh.width(col, w)
    sh.heights[row] = 32


def cur(label):
    """Header text with the currency symbol from Settings, e.g. Rate ($)."""
    return F(f'"{label} ("&{S_CUR}&")"')


DATE_MIN, DATE_MAX = 36526, 73050   # 2000-01-01 .. 2099-12-31
DATE_ERR = "Enter a real date, for example 7/14/2026 (or your usual date format)."
NUM_ERR = "Enter a number of 0 or more (no currency symbols or text)."


# ---------------------------------------------------------------- sheets
def build_start_here():
    sh = Sheet("Start Here", "FF2E7D32")
    sh.gridlines = False
    sh.width("A", 3)
    sh.width("B", 118)
    lines = [
        ("t", "Cleaning Business System - Start Here"),
        ("n", "Clients, jobs, invoices, payments, expenses and a monthly profit report in one file. "
              "Type each thing once - the rest fills in for you."),
        ("", ""),
        ("h", "HOW THE COLORS WORK"),
        ("", "Teal column headers = you type here (many cells have dropdown lists)."),
        ("", "Dark gray headers and light gray cells = automatic. Please don't type over them."),
        ("", "Yellow cells on the Settings tab = your business details and price list."),
        ("", ""),
        ("h", "SET UP IN 5 STEPS"),
        ("", "1. Settings: enter your business name, sales tax rate (0% if you don't charge tax), payment terms and P&L year."),
        ("", "2. Settings: list your services and your standard price for each one (room for 15)."),
        ("", "3. Clients: add each client once - name, address, usual service, and a custom price if they have one."),
        ("", "4. Jobs: add a row for every job. Pick the date, pick the client, and set Status to Completed when it's done."),
        ("", "5. Invoices: completed jobs show up as invoices. When a client pays, log it in the Payments Log on the same tab."),
        ("", "    The Monthly P&L tab then adds up revenue, expenses, profit and cash collected for each month."),
        ("", ""),
        ("h", "YOUR FIRST JOB IN 60 SECONDS"),
        ("", "1. Clients tab: in the first empty row type a name (for example Jane Test) and an address, "
             "then pick Standard Clean under Usual Service."),
        ("", "2. Jobs tab: in the first empty row type today's date (Ctrl + ; in Excel), pick Jane Test under Client,"),
        ("", "    and pick Completed under Status."),
        ("", "3. Look right on the same row: address, price, tax, total and invoice number are already filled in."),
        ("", "4. Invoices tab: Jane Test's invoice is at the end of the list with its due date and status."),
        ("", "5. Monthly P&L tab: this month's revenue went up by that job (if the P&L year on Settings is this year)."),
        ("", "    Done testing? Clear the cells you typed on Clients and Jobs and everything goes back."),
        ("", ""),
        ("h", "HOW TO CLEAR THE SAMPLE DATA (all formulas stay in place)"),
        ("", "Select each range below and press the Delete key. Use Delete only - not 'Delete Row' or 'Delete Sheet'."),
        ("", "    Clients tab:   A5:H204"),
        ("", "    Jobs tab:   A5:H504"),
        ("", "    Expenses tab:   A5:F504"),
        ("", "    Invoices tab (Payments Log):   M5:P1004"),
        ("", "Quick way: click the Name Box (left of the formula bar), type the range, press Enter, then press Delete."),
        ("", "Then replace the sample business details and prices on the Settings tab with your own."),
        ("", ""),
        ("h", "GOOD TO KNOW"),
        ("", "Dates: type real dates like 7/14/2026. The file refuses text that isn't a date, so due dates and "
             "days overdue always work."),
        ("", "Leave Service blank on a job to use the client's usual service and custom price. Pick a different "
             "service to charge its standard price."),
        ("", "Rate Override on a job beats every other price. Qty is visits or hours (blank = 1). Extras adds a "
             "flat amount, e.g. inside oven."),
        ("", "Only Completed jobs become invoices and count as revenue. Scheduled and Cancelled jobs are kept "
             "for your records."),
        ("", "Add new rows at the bottom of each list. Don't sort the Jobs tab or delete job rows - invoice "
             "numbers come from the job's line."),
        ("", "    To drop a job, set its Status to Cancelled instead."),
        ("", "Client names must be unique (the file will warn you). Avoid * and ? in client names."),
        ("", "Room for 200 clients, 500 jobs, 500 expenses and 1,000 payments."),
    ]
    r = 1
    for kind, text in lines:
        style = {"t": ST_TITLE, "n": ST_NOTE, "h": ST_SECTION}.get(kind, ST_PLAIN)
        if text:
            sh.set(f"B{r}", text, style)
        r += 1
    sh.heights[1] = 24
    return sh


def build_settings():
    sh = Sheet("Settings", "FFB8860B")
    sh.set("A1", "Settings", ST_TITLE)
    sh.set("A2", "Yellow cells are yours to change. Everything else in the file reads from here.", ST_NOTE)
    header_row(sh, 4, [("A", "Setting", True, 44), ("B", "Your value", True, 30), ("C", "", True, 3),
                       ("D", "Service type", True, 28), ("E", cur("Standard rate"), True, 15),
                       ("F", "Rate basis / notes", True, 30), ("G", "", True, 3),
                       ("H", "Expense categories", True, 26), ("I", "", True, 3),
                       ("J", "Payment methods", True, 20)])
    for col in ("C", "G", "I"):
        sh.set(f"{col}4", None, 0)
    for i, (label, value, kind) in enumerate(SETTINGS):
        r = 5 + i
        sh.set(f"A{r}", label, ST_BOLD)
        sh.set(f"B{r}", value, {"text": ST_SET_TEXT, "pct": ST_SET_PCT, "int": ST_SET_INT}[kind])
    sh.dv("decimal", "B9", "0", "1", operator="between", prompt="Type a percent, e.g. 6%. Use 0% if you don't charge tax.",
          error="Enter a percent between 0% and 100%.")
    sh.dv("whole", "B11", "0", "365", operator="between", prompt="Days a client has to pay, e.g. 14.",
          error="Enter whole days between 0 and 365.")
    sh.dv("whole", "B12", "2000", "2099", operator="between", prompt="Year shown on the Monthly P&L tab.",
          error="Enter a 4-digit year, e.g. 2026.")
    for i in range(LS_LAST - LS_FIRST + 1):
        r = LS_FIRST + i
        s = SERVICES[i] if i < len(SERVICES) else (None, None, None)
        sh.set(f"D{r}", s[0], ST_SET_TEXT)
        sh.set(f"E{r}", s[1], ST_SET_MONEY)
        sh.set(f"F{r}", s[2], ST_SET_TEXT)
        sh.set(f"H{r}", CATEGORIES[i] if i < len(CATEGORIES) else None, ST_SET_TEXT)
        if r <= PM_LAST:
            sh.set(f"J{r}", METHODS[i] if i < len(METHODS) else None, ST_SET_TEXT)
    sh.dv("decimal", f"E{LS_FIRST}:E{LS_LAST}", "0", operator="greaterThanOrEqual", error=NUM_ERR)
    sh.set("D21", "Renaming a service? Update the jobs and clients that use the old name too.", ST_NOTE)
    sh.freeze_rows = 4
    return sh


def build_clients(clients):
    sh = Sheet("Clients")
    sh.set("A1", F(f'{S_NAME}&" - Clients"'), ST_TITLE)
    sh.set("A2", "One row per client. Jobs, invoices and balances look up this list - enter each client only once.",
           ST_NOTE)
    header_row(sh, 4, [("A", "Client name", True, 28), ("B", "Phone", True, 16), ("C", "Email", True, 28),
                       ("D", "Service address", True, 40), ("E", "Usual service", True, 24),
                       ("F", cur("Custom rate (optional)"), True, 14), ("G", "Frequency", True, 15),
                       ("H", "Notes", True, 36), ("I", "Completed jobs", False, 11),
                       ("J", cur("Total billed"), False, 14), ("K", cur("Balance owed"), False, 14)])
    for i in range(CL_LAST - CL_FIRST + 1):
        r = CL_FIRST + i
        c = clients[i] if i < len(clients) else (None,) * 8
        for col, v, st in zip("ABCDEFGH", c, (ST_IN_TEXT, ST_IN_TEXT, ST_IN_TEXT, ST_IN_TEXT, ST_IN_TEXT,
                                              ST_IN_MONEY, ST_IN_TEXT, ST_IN_TEXT)):
            sh.set(f"{col}{r}", v if v != "" else None, st)
        sh.set(f"I{r}", F(f'IF($A{r}="","",COUNTIFS(Jobs!$B${JB_FIRST}:$B${JB_LAST},$A{r},'
                          f'Jobs!$G${JB_FIRST}:$G${JB_LAST},"Completed"))'), ST_AUTO_INT)
        sh.set(f"J{r}", F(f'IF($A{r}="","",SUMIFS(Invoices!$E${JB_FIRST}:$E${JB_LAST},'
                          f'Invoices!$C${JB_FIRST}:$C${JB_LAST},$A{r}))'), ST_AUTO_MONEY)
        sh.set(f"K{r}", F(f'IF($A{r}="","",SUMIFS(Invoices!$G${JB_FIRST}:$G${JB_LAST},'
                          f'Invoices!$C${JB_FIRST}:$C${JB_LAST},$A{r}))'), ST_AUTO_MONEY)
    sh.dv("custom", f"A{CL_FIRST}:A{CL_LAST}", f"COUNTIF($A${CL_FIRST}:$A${CL_LAST},A{CL_FIRST})=1",
          prompt="Each client name must be unique - add a last name or location if two clients share a name.",
          error="This client name is already on the list. Use a unique name.")
    sh.list_dv(f"E{CL_FIRST}:E{CL_LAST}", f"Settings!$D${LS_FIRST}:$D${LS_LAST}",
               prompt="The service this client usually books. Jobs use it when Service is left blank.")
    sh.dv("decimal", f"F{CL_FIRST}:F{CL_LAST}", "0", operator="greaterThanOrEqual",
          prompt="Optional. A special price for this client's usual service. Leave blank to use the Settings price.",
          error=NUM_ERR)
    sh.list_dv(f"G{CL_FIRST}:G{CL_LAST}", f'"{FREQUENCIES}"')
    sh.freeze_rows = 4
    return sh


def build_jobs(jobs):
    sh = Sheet("Jobs")
    sh.set("A1", F(f'{S_NAME}&" - Jobs"'), ST_TITLE)
    sh.set("A2", "Type in the teal columns (A-H). Gray columns fill in by themselves. "
                 "Only Completed jobs become invoices.", ST_NOTE)
    header_row(sh, 4, [("A", "Date", True, 12), ("B", "Client", True, 26), ("C", "Service (blank = usual)", True, 24),
                       ("D", "Qty / hours (blank = 1)", True, 11), ("E", cur("Extras"), True, 11),
                       ("F", cur("Rate override"), True, 11), ("G", "Status", True, 13), ("H", "Notes", True, 30),
                       ("I", "Address", False, 36), ("J", "Service used", False, 24), ("K", cur("Rate"), False, 11),
                       ("L", cur("Subtotal"), False, 12), ("M", cur("Tax"), False, 10), ("N", cur("Total"), False, 12),
                       ("O", "Invoice #", False, 12), ("P", "Payment status", False, 13),
                       ("Q", "Invoice line (helper)", False, 8)])
    sh.width("Q", 8, hidden=True)
    CA, CD, CE, CF_ = (f"Clients!$A${CL_FIRST}:$A${CL_LAST}", f"Clients!$D${CL_FIRST}:$D${CL_LAST}",
                       f"Clients!$E${CL_FIRST}:$E${CL_LAST}", f"Clients!$F${CL_FIRST}:$F${CL_LAST}")
    SD, SE = f"Settings!$D${LS_FIRST}:$D${LS_LAST}", f"Settings!$E${LS_FIRST}:$E${LS_LAST}"
    for i in range(JB_LAST - JB_FIRST + 1):
        r = JB_FIRST + i
        j = jobs[i] if i < len(jobs) else (None,) * 8
        for col, v, st in zip("ABCDEFGH", j, (ST_IN_DATE, ST_IN_TEXT, ST_IN_TEXT, ST_IN_NUM, ST_IN_MONEY,
                                              ST_IN_MONEY, ST_IN_TEXT, ST_IN_TEXT)):
            sh.set(f"{col}{r}", v if v != "" else None, st)
        m = f"MATCH($B{r},{CA},0)"
        sh.set(f"I{r}", F(f'IF($B{r}="","",IFERROR(INDEX({CD},{m})&"",""))'), ST_AUTO_TEXT)
        sh.set(f"J{r}", F(f'IF($B{r}="","",IF($C{r}<>"",$C{r},IFERROR(INDEX({CE},{m})&"","")))'), ST_AUTO_TEXT)
        sh.set(f"K{r}", F(f'IF($B{r}="","",IF($F{r}<>"",$F{r},IFERROR(IF(AND(INDEX({CF_},{m})&""<>"",'
                          f'OR($C{r}="",$C{r}=INDEX({CE},{m}))),INDEX({CF_},{m}),'
                          f'INDEX({SE},MATCH($J{r},{SD},0))),0)))'), ST_AUTO_MONEY)
        sh.set(f"L{r}", F(f'IF($B{r}="","",$K{r}*IF($D{r}="",1,$D{r})+$E{r})'), ST_AUTO_MONEY)
        sh.set(f"M{r}", F(f'IF($B{r}="","",ROUND($L{r}*{S_TAX},2))'), ST_AUTO_MONEY)
        sh.set(f"N{r}", F(f'IF($B{r}="","",$L{r}+$M{r})'), ST_AUTO_MONEY)
        sh.set(f"O{r}", F(f'IF($Q{r}="","",{S_PREFIX}&TEXT(ROWS($A${JB_FIRST}:$A{r}),"0000"))'), ST_AUTO_TEXT)
        sh.set(f"P{r}", F(f'IF($O{r}="","",IFERROR(INDEX(Invoices!$I${JB_FIRST}:$I${JB_LAST},'
                          f'MATCH($O{r},Invoices!$A${JB_FIRST}:$A${JB_LAST},0)),""))'), ST_AUTO_TEXT)
        sh.set(f"Q{r}", F(f'IF(AND($A{r}<>"",$B{r}<>"",$G{r}="Completed"),COUNTIFS($A${JB_FIRST}:$A{r},"<>",'
                          f'$B${JB_FIRST}:$B{r},"<>",$G${JB_FIRST}:$G{r},"Completed"),"")'), ST_HELPER)
    sh.dv("date", f"A{JB_FIRST}:A{JB_LAST}", str(DATE_MIN), str(DATE_MAX), operator="between",
          prompt="Job date, e.g. 7/14/2026. In Excel, Ctrl + ; types today's date.", error=DATE_ERR)
    sh.list_dv(f"B{JB_FIRST}:B{JB_LAST}", CA, prompt="Pick a client. New client? Add them on the Clients tab first.")
    sh.list_dv(f"C{JB_FIRST}:C{JB_LAST}", SD,
               prompt="Leave blank to use the client's usual service and price, or pick a different service.")
    sh.dv("decimal", f"D{JB_FIRST}:F{JB_LAST}", "0", operator="greaterThanOrEqual", error=NUM_ERR)
    sh.list_dv(f"G{JB_FIRST}:G{JB_LAST}", f'"{STATUSES}"',
               prompt="Set to Completed when the job is done - that creates the invoice.")
    sh.cond.append((f"P{JB_FIRST}:P{JB_LAST}", [("equal", '"Overdue"', 0), ("equal", '"Paid"', 1),
                                                 ("equal", '"Partial"', 2)]))
    sh.freeze_rows = 4
    return sh


def build_invoices(payments):
    sh = Sheet("Invoices", "FF8E44AD")
    sh.set("A1", F(f'{S_NAME}&" - Invoices & Payments"'), ST_TITLE)
    sh.set("A2", "Total billed", ST_BOLD)
    sh.set("B2", F(f"SUM(E{JB_FIRST}:E{JB_LAST})"), ST_KPI_MONEY)
    sh.set("C2", "Collected", ST_BOLD)
    sh.set("D2", F(f"SUM(F{JB_FIRST}:F{JB_LAST})"), ST_KPI_MONEY)
    sh.set("E2", "Outstanding", ST_BOLD)
    sh.set("F2", F(f"SUM(G{JB_FIRST}:G{JB_LAST})"), ST_KPI_MONEY)
    sh.set("G2", "Overdue invoices", ST_BOLD)
    sh.set("H2", F(f'COUNTIFS(I{JB_FIRST}:I{JB_LAST},"Overdue")'), ST_KPI_INT)
    sh.set("A3", "Left: every Completed job becomes an invoice automatically (don't type here). "
                 "Right: log each payment you receive.", ST_NOTE)
    header_row(sh, 4, [("A", "Invoice #", False, 12), ("B", "Invoice date", False, 12), ("C", "Client", False, 26),
                       ("D", "Service", False, 22), ("E", cur("Amount due"), False, 13), ("F", cur("Paid"), False, 12),
                       ("G", cur("Balance"), False, 12), ("H", "Due date", False, 12), ("I", "Status", False, 11),
                       ("J", "Days overdue", False, 10), ("K", "Job line (helper)", False, 8),
                       ("L", "", True, 3), ("M", "Payment date", True, 12), ("N", "Invoice #", True, 13),
                       ("O", "Paid with", True, 15), ("P", cur("Amount"), True, 12), ("Q", "Client", False, 26)])
    sh.set("L4", None, 0)
    sh.width("K", 8, hidden=True)
    PAY_AMT, PAY_INV, PAY_DATE = (f"$P${PY_FIRST}:$P${PY_LAST}", f"$N${PY_FIRST}:$N${PY_LAST}",
                                  f"$M${PY_FIRST}:$M${PY_LAST}")
    for i in range(JB_LAST - JB_FIRST + 1):
        r = JB_FIRST + i
        sh.set(f"K{r}", F(f'IFERROR(MATCH(ROWS($K${JB_FIRST}:$K{r}),Jobs!$Q${JB_FIRST}:$Q${JB_LAST},0),"")'), ST_HELPER)
        for col, src, st in (("A", "O", ST_AUTO_TEXT), ("B", "A", ST_AUTO_DATE), ("C", "B", ST_AUTO_TEXT),
                             ("D", "J", ST_AUTO_TEXT), ("E", "N", ST_AUTO_MONEY)):
            sh.set(f"{col}{r}", F(f'IF($K{r}="","",INDEX(Jobs!${src}${JB_FIRST}:${src}${JB_LAST},$K{r}))'), st)
        sh.set(f"F{r}", F(f'IF($A{r}="","",SUMIFS({PAY_AMT},{PAY_INV},$A{r}))'), ST_AUTO_MONEY)
        sh.set(f"G{r}", F(f'IF($A{r}="","",$E{r}-$F{r})'), ST_AUTO_MONEY)
        sh.set(f"H{r}", F(f'IF($A{r}="","",$B{r}+{S_TERMS})'), ST_AUTO_DATE)
        sh.set(f"I{r}", F(f'IF($A{r}="","",IF($G{r}<=0,"Paid",IF(TODAY()>$H{r},"Overdue",'
                          f'IF($F{r}>0,"Partial","Unpaid"))))'), ST_AUTO_TEXT)
        sh.set(f"J{r}", F(f'IF($A{r}="","",IF(AND($G{r}>0,TODAY()>$H{r}),TODAY()-$H{r},""))'), ST_AUTO_INT)
    for i in range(PY_LAST - PY_FIRST + 1):
        r = PY_FIRST + i
        p = payments[i] if i < len(payments) else (None,) * 4
        for col, v, st in zip("MNOP", p, (ST_IN_DATE, ST_IN_TEXT, ST_IN_TEXT, ST_IN_MONEY)):
            sh.set(f"{col}{r}", v, st)
        sh.set(f"Q{r}", F(f'IF($N{r}="","",IFERROR(INDEX($C${JB_FIRST}:$C${JB_LAST},'
                          f'MATCH($N{r},$A${JB_FIRST}:$A${JB_LAST},0)),"Invoice not found"))'), ST_AUTO_TEXT)
    sh.dv("date", f"M{PY_FIRST}:M{PY_LAST}", str(DATE_MIN), str(DATE_MAX), operator="between",
          prompt="Date the money arrived. Counts toward Cash Collected on the Monthly P&L.", error=DATE_ERR)
    sh.list_dv(f"N{PY_FIRST}:N{PY_LAST}", f"$A${JB_FIRST}:$A${JB_LAST}",
               prompt="Pick the invoice this payment is for. Part payments: add one row per payment.")
    sh.list_dv(f"O{PY_FIRST}:O{PY_LAST}", f"Settings!$J${PM_FIRST}:$J${PM_LAST}")
    sh.dv("decimal", f"P{PY_FIRST}:P{PY_LAST}", "0", operator="greaterThanOrEqual", error=NUM_ERR)
    sh.cond.append((f"I{JB_FIRST}:I{JB_LAST}", [("equal", '"Overdue"', 0), ("equal", '"Paid"', 1),
                                                 ("equal", '"Partial"', 2)]))
    sh.freeze_rows = 4
    return sh


def build_expenses(expenses):
    sh = Sheet("Expenses", "FFC0392B")
    sh.set("A1", F(f'{S_NAME}&" - Expenses"'), ST_TITLE)
    sh.set("A2", "Total on this list", ST_BOLD)
    sh.set("B2", F(f"SUM(D{EX_FIRST}:D{EX_LAST})"), ST_KPI_MONEY)
    sh.set("A3", "One row per business expense. Categories come from the Settings tab.", ST_NOTE)
    header_row(sh, 4, [("A", "Date", True, 12), ("B", "Category", True, 26), ("C", "Description / vendor", True, 40),
                       ("D", cur("Amount"), True, 12), ("E", "Paid with", True, 15), ("F", "Notes", True, 30)])
    for i in range(EX_LAST - EX_FIRST + 1):
        r = EX_FIRST + i
        e = expenses[i] if i < len(expenses) else (None,) * 6
        for col, v, st in zip("ABCDEF", e, (ST_IN_DATE, ST_IN_TEXT, ST_IN_TEXT, ST_IN_MONEY, ST_IN_TEXT, ST_IN_TEXT)):
            sh.set(f"{col}{r}", v if v != "" else None, st)
    sh.dv("date", f"A{EX_FIRST}:A{EX_LAST}", str(DATE_MIN), str(DATE_MAX), operator="between",
          prompt="Date of the expense, e.g. 7/14/2026.", error=DATE_ERR)
    sh.list_dv(f"B{EX_FIRST}:B{EX_LAST}", f"Settings!$H${LS_FIRST}:$H${LS_LAST}")
    sh.dv("decimal", f"D{EX_FIRST}:D{EX_LAST}", "0", operator="greaterThanOrEqual", error=NUM_ERR)
    sh.list_dv(f"E{EX_FIRST}:E{EX_LAST}", f"Settings!$J${PM_FIRST}:$J${PM_LAST}")
    sh.freeze_rows = 4
    return sh


def build_pnl():
    sh = Sheet("Monthly P&L", "FF1F4E5A")
    sh.set("A1", F(f'{S_NAME}&" - Monthly Profit & Loss "&{S_YEAR}'), ST_TITLE)
    sh.set("A2", "Revenue = Completed jobs by job date (before tax). Cash collected = payments by payment date. "
                 "Change the year on Settings.", ST_NOTE)
    header_row(sh, 4, [("A", "Month", False, 14), ("B", "Jobs completed", False, 11),
                       ("C", cur("Revenue (before tax)"), False, 14), ("D", cur("Sales tax collected"), False, 13),
                       ("E", cur("Expenses"), False, 13), ("F", cur("Net profit"), False, 13),
                       ("G", "Profit margin", False, 10), ("H", cur("Cash collected"), False, 13)])
    JA, JG = f"Jobs!$A${JB_FIRST}:$A${JB_LAST}", f"Jobs!$G${JB_FIRST}:$G${JB_LAST}"
    EA, ED = f"Expenses!$A${EX_FIRST}:$A${EX_LAST}", f"Expenses!$D${EX_FIRST}:$D${EX_LAST}"
    PM, PA = f"Invoices!$M${PY_FIRST}:$M${PY_LAST}", f"Invoices!$P${PY_FIRST}:$P${PY_LAST}"
    for m in range(1, 13):
        r = 4 + m
        win = lambda rng: f'{rng},">="&$A{r},{rng},"<="&EOMONTH($A{r},0)'
        sh.set(f"A{r}", F(f"DATE({S_YEAR},{m},1)"), ST_AUTO_MONTH)
        sh.set(f"B{r}", F(f'COUNTIFS({JG},"Completed",{win(JA)})'), ST_AUTO_INT)
        sh.set(f"C{r}", F(f'SUMIFS(Jobs!$L${JB_FIRST}:$L${JB_LAST},{JG},"Completed",{win(JA)})'), ST_AUTO_MONEY)
        sh.set(f"D{r}", F(f'SUMIFS(Jobs!$M${JB_FIRST}:$M${JB_LAST},{JG},"Completed",{win(JA)})'), ST_AUTO_MONEY)
        sh.set(f"E{r}", F(f"SUMIFS({ED},{win(EA)})"), ST_AUTO_MONEY)
        sh.set(f"F{r}", F(f"$C{r}-$E{r}"), ST_AUTO_MONEY)
        sh.set(f"G{r}", F(f'IF($C{r}=0,"",$F{r}/$C{r})'), ST_AUTO_PCT)
        sh.set(f"H{r}", F(f"SUMIFS({PA},{win(PM)})"), ST_AUTO_MONEY)
    sh.set("A17", "Year total", ST_TOT_LABEL)
    sh.set("B17", F("SUM(B5:B16)"), ST_TOT_INT)
    for col in "CDEFH":
        sh.set(f"{col}17", F(f"SUM({col}5:{col}16)"), ST_TOT_MONEY)
    sh.set("G17", F('IF($C17=0,"",$F17/$C17)'), ST_TOT_PCT)

    sh.set("A19", "Expenses by category", ST_SECTION)
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    sh.set("A20", "Category", ST_HDR_AUTO)
    for k, name in enumerate(months):
        sh.set(f"{col_letter(2 + k)}20", name, ST_HDR_AUTO)
    sh.set("N20", "Year total", ST_HDR_AUTO)
    sh.heights[20] = 20
    sh.width("I", 12)
    for k in range(9, 14):
        sh.width(col_letter(k + 1), 12)
    for i in range(LS_LAST - LS_FIRST + 1):
        r = 21 + i
        srow = LS_FIRST + i
        sh.set(f"A{r}", F(f'IF(Settings!$H${srow}="","",Settings!$H${srow})'), ST_AUTO_TEXT)
        for k in range(12):
            col = col_letter(2 + k)
            start = f"DATE({S_YEAR},{k + 1},1)"
            sh.set(f"{col}{r}", F(f'IF($A{r}="","",SUMIFS({ED},Expenses!$B${EX_FIRST}:$B${EX_LAST},$A{r},'
                                  f'{EA},">="&{start},{EA},"<="&EOMONTH({start},0)))'), ST_AUTO_MONEY)
        sh.set(f"N{r}", F(f'IF($A{r}="","",SUM(B{r}:M{r}))'), ST_TOT_MONEY)
    last = 21 + LS_LAST - LS_FIRST
    sh.set(f"A{last + 1}", "Total", ST_TOT_LABEL)
    for k in range(13):
        col = col_letter(2 + k)
        sh.set(f"{col}{last + 1}", F(f"SUM({col}21:{col}{last})"), ST_TOT_MONEY)
    sh.set(f"A{last + 2}", "If this total is lower than Expenses above, some expenses have a blank or unlisted category.",
           ST_NOTE)
    sh.freeze_rows = 4
    return sh


# ---------------------------------------------------------------- package
def build_workbook(out_path, clients=None, jobs=None, expenses=None, payments=None):
    clients = CLIENTS if clients is None else clients
    jobs = JOBS if jobs is None else jobs
    expenses = EXPENSES if expenses is None else expenses
    payments = sample_payments(CLIENTS, JOBS) if payments is None else payments
    sheets = [build_start_here(), build_settings(), build_clients(clients), build_jobs(jobs),
              build_invoices(payments), build_expenses(expenses), build_pnl()]
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          '<Override PartName="/docProps/core.xml" '
          'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    for i in range(1, len(sheets) + 1)) + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
            'core-properties" Target="docProps/core.xml"/></Relationships>')
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>Cleaning Business System</dc:title><dc:creator>Template Studio</dc:creator>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
    wb = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{ns_main}" xmlns:r="{ns_r}">'
          '<workbookPr/><bookViews><workbookView activeTab="0"/></bookViews><sheets>'
          + "".join(f'<sheet name="{esc(s.name)}" sheetId="{i}" r:id="rId{i}"/>' for i, s in enumerate(sheets, 1))
          + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                         f'relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
               + f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                 'relationships/styles" Target="styles.xml"/></Relationships>')
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, s in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", s.xml(selected=(i == 1)))
    return out_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("out", nargs="?", default=DEFAULT_OUT)
    ap.add_argument("--empty", action="store_true", help="clear Clients, Jobs, Expenses and Payments inputs")
    a = ap.parse_args()
    if a.empty:
        build_workbook(a.out, clients=[], jobs=[], expenses=[], payments=[])
    else:
        build_workbook(a.out)
    print("wrote", a.out)


if __name__ == "__main__":
    main()
