"""Daily car sales report: Input -> adjust -> output.

  1. reads every raw export in Input (sale, order, stock) and works out which part of the report it is
  2. fixes it and writes the fixed rows to adjust (one sheet per target sheet of the main workbook)
  3. applies the rows to a COPY of the main workbook, saved in output with a QC report:
     sales are added below the rows already there, orders and stock replace the rows of their sheet

The workbook in Main is only read, never changed.
"""
import datetime as dt
import os
import re
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

import openpyxl
import xlrd
from openpyxl.styles import Font

ROOT = Path(__file__).resolve().parent
INPUT, ADJUST, MAIN, OUTPUT = ROOT / "Input", ROOT / "adjust", ROOT / "Main", ROOT / "output"

# ---- settings ---------------------------------------------------------------------------------
# sheet name -> header row in the main workbook (data starts on the next row)
SHEETS = {"KDC-CUS": 5, "KDC X-CUS": 5, "TOUK2+Other": 2}
# Maker -> sheet. Any other maker (e.g. KR MOTORS motorcycles) goes to OTHER_SHEET and is pointed out in the QC report.
OTHER_SHEET = "TOUK2+Other"
MAKER_SHEET = {"KIA": "KDC-CUS", "HYUNDAI": "KDC-CUS", "DAEHAN": "KDC-CUS", "TERACO": "KDC-CUS",
               "CHANGAN": "KDC X-CUS", "NEVO": "KDC X-CUS", "GEELY": "KDC X-CUS"}
MODEL_RENAME = {("HYUNDAI", "TUCSON"): "NEW TUCSON"}  # (Maker, Model in export) -> name the report counts
NOT_REPORTED = ("EV TUK TUK",)  # CarType of sales the report does not carry: left out, and taken out if already in the workbook
KEY = "Frame#"
# Every export a full day should bring, for the checklist in the QC report:
# (kind, source, sheets of the main workbook whose columns identify the export, brands sharing those columns)
# (NETA stands before MBS: a row naming both, e.g. a NETA showroom with a -KLM- booking number, is NETA's)
BRANDS = {"Geely": ("GEELY", "GE TOPIA", "-GEE", "-GSVK"), "BMW": ("BMW", "PREMIUM LAO AUTO", "LX TOPIA"),
          "NETA": ("NETA",), "MBS": ("MITSUBISHI", "MITSU", "MBS", "-KLM-")}
# brands with no download of their own: their rows come inside another brand's file, so that file without
# such rows means the brand has none today
INSIDE = {"BMW": "Geely"}
# parts with no download to get: not counted as missing (the sheet stays as it is in the main workbook)
NOT_EXPECTED = []
# brands no longer sold: the export still lists their cars, which are taken out and applied nowhere
DROPPED = ("NETA",)
LAYOUTS = [("Sale", "KDC", [("KDC-CUS", 5)], None), ("Order", "KDC", [("SR ORDER STATUS", 2)], None),
           ("Stock", "KDC", [("stock kdc", 2)], None),
           ("Sale", "Dealer", [("DL-CUTOSMER", 5)], None), ("Order", "Dealer", [("DL ORDER STATUS", 2)], None),
           ("Stock", "Dealer", [("STOCK DL", 2)], None),
           ("Sale", "brand", [("Geely sale", 2), ("BMW Sale", 2), ("MBS Sale", 2)], BRANDS),
           ("Order", "brand", [("Geely order", 2), ("BMW Order", 2), ("MBS Order", 2)], BRANDS),
           ("Stock", "brand", [("Geely stock", 2), ("BMW Stock", 2), ("MBS Stock", 2)], BRANDS)]
# where a brand can be read in the rows; the order exports have no Maker column, so the showroom and booking number are used
BRAND_COLUMNS = ("maker", "showroom", "srm", "warehouse", "sitecode", "bookno", "booking#", "s-booking#")
# parts the script applies; the others are only checked for presence: (how, sheet, header row, key column)
#   "add"     = rows whose key is not in the sheet yet are added below the others (sales build up over the month)
#   "replace" = the rows of the sheet are replaced by the export, which is the full list of today (orders, stock)
# sheet None = split by maker: sales over SHEETS, orders over ORDER_SHEET
ORDER_SHEET = {"KDC-CUS": "SR ORDER STATUS", "KDC X-CUS": "KDCX ORDER", "TOUK2+Other": "SR ORDER STATUS"}
# (column, column to take the value from when it is empty). The order export gives a Lead / Booking order only a
# BookDate, but the report counts orders by Pre Date, so the finished report carries the BookDate there too.
# The dealer order sheet is not filled this way: there a Booking order keeps an empty Pre Date in the finished report.
FILL_FROM = {("Order", "KDC"): [("Pre Date", "BookDate")]}
CANCEL_COLS = ("Cancel Reason", "Cancel Date")  # an order with one of these filled is cancelled and is left out
TARGETS = {("Sale", "KDC"): ("add", None, 5, "Frame#"), ("Order", "KDC"): ("replace", None, 2, "BookNo"),
           ("Stock", "KDC"): ("replace", "stock kdc", 2, "Frame#"),
           ("Sale", "Dealer"): ("add", "DL-CUTOSMER", 5, "Frame#"), ("Order", "Dealer"): ("replace", "DL ORDER STATUS", 2, "BookNo"),
           ("Stock", "Dealer"): ("replace", "STOCK DL", 2, "Frame#"),
           ("Sale", "Geely"): ("add", "Geely sale", 2, "Frame#"), ("Sale", "BMW"): ("add", "BMW Sale", 2, "Frame#"),
           ("Sale", "MBS"): ("add", "MBS Sale", 2, "Frame#"),
           ("Order", "Geely"): ("replace", "Geely order", 2, "BookNo"), ("Order", "BMW"): ("replace", "BMW Order", 2, "BookNo"),
           ("Order", "MBS"): ("replace", "MBS Order", 2, "BookNo"),
           ("Stock", "Geely"): ("replace", "Geely stock", 2, "Frame#"), ("Stock", "BMW"): ("replace", "BMW Stock", 2, "Frame#"),
           ("Stock", "MBS"): ("replace", "MBS Stock", 2, "Frame#")}
# test run "--until=2026-10-05": the date columns that decide whether a row is later than that day and is left out
UNTIL_COLS = {"Sale": ("sales date",), "Order": ("pre date", "bookdate"), "Stock": ("import date", "imp date")}
MATCH_MIN = 0.7                        # share of column names that must match to recognise an export
MONTH_ROW = ("By Model.DATE - with MOS", 4)  # (sheet, row) holding the day columns; they tell which month the workbook is for
DATE_COL = "Sales date"
# (sheet, range, file name, columns to hide for this picture only ...): hidden rows/columns stay out of the picture
PICTURES = [("Daily Report", "F1:AJ94", "Daily Report"),
            ("KOLAOTOPIA byModel.DAET", "A1:BX102", "KOLAOTOPIA by model"),
            ("By Model.DATE - with MOS", "A4:CJ101", "By model with MOS"),
            ("By Model.DATE - with MOS", "B1:CD101", "By model sales", "I:K")]  # the same sheet without the stock columns
REFILTER = ["KOLAOTOPIA byModel.DAET"]   # sheets whose saved filter is applied again after the new numbers are in
BLANK_HEADERS ={15: "Company"}          # column P has no title on some sheets; position counted from 0
MAX_COLS = 150                          # how far right to read a sheet's header row
HELPER_SCAN = 8                          # formula columns may sit this far right of the last header
QC_COLS = ("Showroom", "Model", "Sales date", "Selling AMT", "Status", "Cust#")
QC_CELLS = [("Daily Report", "O31", "KDC units this month"), ("Daily Report", "AD31", "KDC amount this month"),
            ("Daily Report", "O34", "CHANGAN units this month"), ("Daily Report", "O67", "New car sales total"),
            ("By Model.DATE - with MOS", "AG8", "By-model month total")]
# -----------------------------------------------------------------------------------------------


def text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return " ".join(str(v).split())


def read_raw(path):
    """Return (header, rows). Each row is a list of (value, is_date) with dates as Excel serial numbers."""
    if path.suffix.lower() == ".xls":
        sh = xlrd.open_workbook(str(path), logfile=open(os.devnull, "w")).sheet_by_index(0)
        grid = [[(c.value if c.ctype not in (0, 6) and c.value != "" else None, c.ctype == 3) for c in sh.row(r)]
                for r in range(sh.nrows)]
    else:
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
        epoch = dt.datetime(1899, 12, 30)
        grid = []
        for row in ws.iter_rows(values_only=True):
            cells = []
            for v in row:
                if isinstance(v, dt.datetime):
                    cells.append(((v - epoch).total_seconds() / 86400, True))
                elif isinstance(v, dt.date):
                    cells.append((float((v - epoch.date()).days), True))
                else:
                    cells.append((v if v != "" else None, False))
            grid.append(cells)
    # the header is the first row near the top that is mostly text (a stock export has fewer header cells than a
    # car row has text cells, so the fullest row is not always it); some exports put part of it on the row below,
    # under a group title such as 'Car Info'
    is_text = lambda cell: isinstance(cell[0], str) and bool(cell[0].strip()) and not cell[1]
    if not grid:
        return [], []
    counts = [sum(map(is_text, grid[i])) for i in range(min(10, len(grid)))]
    hr = next(i for i, c in enumerate(counts) if c * 2 >= max(counts))
    top = [text(v) for v, _ in grid[hr]]
    nxt = grid[hr + 1] if hr + 1 < len(grid) else []
    below = [text(v) for v, _ in nxt]
    filled = [i for i, b in enumerate(below) if b]
    overlap = sum(1 for i in filled if i < len(top) and top[i])  # the group title sits over the first name below it
    two_rows = (bool(filled) and all(is_text(nxt[i]) for i in filled) and any(i >= len(top) or not top[i] for i in filled)
                and (len(filled) < sum(1 for t in top if t) or overlap <= 1))
    width = max(len(top), len(below)) if two_rows else len(top)
    top += [""] * (width - len(top))
    header = [(below[i] if two_rows and i < len(below) and below[i] else top[i]) for i in range(width)]
    rows = [row + [(None, False)] * (width - len(row)) for row in grid[hr + (2 if two_rows else 1):]]
    return header, [r for r in rows if any(text(v) for v, _ in r)]


def main_signatures(main_path):
    """The header names of every sheet used to recognise an export: {(kind, source): [set of names, ...]}."""
    wb = openpyxl.load_workbook(main_path, read_only=True, data_only=True)
    sig = {}
    for kind, source, sheets, _ in LAYOUTS:
        for name, hr in sheets:
            if name in wb.sheetnames:
                row = next(wb[name].iter_rows(min_row=hr, max_row=hr, max_col=MAX_COLS, values_only=True), ())
                names = {text(v).lower() for v in row if text(v)}
                if names:
                    sig.setdefault((kind, source), []).append(names)
    wb.close()
    return sig


def report_month(main_path):
    """The (year, month) the main workbook is laid out for, from its day columns. None if it cannot be read."""
    wb = openpyxl.load_workbook(main_path, read_only=True, data_only=True)
    sheet, r = MONTH_ROW
    row = next(wb[sheet].iter_rows(min_row=r, max_row=r, max_col=MAX_COLS, values_only=True), ()) if sheet in wb.sheetnames else ()
    wb.close()
    months = Counter((v.year, v.month) for v in row if isinstance(v, dt.datetime))
    return months.most_common(1)[0][0] if months else None


def other_months(header, fixed, month):
    """Count the export rows whose sales date is outside the workbook's month: {'Nov 2026': 12, ...}."""
    if month is None or DATE_COL not in header:
        return Counter()
    at, epoch, found = header.index(DATE_COL), dt.datetime(1899, 12, 30), Counter()
    for rows in fixed.values():
        for row in rows:
            v, is_date = row[at]
            if is_date and v is not None:
                day = epoch + dt.timedelta(days=v)
                if (day.year, day.month) != month:
                    found[f"{day:%b %Y}"] += 1
    return found


def classify(path, header, rows, sig):
    """Which part of the report is this export? Returns (kind, source) or None, from its columns and its rows."""
    names = {text(h).lower() for h in header if text(h)}
    if not names:
        return None
    score = {k: max(len(names & s) / max(len(names), len(s)) for s in sets) for k, sets in sig.items()}
    best = max(score, key=score.get, default=None)
    if best is None or score[best] < MATCH_MIN:
        return None
    brands = next(b for kind, source, _, b in LAYOUTS if (kind, source) == best)
    if not brands:
        return best
    # the same columns serve several brands: read the brand from the rows, then from the file name
    cols = [i for i, h in enumerate(header) if text(h).lower() in BRAND_COLUMNS]
    seen = " ".join(text(r[i][0]).upper() for r in rows[:300] for i in cols)
    hits = {b: sum(seen.count(w) for w in words) for b, words in brands.items()}
    if not any(hits.values()):
        hits = {b: sum(path.stem.upper().count(w) for w in words) for b, words in brands.items()}
    brand = max(hits, key=hits.get)
    return (best[0], brand if hits[brand] else "brand not found")


def split_brands(header, rows, brands, default):
    """One export can hold several brands (Geely and BMW come out of the same system): {brand: [row]}.

    A row that names no brand (the total row, for one) stays with the brand of the file.
    """
    cols = [i for i, h in enumerate(header) if text(h).lower() in BRAND_COLUMNS]
    out = {default: []}
    for row in rows:
        seen = " ".join(text(row[i][0]).upper() for i in cols)
        hits = {b: sum(seen.count(w) for w in words) for b, words in brands.items()}
        brand = max(hits, key=hits.get)
        out.setdefault(brand if hits[brand] else default, []).append(row)
    return {b: r for b, r in out.items() if r}


def clean(header, rows, key, sheet, fills=()):
    """For a part that replaces its sheet: drop the total row and rows that come twice, fill the columns in fills.

    Returns ({sheet: [row]}, removed, notes).
    """
    col = {name: i for i, name in enumerate(header)}
    for need in (key,) + (("Maker",) if sheet is None else ()):
        if need not in col:
            raise ValueError(f"column '{need}' is missing from the export")
    out = {name: [] for name in ([sheet] if sheet else dict.fromkeys(ORDER_SHEET.values()))}
    removed, seen, filled = [], set(), Counter()
    for n, row in enumerate(rows, start=1):
        if not any(isinstance(v, str) and v.strip() for v, _ in row):  # a total row holds numbers only
            removed.append((n, "total / blank row", row))
            continue
        cancel = next((c for c in CANCEL_COLS if c in col and text(row[col[c]][0])), None)
        if cancel:  # the finished report does not carry a cancelled order
            removed.append((n, f"cancelled order {text(row[col[key]][0])} ({cancel}: {text(row[col[cancel]][0])})", row))
            continue
        whole = tuple(text(v) for v, _ in row)
        if whole in seen:
            removed.append((n, "same row twice in the export files", row))
            continue
        seen.add(whole)
        for empty, source in fills:
            if empty in col and source in col and not text(row[col[empty]][0]) and text(row[col[source]][0]):
                row[col[empty]] = row[col[source]]
                filled[(empty, source)] += 1
        out[sheet or ORDER_SHEET[MAKER_SHEET.get(text(row[col["Maker"]][0]).upper(), OTHER_SHEET)]].append(row)
    return out, removed, [f"FILLED {n} rows: empty {a} takes the {b}" for (a, b), n in filled.items()]


def fix(header, rows, sheet=None):
    """Apply the fixes to a sales export. Returns ({sheet: [row]}, removed, changes, drop).

    sheet None = the KDC export, split over SHEETS by maker; otherwise every row goes to that sheet.
    drop holds the Frame# of the NOT_REPORTED cars, to take out of the workbook if they are in it.
    """
    col = {name: i for i, name in enumerate(header)}
    for need in (KEY, "Maker", "Model") if sheet is None else (KEY,):
        if need not in col:
            raise ValueError(f"column '{need}' is missing from the export")
    out = {name: [] for name in (SHEETS if sheet is None else [sheet])}
    removed, changes, seen, drop = [], [], set(), set()
    for n, row in enumerate(rows, start=1):
        get = lambda name: text(row[col[name]][0])
        frame = get(KEY)
        if not any(text(v) for v, _ in row):
            continue
        if not frame:
            removed.append((n, "no Frame# (total / blank row)", row))
            continue
        if frame in seen:
            removed.append((n, f"Frame# {frame} repeated in the export", row))
            continue
        seen.add(frame)
        if "CarType" in col and get("CarType").upper() in NOT_REPORTED:
            removed.append((n, f"{frame} {get('CarType')}: this kind of car is not reported", row))
            drop.add(frame)
            continue
        if sheet is not None:
            out[sheet].append(row)
            continue
        target = MAKER_SHEET.get(get("Maker").upper())
        if target is None:
            target = OTHER_SHEET
            changes.append((frame, "Sheet", f"maker '{get('Maker')}' is not in the list", OTHER_SHEET))
        new_model = MODEL_RENAME.get((get("Maker").upper(), get("Model").upper()))
        if new_model:
            changes.append((frame, "Model", get("Model"), new_model))
            row[col["Model"]] = (new_model, False)
        out[target].append(row)
    return out, removed, changes, drop


def sheet_heads(main_path):
    """The header row of every sheet the script writes to, as it stands in the main workbook: {sheet: [name, ...]}."""
    where = dict(SHEETS)
    for part, (_, sheet, hr, _) in TARGETS.items():
        for name in [sheet] if sheet else (ORDER_SHEET.values() if part[0] == "Order" else []):
            where[name] = hr
    wb = openpyxl.load_workbook(main_path, read_only=True, data_only=True)
    heads = {}
    for name, hr in where.items():
        if name not in wb.sheetnames:
            continue
        head = [text(v) for v in next(wb[name].iter_rows(min_row=hr, max_row=hr, max_col=MAX_COLS, values_only=True), ())]
        while head and not head[-1]:
            head.pop()
        if name in SHEETS:
            for j, title in BLANK_HEADERS.items():
                if j < len(head) and not head[j]:
                    head[j] = title
        heads[name] = head
    wb.close()
    return heads


def write_adjust(path, header, fixed, removed, changes, heads):
    """The fixed rows, one sheet per target sheet, laid out under that sheet's own header: the export is made to
    fit the main workbook, never the other way round."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    epoch = dt.datetime(1899, 12, 30)

    def put(ws, row, extra=()):
        ws.append(list(extra) + [epoch + dt.timedelta(days=v) if d and v is not None else v for v, d in row])
        for c, (v, d) in enumerate(row, start=1 + len(extra)):
            if d and v is not None:
                ws.cell(ws.max_row, c).number_format = "mm-dd-yy hh:mm" if v % 1 else "mm-dd-yy"

    for name, rows in fixed.items():
        ws = wb.create_sheet(name)
        head = heads.get(name)
        m = map_columns(header, head)[0] if head else None
        ws.append(head or header)
        for c in ws[1]:
            c.font = Font(bold=True)
        ws.freeze_panes = "A2"
        for row in rows:
            put(ws, row if m is None else [row[i] if i is not None else (None, False) for i in m])
    ws = wb.create_sheet("Removed")
    ws.append(["Export row", "Reason"] + header)
    for n, reason, row in removed:
        put(ws, row, (n, reason))
    ws = wb.create_sheet("Changes")
    ws.append([KEY, "Column", "Export value", "Changed to"])
    for ch in changes:
        ws.append(list(ch))
    wb.save(path)


def col_values(ws, col, r1, r2):
    """One column of the sheet as a list of text."""
    if r2 < r1:
        return []
    v = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Value2
    return [text(v)] if r1 == r2 else [text(x[0]) for x in v]


def map_columns(header, head):
    """Match the sheet's columns to the export's by header name (the order may differ).

    Returns (m, missing, extra, guessed): m[j] is the export column for sheet column j or None, missing are
    sheet headers the export does not have, extra are export headers the sheet does not have, guessed are
    (sheet header, export header) pairs matched by position only.
    A repeated header is matched in order. A sheet column left without a match takes the export column in
    the same position if nothing else uses it and, unless the sheet header is blank, the sheet has no column
    of that name (e.g. the sheet says 'CSdate' or a second 'Deposit' where the export says 'Sale date' / 'Tel1').
    """
    key = lambda s: text(s).lower()
    where, seen = {}, Counter()
    for i, h in enumerate(header):
        if key(h):
            where[(key(h), seen[key(h)])] = i
            seen[key(h)] += 1
    m, seen = [None] * len(head), Counter()
    for j, h in enumerate(head):
        if key(h):
            m[j] = where.get((key(h), seen[key(h)]))
            seen[key(h)] += 1
    used = {i for i in m if i is not None}
    named, guessed = {key(h) for h in head}, []
    for j, h in enumerate(head):
        if m[j] is None and j < len(header) and key(header[j]) and j not in used and (not key(h) or key(header[j]) not in named):
            m[j] = j
            used.add(j)
            if key(h):
                guessed.append((h, header[j]))
    missing = [head[j] for j in range(len(head)) if m[j] is None and key(head[j])]
    extra = [header[i] for i in range(len(header)) if i not in used and key(header[i])]
    return m, missing, extra, guessed


def sheet_head(ws, name, hr):
    """The sheet's header names, left to right, without the empty ones at the end."""
    head = [text(x) for x in ws.Range(ws.Cells(hr, 1), ws.Cells(hr, MAX_COLS)).Value2[0]]
    while head and not head[-1]:
        head.pop()
    if name in SHEETS:
        for j, title in BLANK_HEADERS.items():
            if j < len(head) and not head[j]:
                head[j] = title
    return head


def match_sheet(ws, name, hr, header, key, log):
    """Line the export up with the sheet. Returns (head, m); stops the run if the export is not this sheet's."""
    head = sheet_head(ws, name, hr)
    m, missing, extra, guessed = map_columns(header, head)
    if key not in head or m[head.index(key)] is None or len(missing) > len(head) / 2:
        raise ValueError(f"{name}: the export does not look like this sheet ({len(missing)} of {len(head)} columns not found)")
    if missing:
        log(f"NOTE {name}: columns the export does not have (left empty unless the sheet has a formula there): " + ", ".join(missing))
    if extra:
        log(f"NOTE {name}: export columns the sheet does not have (ignored): " + ", ".join(extra))
    if guessed:
        log(f"NOTE {name}: columns matched by position, the names differ (sheet <- export): " + ", ".join(f"{a} <- {b}" for a, b in guessed))
    return head, m


def col_runs(m):
    """Neighbouring sheet columns that come from the export, as (first, last) pairs: each run is written as one block."""
    runs, j = [], 0
    while j < len(m):
        if m[j] is None:
            j += 1
            continue
        k = j
        while k + 1 < len(m) and m[k + 1] is not None:
            k += 1
        runs.append((j, k))
        j = k + 1
    return runs


def write_rows(ws, r1, m, rows):
    """Write the export rows from row r1 down; dates landing in an unformatted column are shown as dates."""
    r2 = r1 + len(rows) - 1
    # text that Excel would turn into a date, time or number on entry ('2026-09-13', '16:46') goes in as text
    keep = lambda v: "'" + v if isinstance(v, str) and v and (v[0] in "=+-'" or any(ch.isdigit() for ch in v)) else v
    for a, b in col_runs(m):
        ws.Range(ws.Cells(r1, a + 1), ws.Cells(r2, b + 1)).Value2 = [[keep(row[m[j]][0]) for j in range(a, b + 1)] for row in rows]
    for j, i in enumerate(m):
        if i is not None and any(r[i][1] and r[i][0] is not None for r in rows):
            block = ws.Range(ws.Cells(r1, j + 1), ws.Cells(r2, j + 1))
            if block.NumberFormat == "General":
                block.NumberFormat = "mm-dd-yy"
    # Excel switches 'wrap text' on for a cell that receives a line break (long remarks) and the row grows tall:
    # switch it off again so the rows keep the height of the others
    tall = set()
    for n, row in enumerate(rows):
        for j, i in enumerate(m):
            if i is not None and isinstance(row[i][0], str) and "\n" in row[i][0]:
                ws.Cells(r1 + n, j + 1).WrapText = False
                tall.add(r1 + n)
    for r in tall:
        ws.Rows(r).AutoFit()


def fill_formulas(ws, first, r1, r2, width, old_last):
    """Columns the sheet works out itself (lookups etc.): copy the formula of the first data row down to rows r1-r2.

    old_last is the last data row the sheet had. A column that held its formula in only a few of those rows is not
    such a column but a small table standing beside the data (STOCK DL has one), and is left alone.
    """
    for c in range(1, width + HELPER_SCAN + 1):
        src = ws.Cells(first, c)
        if not src.HasFormula:
            continue
        if old_last > first + 3:
            had = ws.Range(ws.Cells(first, c), ws.Cells(old_last, c)).Formula
            if sum(1 for x in had if str(x[0]).startswith("=")) * 2 < len(had):
                continue
        have = ws.Range(ws.Cells(r1, c), ws.Cells(r2, c)).Formula
        have = [have] if r1 == r2 else [x[0] for x in have]
        runs = []  # empty cells next to each other are filled in one copy: a copy per cell takes minutes on a long sheet
        for i, f in enumerate(have):
            if f == "" and r1 + i != first:
                if runs and runs[-1][1] == r1 + i - 1:
                    runs[-1][1] = r1 + i
                else:
                    runs.append([r1 + i, r1 + i])
        for a, b in runs:
            src.Copy(ws.Range(ws.Cells(a, c), ws.Cells(b, c)))


def grow_filter(ws, r2):
    """Stretch the sheet's filter so it covers the rows down to r2."""
    if not ws.AutoFilterMode:
        return
    box = ws.AutoFilter.Range
    top, left, cols, bottom = box.Row, box.Column, box.Columns.Count, box.Row + box.Rows.Count - 1
    if bottom < r2:
        ws.AutoFilterMode = False
        ws.Range(ws.Cells(top, left), ws.Cells(r2, left + cols - 1)).AutoFilter()


def add_rows(xl, wb, sheets, header, fixed, drop, log):
    """Add the export rows whose Frame# is in none of these sheets yet. Returns {sheet: (ws, hr, last, count, m, mcol)}.

    drop: Frame# of cars the report does not carry; a row of the sheet holding one is deleted.
    """
    col = {name: i for i, name in enumerate(header)}
    amount = next((c for c in ("Selling AMT", "sales") if c in col), None)

    # frames already in the workbook, across all the sheets
    state, existing = {}, {}
    for name, hr in sheets.items():
        ws = wb.Worksheets(name)
        if ws.FilterMode:
            ws.ShowAllData()
        used_last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        head, m = match_sheet(ws, name, hr, header, KEY, log)
        mcol = {c: head.index(c) for c in (KEY,) + QC_COLS if c in head}  # sheet position of the columns used below
        frames = col_values(ws, mcol[KEY] + 1, hr + 1, used_last)
        for i in reversed(range(len(frames))):
            if frames[i] in drop:
                ws.Rows(hr + 1 + i).Delete()
                log(f"TAKEN OUT {name} row {hr + 1 + i}: {frames[i]} (this kind of car is not reported)")
                del frames[i]
        last = hr + max((i + 1 for i, f in enumerate(frames) if f), default=0)
        state[name] = (ws, hr, last, len([f for f in frames if f]), m, mcol)
        for i, f in enumerate(frames):
            if f:
                existing.setdefault(f, []).append((name, hr + 1 + i))
    for f, where in existing.items():
        if len(where) > 1:
            log(f"WARNING already in the main file {len(where)} times: {f} at " + ", ".join(f"{n} row {r}" for n, r in where))

    # the export covers the month so far, so every car in the main file should be in it
    sent = {text(r[col[KEY]][0]) for rows in fixed.values() for r in rows}
    gone = [(f, w[0]) for f, w in existing.items() if f not in sent]
    if gone:
        per = Counter(n for _, (n, _) in gone)
        log(f"ALERT {len(gone)} cars are in the main file but not in today's export (a download may be missing): "
            + ", ".join(f"{n} {c}" for n, c in per.items()) + "; e.g. " + ", ".join(f"{f} ({n} row {r})" for f, (n, r) in gone[:5]))

    for name, rows in fixed.items():
        ws, hr, last, count, m, mcol = state[name]
        width = len(m)
        new = [r for r in rows if text(r[col[KEY]][0]) not in existing]
        for row in rows:  # rows already there are left alone, differences are only reported
            frame = text(row[col[KEY]][0])
            for sheet, r in existing.get(frame, [])[:1]:
                if sheet != name:
                    log(f"WARNING {frame} belongs in {name} but is already in {sheet} row {r} - not added again")
                    continue
                old = ws.Range(ws.Cells(r, 1), ws.Cells(r, width)).Value2[0]
                d = [f"{c}: main '{text(old[mcol[c]])}' / export '{text(row[col[c]][0])}'" for c in QC_COLS
                     if c in mcol and c in col and text(old[mcol[c]]) != text(row[col[c]][0])]
                if d:
                    log(f"DIFFERENT (not changed) {name} row {r} {frame}: " + "; ".join(d))
        if new:
            first, r1, r2 = hr + 1, last + 1, last + len(new)
            blocks = [ws.Range(ws.Cells(r1, a + 1), ws.Cells(r2, b + 1)) for a, b in col_runs(m)]
            if any(xl.WorksheetFunction.CountA(b) for b in blocks):
                raise ValueError(f"{name}: rows {r1}-{r2} are not empty, nothing was written")
            if r1 > first:
                ws.Range(ws.Cells(first, 1), ws.Cells(first, width)).Copy()
                ws.Range(ws.Cells(r1, 1), ws.Cells(r2, width)).PasteSpecial(-4122)  # formats only: text stays text, dates show as dates
                xl.CutCopyMode = False
            write_rows(ws, r1, m, new)
            fill_formulas(ws, first, r1, r2, width, last)
            xl.CutCopyMode = False
            grow_filter(ws, r2)
            for i, row in enumerate(new):
                g = lambda c: text(row[col[c]][0]) if c in col else ""
                log(f"ADDED {name} row {r1 + i}: {g(KEY)} | {g('Showroom')} | {g('Maker')} {g('Model')} | {g('Status')} | {g(amount)}")
        log(f"{name}: {count} rows before, {len(new)} added, {len(rows) - len(new)} already there, {count + len(new)} now")
        state[name] = (ws, hr, last + len(new), count + len(new), m, mcol)
    return state


def replace_rows(xl, wb, name, hr, header, rows, key, log):
    """Put today's export in place of the sheet's rows. Formats and the sheet's own formula columns stay."""
    col = {h: i for i, h in enumerate(header)}
    ws = wb.Worksheets(name)
    if ws.FilterMode:
        ws.ShowAllData()
    head, m = match_sheet(ws, name, hr, header, key, log)
    width, first = len(m), hr + 1
    used_last = max(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, first)
    old = [k for k in col_values(ws, head.index(key) + 1, first, used_last) if k]
    for a, b in col_runs(m):
        ws.Range(ws.Cells(first, a + 1), ws.Cells(used_last, b + 1)).ClearContents()
    if rows:
        r2 = hr + len(rows)
        if r2 > first:
            wide = max(j for j, i in enumerate(m) if i is not None) + 1  # not past the export's columns: a side table may stand there
            ws.Range(ws.Cells(first, 1), ws.Cells(first, wide)).Copy()
            ws.Range(ws.Cells(first + 1, 1), ws.Cells(r2, wide)).PasteSpecial(-4122)  # formats only
            xl.CutCopyMode = False
        write_rows(ws, first, m, rows)
        fill_formulas(ws, first, first, r2, width, hr + len(old))
        xl.CutCopyMode = False
        grow_filter(ws, r2)
    now = [k for k in (text(r[col[key]][0]) for r in rows) if k]
    log(f"{name}: replaced with today's export - {len(old)} rows before, {len(rows)} now "
        f"({len(set(now) - set(old))} new, {len(set(old) - set(now))} no longer in the export, counted by {key})")


def apply(main_path, out_path, jobs, log, report_day=None):
    """jobs: ("add", {sheet: header row}, header, {sheet: rows}) or ("replace", sheet, header row, header, rows, key).

    report_day (an Excel day number) is for test runs: the report date cell gets that day instead of TODAY().
    """
    import pythoncom
    import win32com.client

    shutil.copy2(main_path, out_path)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    wb = None
    try:
        wb = xl.Workbooks.Open(str(out_path), 0, False)
        if report_day:
            wb.Worksheets("Daily Report").Range("G3").Value2 = report_day
        xl.CalculateFull()
        before = {ref: wb.Worksheets(s).Range(ref).Value2 for s, ref, _ in QC_CELLS}

        # Excel recalculates the whole workbook after every write: switched off while the rows go in, and back on
        # before the workbook is calculated and saved (the setting is saved with the workbook)
        calc = xl.Calculation
        xl.Calculation = -4135  # manual
        state = {}
        try:
            for job in jobs:
                if job[0] == "add":
                    state.update(add_rows(xl, wb, *job[1:], log))
                else:
                    replace_rows(xl, wb, *job[1:], log)
        finally:
            xl.Calculation = calc

        xl.CalculateFull()
        log("")
        log("Report numbers (before -> after):")
        for s, ref, label in QC_CELLS:
            log(f"  {label:28s} {s}!{ref}: {text(before[ref])} -> {text(wb.Worksheets(s).Range(ref).Value2)}")
        total, by_model = (wb.Worksheets(s).Range(ref).Value2 for s, ref, _ in QC_CELLS[-2:])
        if total != by_model:
            log(f"CHECK the two month totals differ: Daily Report {text(total)} vs by-model sheet {text(by_model)}")
        g3 = wb.Worksheets("Daily Report").Range("G3")
        log(f"  Report date (Daily Report!G3): {(dt.datetime(1899, 12, 30) + dt.timedelta(days=float(g3.Value2))):%d-%b-%Y}"
            + ("  <- formula TODAY(), so the 'TODAY' columns follow the day the file is opened" if "TODAY" in str(g3.Formula).upper() else ""))

        # rows the report formulas will not pick up
        dr, bm = wb.Worksheets("Daily Report"), wb.Worksheets("By Model.DATE - with MOS")
        sites = {v.upper() for c in (2, 3, 4) for v in col_values(dr, c, 7, 70) if v}
        models = {v.upper() for c in (3, 90) for v in col_values(bm, c, 9, 101) if v}
        for name in ("KDC-CUS", "KDC X-CUS"):
            if name not in state:
                continue
            ws, hr, last, _, _, mcol = state[name]
            if not all(c in mcol for c in ("Showroom", "Cust#", "Model")):
                continue
            show, cust, model, frames = (col_values(ws, mcol[c] + 1, hr + 1, last) for c in ("Showroom", "Cust#", "Model", KEY))
            for i, f in enumerate(frames):
                if f and show[i].upper() not in sites and cust[i].upper() not in sites:
                    log(f"CHECK {name} row {hr + 1 + i} {f}: showroom '{show[i]}' / customer '{cust[i]}' has no row on Daily Report")
                if f and model[i].upper() not in models:
                    log(f"CHECK {name} row {hr + 1 + i} {f}: model '{model[i]}' has no row on the by-model sheet")
        for sheet in REFILTER:
            ws = wb.Worksheets(sheet)
            if ws.AutoFilterMode:
                ws.AutoFilter.ApplyFilter()
        wb.Save()

        # report pictures; the chart used to export them is never saved into the workbook
        for sheet, ref, title, *hide in PICTURES:
            ws = wb.Worksheets(sheet)
            for h in hide:  # only for the picture: the workbook was saved above and is closed without saving
                ws.Range(h).EntireColumn.Hidden = True
            rng = ws.Range(ref)
            png = out_path.with_name(f"{title} {out_path.stem.rsplit(' ', 1)[-1]}.png")
            for attempt in range(5):  # the clipboard is sometimes busy
                try:
                    rng.CopyPicture(1, 2)  # as shown on screen, bitmap
                    box = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
                    box.Activate()
                    box.Chart.ChartArea.Format.Line.Visible = False
                    box.Chart.Paste()
                    box.Chart.Export(str(png))
                    box.Delete()
                    log(f"Picture: {png}")
                    break
                except pythoncom.com_error:
                    time.sleep(1)
            else:
                log(f"WARNING could not export the picture of {sheet}!{ref}")
            for h in hide:
                ws.Range(h).EntireColumn.Hidden = False
    finally:
        if wb is not None:
            wb.Close(False)
        xl.Quit()
        pythoncom.CoUninitialize()


# ---- Thai copy of the QC report -----------------------------------------------------------------
# The run logs in English; the Thai file is made from those lines, so the two always say the same thing.
# (pattern, replacement), all applied to every line in this order. The word a line starts with (OK, MISSING,
# ALERT ...) stays in English, as do sheet, column and file names. A line no pattern knows is copied as it is.
TH_LABELS = {"KDC units this month": "KDC จำนวนคันเดือนนี้", "KDC amount this month": "KDC ยอดเงินเดือนนี้",
             "CHANGAN units this month": "CHANGAN จำนวนคันเดือนนี้", "New car sales total": "ยอดขายรถใหม่รวม",
             "By-model month total": "ยอดรวมเดือนตามรุ่น"}
TH = [(re.compile(p), r) for p, r in [
    (r"^Run (\S+)$", r"รอบ \1"),
    (r"^Main workbook: REPLACED with the output\. The old one is at ", "ไฟล์หลัก: แทนที่ด้วยไฟล์ผลลัพธ์แล้ว ไฟล์เดิมอยู่ที่ "),
    (r"^Main workbook: kept as it was\.$", "ไฟล์หลัก: คงไว้เหมือนเดิม"),
    (r"^Main workbook: NOT replaced \(a file was open in Excel\)\.$", "ไฟล์หลัก: ไม่ได้แทนที่ (มีไฟล์เปิดอยู่ใน Excel)"),
    (r"^Main workbook: ", "ไฟล์หลัก: "),
    (r"^TEST RUN as of (\S+): later rows are left out, the report date is set to that day, Main and Input are left alone$",
     r"TEST RUN ณ วันที่ \1: แถวที่วันที่หลังจากนั้นถูกตัดออก วันที่รายงานตั้งเป็นวันนั้น ไม่แตะ Main และ Input"),
    # checklist
    (r"^CHECKLIST of today's files:$", "CHECKLIST ไฟล์ของวันนี้:"),
    (r"  <- more than one file for this part", "  <- มีมากกว่าหนึ่งไฟล์สำหรับส่วนนี้"),
    (r"  \(received, but this part is not built yet\)", "  (ได้รับแล้ว แต่โปรแกรมยังไม่ทำส่วนนี้)"),
    (r"\(brand no longer sold\)", "(ยี่ห้อที่เลิกขายแล้ว)"),
    (r": could not be read \((.*)\)$", r": อ่านไฟล์ไม่ได้ (\1)"),
    (r": looks like a (\w+) export of Geely / BMW / MBS but the brand could not be read - put GEELY, BMW or MBS in the file name$",
     r": ดูเหมือน export ประเภท \1 ของ Geely / BMW / MBS แต่อ่านยี่ห้อไม่ออก - ให้ใส่ GEELY, BMW หรือ MBS ในชื่อไฟล์"),
    (r": its columns match no part of the report$", ": หัวคอลัมน์ไม่ตรงกับส่วนใดของรายงาน"),
    (r"^ALERT: (\d+) of (\d+) parts are missing: ", r"ALERT: ขาด \1 จาก \2 ส่วน: "),
    (r"^All (\d+) parts were received\.$", r"ได้รับครบทั้ง \1 ส่วน"),
    (r"^STOP: Input holds no export this script applies, so nothing was written\.$",
     "STOP: ใน Input ไม่มี export ที่โปรแกรมนี้ใช้ จึงไม่ได้เขียนอะไร"),
    # each part
    (r": (\d+) rows read, (\d+) kept, (\d+) removed$", r": อ่าน \1 แถว, เก็บ \2, ตัดออก \3"),
    (r"^REMOVED export row (\d+): ", r"REMOVED แถวที่ \1 ของ export: "),
    (r"no Frame# \(total / blank row\)", "ไม่มี Frame# (แถวรวม / แถวว่าง)"),
    (r"total / blank row", "แถวรวม / แถวว่าง"),
    (r"cancelled order", "order ที่ยกเลิก"),
    (r"same row twice in the export files", "แถวซ้ำกันในไฟล์ export"),
    (r"Frame# (\S+) repeated in the export", r"Frame# \1 ซ้ำใน export"),
    (r"this kind of car is not reported", "รถประเภทนี้ไม่อยู่ในรายงาน"),
    (r"maker '(.*?)' is not in the list", r"ยี่ห้อ '\1' ไม่อยู่ในรายการ"),
    (r"^FILLED (\d+) rows: empty (.+) takes the (.+)$", r"FILLED \1 แถว: \2 ที่ว่างใช้ค่าจาก \3"),
    (r"^Split: ", "แยกลง sheet: "),
    (r"^Brands in the sales export: ", "ยี่ห้อใน export ยอดขาย: "),
    (r"^NOTE no sales rows this month for: (.*?)  \(fine if they sold nothing; if they did, a download is missing\)$",
     r"NOTE เดือนนี้ไม่มีแถวยอดขายของ: \1  (ไม่เป็นไรถ้าไม่มียอดขายจริง ถ้ามียอดขาย แปลว่าไฟล์ดาวน์โหลดขาด)"),
    (r"^ALERT the main workbook is for (.+?) but the export has sales dated (.*) - check that Main holds the right month's workbook$",
     r"ALERT ไฟล์หลักเป็นของเดือน \1 แต่ export มียอดขายลงวันที่ \2 - ตรวจว่าไฟล์ใน Main เป็นของเดือนที่ถูกต้อง"),
    (r"^Adjusted file: ", "ไฟล์ที่แก้แล้ว: "),
    (r": LEFT OUT (\d+) rows dated after the test day$", r": LEFT OUT \1 แถวที่วันที่หลังวันทดสอบ"),
    # the workbook
    (r": columns the export does not have \(left empty unless the sheet has a formula there\): ",
     ": คอลัมน์ที่ export ไม่มี (ปล่อยว่าง ยกเว้น sheet มีสูตรอยู่ตรงนั้น): "),
    (r": export columns the sheet does not have \(ignored\): ", ": คอลัมน์ของ export ที่ sheet ไม่มี (ไม่ใช้): "),
    (r": columns matched by position, the names differ \(sheet <- export\): ",
     ": คอลัมน์ที่จับคู่ตามตำแหน่ง ชื่อไม่ตรงกัน (sheet <- export): "),
    (r"^WARNING already in the main file (\d+) times: (\S+) at ", r"WARNING มีอยู่ในไฟล์หลัก \1 ครั้ง: \2 ที่ "),
    (r"^ALERT (\d+) cars are in the main file but not in today's export \(a download may be missing\): (.*); e\.g\. ",
     r"ALERT มีรถ \1 คันในไฟล์หลักที่ไม่อยู่ใน export ของวันนี้ (อาจขาดไฟล์ดาวน์โหลด): \2; เช่น "),
    (r"^WARNING (\S+) belongs in (.+) but is already in (.+) row (\d+) - not added again$",
     r"WARNING \1 ควรอยู่ใน \2 แต่มีอยู่แล้วใน \3 แถว \4 - ไม่เพิ่มซ้ำ"),
    (r"^DIFFERENT \(not changed\) (.*)$", lambda m: "DIFFERENT (ไม่ได้แก้) " + m.group(1).replace(": main '", ": ไฟล์หลัก '")),
    (r"^(.+): (\d+) rows before, (\d+) added, (\d+) already there, (\d+) now$",
     r"\1: เดิม \2 แถว, เพิ่ม \3, มีอยู่แล้ว \4, ตอนนี้ \5"),
    (r"^(.+): replaced with today's export - (\d+) rows before, (\d+) now \((\d+) new, (\d+) no longer in the export, counted by (.+)\)$",
     r"\1: แทนที่ด้วย export ของวันนี้ - เดิม \2 แถว, ตอนนี้ \3 (ใหม่ \4, ไม่อยู่ใน export แล้ว \5, นับตาม \6)"),
    # report numbers, checks, pictures
    (r"^Report numbers \(before -> after\):$", "ตัวเลขรายงาน (ก่อน -> หลัง):"),
    (r"^  Report date \(Daily Report!G3\): ", "  วันที่รายงาน (Daily Report!G3): "),
    (r"  <- formula TODAY\(\), so the 'TODAY' columns follow the day the file is opened$",
     "  <- เป็นสูตร TODAY() คอลัมน์ 'TODAY' จึงเปลี่ยนตามวันที่เปิดไฟล์"),
    (r"^  (.{28}) (.+![A-Z]+\d+: .*)$", lambda m: f"  {TH_LABELS.get(m.group(1).strip(), m.group(1).strip())}  {m.group(2)}"),
    (r"^CHECK the two month totals differ: Daily Report (.+) vs by-model sheet (.+)$",
     r"CHECK ยอดรวมเดือนสองที่ไม่เท่ากัน: Daily Report \1 กับ sheet ตามรุ่น \2"),
    (r": showroom '(.*)' / customer '(.*)' has no row on Daily Report$", r": showroom '\1' / ลูกค้า '\2' ไม่มีแถวใน Daily Report"),
    (r": model '(.*)' has no row on the by-model sheet$", r": รุ่น '\1' ไม่มีแถวใน sheet ตามรุ่น"),
    (r"^Picture: ", "รูป: "),
    (r"^WARNING could not export the picture of ", "WARNING ถ่ายรูปไม่สำเร็จ: "),
    (r"^Output for QC: ", "ไฟล์ผลลัพธ์สำหรับ QC: "),
    (r"^Input: files kept\.$", "Input: เก็บไฟล์ไว้"),
    (r"^Input: (\d+) file\(s\) sent to the Recycle Bin: ", r"Input: ส่ง \1 ไฟล์ไป Recycle Bin: "),
    (r"^Input: the files could NOT be deleted \(one may be open in Excel\)\.$", "Input: ลบไฟล์ไม่ได้ (อาจมีไฟล์เปิดอยู่ใน Excel)"),
    # words left over inside the lines above
    (r"\((\d+) rows\)", r"(\1 แถว)"),
    (r" row (\d+)\b", r" แถว \1"),
]]


def thai(line):
    """One line of the QC report in Thai."""
    for pattern, repl in TH:
        line = pattern.sub(repl, line)
    return line


def ask(question, buttons=4):
    """A Windows message box on top of other windows. buttons: 4 = Yes/No (No is the default), 5 = Retry/Cancel."""
    import ctypes
    flags = buttons | 0x20 | 0x40000 | 0x10000 | (0x100 if buttons == 4 else 0)
    return ctypes.windll.user32.MessageBoxW(0, question, "Daily sales report", flags) in (6, 4)  # Yes or Retry


def replace_main(main_path, out_path, stamp, log, choice):
    """Offer to make the finished workbook the new main file. The old main is kept in Main\\previous."""
    if choice is None:
        choice = ask("The report is ready in the output folder.\n\nCheck it first if you want (this box can wait).\n\n"
                     "Replace the workbook in Main with this output?\n\n"
                     "Yes = Main is updated, the old one is kept in Main\\previous\nNo = Main stays as it is")
    if not choice:
        log("Main workbook: kept as it was.")
        return
    keep = MAIN / "previous"
    keep.mkdir(exist_ok=True)
    backup = keep / f"{main_path.stem} (before {stamp}){main_path.suffix}"
    incoming = keep / "incoming.tmp"
    while True:
        try:
            shutil.copy2(out_path, incoming)   # copy first, so Main is never left empty
            shutil.move(main_path, backup)
            shutil.move(incoming, main_path)
            log(f"Main workbook: REPLACED with the output. The old one is at {backup}")
            return
        except PermissionError:
            incoming.unlink(missing_ok=True)
            if not ask("The main workbook or the output is open in Excel.\n\nClose it, then choose Retry.", 5):
                log("Main workbook: NOT replaced (a file was open in Excel).")
                return


def recycle(paths):
    """Send files to the Windows Recycle Bin (they can be restored from there). True if it worked."""
    import ctypes
    from ctypes import wintypes

    class Op(ctypes.Structure):
        _fields_ = [("hwnd", wintypes.HWND), ("wFunc", wintypes.UINT), ("pFrom", ctypes.c_void_p), ("pTo", ctypes.c_void_p),
                    ("fFlags", ctypes.c_ushort), ("aborted", wintypes.BOOL), ("maps", ctypes.c_void_p), ("title", ctypes.c_void_p)]

    names = ctypes.create_unicode_buffer("\0".join(str(p) for p in paths) + "\0\0")
    op = Op(None, 3, ctypes.cast(names, ctypes.c_void_p), None, 0x40 | 0x10 | 0x4 | 0x400)  # delete, allow undo, no prompts
    return ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op)) == 0 and not op.aborted


def clear_input(files, log, choice):
    """Offer to remove the export files that were used, so Input is empty for the next day."""
    if not files:
        return
    if choice is None:
        choice = ask(f"Delete the {len(files)} export file(s) that were used from the Input folder?\n\n"
                     + "\n".join(f.name for f in files[:12]) + ("\n..." if len(files) > 12 else "")
                     + "\n\nYes = they go to the Recycle Bin\nNo = they stay in Input")
    if not choice:
        log("Input: files kept.")
    elif recycle(files):
        log(f"Input: {len(files)} file(s) sent to the Recycle Bin: " + ", ".join(f.name for f in files))
    else:
        log("Input: the files could NOT be deleted (one may be open in Excel).")


def main():
    # optional: these answer the questions at the end without showing the boxes
    choice = True if "--replace-main" in sys.argv else False if "--keep-main" in sys.argv else None
    clear = True if "--delete-input" in sys.argv else False if "--keep-input" in sys.argv else None
    # test: --until=2026-10-05 builds the report as of that day (later rows left out, report date set to it).
    # A test never replaces Main or empties Input.
    until = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--until=")), None)
    if until:
        until = (dt.datetime.strptime(until, "%Y-%m-%d") - dt.datetime(1899, 12, 30)).days
        choice = clear = False
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M")
    for d in (ADJUST, OUTPUT):
        d.mkdir(exist_ok=True)
    mains = [p for p in MAIN.glob("*.xls*") if not p.name.startswith("~$")]
    raws = [p for p in INPUT.iterdir() if p.suffix.lower() in (".xls", ".xlsx", ".xlsm") and not p.name.startswith("~$")]
    if len(mains) != 1:
        sys.exit(f"STOP: Main must hold exactly one workbook (found {len(mains)}).")

    lines = []

    def log(msg):
        print(msg)
        lines.append(msg)

    day_dir = out_path = None
    try:
        log(f"Run {stamp}")
        log(f"Main workbook: {mains[0].name}")
        if until:
            log(f"TEST RUN as of {dt.datetime(1899, 12, 30) + dt.timedelta(days=until):%d-%b-%Y}: later rows are left out, "
                "the report date is set to that day, Main and Input are left alone")

        # ---- which exports arrived, and which are missing
        sig = main_signatures(mains[0])
        expected = [(kind, b) for kind, source, _, brands in LAYOUTS for b in (brands or [source]) if b not in DROPPED]
        got, unknown, left_out = {}, [], []
        for raw in sorted(raws):
            try:
                h, rows = read_raw(raw)
                part = classify(raw, h, rows, sig)
            except Exception as e:
                unknown.append(f"{raw.name}: could not be read ({e})")
                continue
            if part in expected or (part and part[1] in DROPPED):
                brands = next(b for kind, source, _, b in LAYOUTS if kind == part[0] and part[1] in (b or [source]))
                for src, part_rows in (split_brands(h, rows, brands, part[1]) if brands else {part[1]: rows}).items():
                    if src in DROPPED:
                        left_out.append(f"{part[0]:5s} - {src:7s} {raw.name} ({len(part_rows)} rows)  (brand no longer sold)")
                    else:
                        got.setdefault((part[0], src), []).append((raw, h, part_rows))
            elif part:
                unknown.append(f"{raw.name}: looks like a {part[0].lower()} export of Geely / BMW / MBS but the brand could not be read "
                               "- put GEELY, BMW or MBS in the file name")
            else:
                unknown.append(f"{raw.name}: its columns match no part of the report")
        for kind, brand in expected:  # the file such a brand comes in arrived without rows of it: it has none today
            host = (kind, INSIDE.get(brand))
            if (kind, brand) not in got and host in got:
                got[(kind, brand)] = [(raw, h, []) for raw, h, _ in got[host][:1]]
        expected = [p for p in expected if p in got or p not in NOT_EXPECTED]
        missing = [p for p in expected if p not in got]
        log("")
        log("CHECKLIST of today's files:")
        for p in expected:
            label = f"{p[0]:5s} - {p[1]:7s}"
            if p not in got:
                log(f"  MISSING      {label}")
                continue
            files = ", ".join(f"{f.name} ({len(rows)} rows)" for f, _, rows in got[p])
            more = "  <- more than one file for this part" if len(got[p]) > 1 else ""
            log(f"  {'OK' if p in TARGETS else 'NOT APPLIED':12s} {label} {files}{more}" + ("" if p in TARGETS else "  (received, but this part is not built yet)"))
        for u in left_out:
            log(f"  LEFT OUT     {u}")
        for u in unknown:
            log(f"  UNKNOWN      {u}")
        if missing:
            log(f"ALERT: {len(missing)} of {len(expected)} parts are missing: " + ", ".join(f"{k} {s}" for k, s in missing))
        else:
            log(f"All {len(expected)} parts were received.")
        log("")

        todo = [p for p in expected if p in got and p in TARGETS]
        if not todo:
            log("STOP: Input holds no export this script applies, so nothing was written.")
            return

        # one folder per day, and inside it one numbered folder per run (1, 2, 3 ...): workbook, pictures, QC report
        day = OUTPUT / dt.datetime.now().strftime("%Y-%m-%d")
        day.mkdir(exist_ok=True)
        day_dir = day / str(max((int(p.name) for p in day.iterdir() if p.is_dir() and p.name.isdigit()), default=0) + 1)
        day_dir.mkdir()
        out_path = day_dir / f"{mains[0].stem} - applied {stamp}{mains[0].suffix}"

        # ---- each part: fix, write to adjust, queue for the workbook
        jobs, heads = [], sheet_heads(mains[0])
        for part in todo:
            kind, src = part
            mode, sheet, hr, key = TARGETS[part]
            header, rows = got[part][0][1], []
            for raw, h, part_rows in got[part]:
                if h != header:  # same part, columns in another order: line them up with the first file
                    at = {name: i for i, name in enumerate(h)}
                    part_rows = [[row[at[name]] if name in at else (None, False) for name in header] for row in part_rows]
                rows += part_rows
            if until:
                at = [i for i, h in enumerate(header) if text(h).lower() in UNTIL_COLS[kind]]
                # a date years ahead is a typing slip in the system (Buddhist year 2569), not a later day
                kept = [r for r in rows if not any(r[i][1] and r[i][0] is not None and until + 1 <= r[i][0] < until + 366 for i in at)]
                if len(kept) != len(rows):
                    log(f"{kind} {src}: LEFT OUT {len(rows) - len(kept)} rows dated after the test day")
                rows = kept
            if not rows and mode == "add":
                continue
            notes, changes, drop = [], [], set()
            if mode == "add":
                fixed, removed, changes, drop = fix(header, rows, sheet)
            else:
                fixed, removed, notes = clean(header, rows, key, sheet, FILL_FROM.get(part, ()))
            log(f"{kind} {src} - " + ", ".join(raw.name for raw, _, _ in got[part])
                + f": {len(rows)} rows read, {sum(len(v) for v in fixed.values())} kept, {len(removed)} removed")
            if len(removed) <= 10:
                for n, reason, _ in removed:
                    log(f"REMOVED export row {n}: {reason}")
            else:
                log("REMOVED " + ", ".join(f"{c} x {reason}" for reason, c in Counter(r for _, r, _ in removed).items()))
            for frame, c, old, new in changes:
                log(f"CHANGED {frame}: {c} '{old}' -> '{new}'")
            for note in notes:
                log(note)
            if len(fixed) > 1:
                log("Split: " + ", ".join(f"{name} {len(v)}" for name, v in fixed.items()))
            if part == ("Sale", "KDC"):
                at = header.index("Maker")
                makers = Counter(text(r[at][0]).upper() for v in fixed.values() for r in v)
                log("Brands in the sales export: " + ", ".join(f"{m} {n}" for m, n in makers.most_common()))
                absent = [m for m in MAKER_SHEET if m not in makers]
                if absent:
                    log("NOTE no sales rows this month for: " + ", ".join(absent) + "  (fine if they sold nothing; if they did, a download is missing)")
                month = report_month(mains[0])
                wrong = other_months(header, fixed, month)
                if wrong:
                    log(f"ALERT the main workbook is for {dt.date(month[0], month[1], 1):%b %Y} but the export has sales dated "
                        + ", ".join(f"{m} ({n} rows)" for m, n in wrong.items()) + " - check that Main holds the right month's workbook")
            if rows:
                adjust_path = ADJUST / f"{'Sales' if part == ('Sale', 'KDC') else f'{kind} {src}'} adjusted {stamp}.xlsx"
                write_adjust(adjust_path, header, fixed, removed, changes, heads)
                log(f"Adjusted file: {adjust_path}")
            log("")
            if mode == "add":
                jobs.append(("add", {sheet: hr} if sheet else SHEETS, header, fixed, drop))
            else:
                jobs += [("replace", name, hr, header, v, key) for name, v in fixed.items()]

        apply(mains[0], out_path, jobs, log, until)
        log("")
        log(f"Output for QC: {out_path}")
        replace_main(mains[0], out_path, stamp, log, choice)
        clear_input(list(dict.fromkeys(f for files in got.values() for f, _, _ in files)), log, clear)  # files it did not recognise stay
    except Exception as e:
        log(f"ERROR: {e}")
        if out_path is not None:
            out_path.unlink(missing_ok=True)
        raise
    finally:
        if day_dir is not None:
            (day_dir / f"QC {stamp}.txt").write_text("\n".join(lines), encoding="utf-8")
            (day_dir / f"QC {stamp} (TH).txt").write_text("\n".join(thai(x) for x in lines), encoding="utf-8-sig")


if __name__ == "__main__":
    main()
