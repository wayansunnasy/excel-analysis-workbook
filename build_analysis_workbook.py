"""Point this at any CSV and get an Excel workbook you can paste new data into.

The problem it solves: most "dashboard" workbooks address a fixed block of cells. Paste 700 rows
into a sheet built for 500 and the summary silently ignores the last 200. Paste 300 and it counts
empty rows as real ones.

Here nothing addresses a fixed range. The rows go into a real Excel Table named SourceData, and
every tile and helper formula refers to it by name, for example SUM(SourceData[amount]). An Excel
Table grows and shrinks with whatever you paste, so the formulas follow the data.

No macros and no VBA, deliberately. Macros are the usual reason a workbook behaves differently on
macOS, and they make the file raise a security prompt every time someone opens it.

Column types are detected from the values rather than the header, with one deliberate exception:
headers that name a label rather than a measurement (id, phone, anything ending _no or _code) stay
as text. Storing a phone number as a number eats the leading plus, and summing an id means nothing.

Dates written 03/04/2026 are ambiguous. Each date column is read as a whole: if any value in it can
only be day first (a day above 12) the column is day first, if any can only be month first it is
month first, and if the column contradicts itself it is left as text and reported. A column that
never settles the question is read day first unless you pass --month-first.

Tabs produced:
  Analysis   headline tiles, one breakdown block per usable text column, native Excel charts
  Data       the rows as the SourceData table, which is where new data is pasted
  Read me    where to paste, how it refreshes, and how to add a pivot table and a slicer

Usage:
  python build_analysis_workbook.py <input.csv> [output.xlsx] [--month-first]

Requires openpyxl. Tests: python tests/test_build_analysis_workbook.py
"""
import csv
import pathlib
import re
import sys
from collections import Counter
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

NAVY, LIGHT, GREY = "1F4E78", "DDEBF7", "F2F2F2"
bold = Font(bold=True)
white_bold = Font(bold=True, color="FFFFFF")
title_font = Font(bold=True, size=14, color=NAVY)
navy_fill = PatternFill("solid", fgColor=NAVY)
light_fill = PatternFill("solid", fgColor=LIGHT)
grey_fill = PatternFill("solid", fgColor=GREY)
thin = Side(style="thin", color="BFBFBF")
box = Border(left=thin, right=thin, top=thin, bottom=thin)

# Four digit years only. A compact 20260403 or a two digit 03/04/26 is not worth the false
# positives: a postcode 012345 and a version 1.2.30 both parse as dates under those patterns.
YEAR_FIRST_PATTERNS = ("%Y-%m-%d", "%Y/%m/%d")
DAY_FIRST_PATTERNS = ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y")
MONTH_FIRST_PATTERNS = ("%m/%d/%Y", "%m-%d-%Y", "%m.%d.%Y")
# Characters Excel requires escaping with a single quote inside a structured reference such as
# SourceData[amount ($)]. Taken from Microsoft's list for table column specifiers.
STRUCTURED_REF_SPECIALS = re.compile(r"([\t\n\r,:.\[\]#'\"{}$^&*+=\-<>/])")
# Columns that look numeric but are labels. Storing these as numbers loses the leading plus of a
# phone number and the leading zero of a postcode, and summing them means nothing.
# Matched against the whole header, not against any word inside it: "last_order_amount" is money,
# while "order_no" is a label.
LABEL_HEADERS = re.compile(
    r"^(id|ids|code|zip|postal|postcode|post_code|phone|telephone|tel|mobile|fax|ref|reference|"
    r"sku|upc|ean|isbn|iban|bic|pin|vat|siret|account|acct)$"
    r"|_(id|no|num|number|code|ref)$"
    r"|^(phone|mobile|tel|fax|zip|postal|account|invoice|order)_", re.I)
MAX_BREAKDOWNS = 3          # breakdown blocks on the Analysis tab
MAX_CATEGORIES = 12         # rows per breakdown block, plus one Other row


def as_date(value, order="dmy"):
    """Parse one cell as a date. `order` settles 03/04/2026: "dmy" is 3 April, "mdy" is 4 March."""
    try:
        text = value.strip()
    except AttributeError:
        return None
    if not text:
        return None
    try:                                        # ISO, with or without a time part and a Z
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    ordered = DAY_FIRST_PATTERNS if order == "dmy" else MONTH_FIRST_PATTERNS
    for pattern in YEAR_FIRST_PATTERNS + ordered:
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def date_order(values, default="dmy"):
    """Decide, for a whole column, whether 03/04/2026 means 3 April or 4 March.

    Returns "dmy", "mdy", or "mixed" when the column contains values that can only be one and
    values that can only be the other. A mixed column is not a date column: reading it either way
    would silently put some rows in the wrong month.
    """
    only_dmy = only_mdy = 0
    for v in values:
        d, m = as_date(v, "dmy"), as_date(v, "mdy")
        if d and not m:
            only_dmy += 1
        elif m and not d:
            only_mdy += 1
    if only_dmy and only_mdy:
        return "mixed"
    if only_dmy:
        return "dmy"
    if only_mdy:
        return "mdy"
    return default


def as_number(value):
    """Read a number the way a person would: 1,234.50 and 1 000 and $99 and (120.00) and 12 USD.

    A percent sign is left alone on purpose. Turning "12%" into 12 changes its meaning, so such a
    column stays text rather than being quietly wrong by a factor of a hundred.
    """
    if value is None:
        return None
    text = str(value).strip()
    if not text or "%" in text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    # A currency symbol, or a code of up to three letters set off by a space, on either side.
    # Letters glued to the digits are left alone: SKU123 and A100 are codes, not 123 and 100.
    text = re.sub(r"^(-?)(?:[A-Za-z]{1,3}\s+|[^\w\s.\-]+\s*)", r"\1", text)
    text = re.sub(r"(?:\s+[A-Za-z]{1,3}|\s*[^\w\s.\-]+)$", "", text)
    text = re.sub(r"[\s,]", "", text)             # thousands separators, either style
    if text in ("", "-", "."):
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return -number if negative else number


def looks_like_a_label(value):
    """A leading plus or a leading zero that must survive: +230..., 01234. Not 0.85, and not a date,
    which is checked before this ever runs."""
    text = str(value).strip()
    return len(text) > 5 and re.match(r"^(\+|0\d)", text) is not None


def sniff(rows, header, default_order="dmy"):
    """Decide what each column is, from the values rather than from the header name.

    Order matters. Dates are tested before the leading-zero rule, because 03/04/2026 starts with a
    zero and is still a date. Labels are tested before numbers, because 01234 parses as a number and
    is still a postcode.
    """
    kinds = []
    for i, name in enumerate(header):
        values = [r[i] for r in rows if i < len(r) and str(r[i]).strip() != ""]
        if not values:
            kinds.append("empty")
            continue
        if LABEL_HEADERS.search(str(name)):
            kinds.append("text")
            continue
        order = date_order(values, default_order)
        dates = 0 if order == "mixed" else sum(1 for v in values if as_date(v, order) is not None)
        if dates >= 0.8 * len(values):
            kinds.append("date")
            continue
        if sum(1 for v in values if looks_like_a_label(v)) >= 0.3 * len(values):
            kinds.append("text")
            continue
        numbers = sum(1 for v in values if as_number(v) is not None)
        kinds.append("number" if numbers >= 0.8 * len(values) else "text")
    return kinds


def clean_headers(header):
    """Headers become Excel Table column names, which must be unique and are better without line
    breaks. Returns the cleaned list and a note for every header that had to change."""
    out, notes, seen = [], [], set()
    for i, raw in enumerate(header, 1):
        name = re.sub(r"\s+", " ", str(raw)).strip() or f"column_{i}"
        base, n = name, 2
        while name.lower() in seen:
            name = f"{base}_{n}"
            n += 1
        if name != str(raw):
            notes.append(f"header {i} {raw!r} written as {name!r}")
        seen.add(name.lower())
        out.append(name)
    return out, notes


def col(name):
    """A column reference inside the SourceData table, escaped the way Excel requires."""
    return "SourceData[" + STRUCTURED_REF_SPECIALS.sub(r"'\1", name) + "]"


def build(csv_path, out_path, month_first=False):
    with open(csv_path, newline="", encoding="utf-8-sig", errors="replace") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        rows = [r for r in reader if any(str(c).strip() for c in r)]
    if not rows:
        raise SystemExit(f"{csv_path} has a header but no data rows")

    header, notes = clean_headers(header)
    width = len(header)
    rows = [(r + [""] * width)[:width] for r in rows]
    default_order = "mdy" if month_first else "dmy"
    kinds = sniff(rows, header, default_order)
    orders = {}
    for i, (name, kind) in enumerate(zip(header, kinds)):
        values = [r[i] for r in rows if str(r[i]).strip()]
        order = date_order(values, default_order)
        if order == "mixed":
            notes.append(f"column {name!r} mixes day-first and month-first dates, left as text")
        elif kind == "date":
            orders[name] = order

    wb = Workbook()

    # ---- Data tab: the only place raw rows live, and the only thing you replace -------------------
    data_ws = wb.active
    data_ws.title = "Data"
    data_ws.append(header)
    for raw in rows:
        out = []
        for value, kind, name in zip(raw, kinds, header):
            if kind == "date":
                out.append(as_date(value, orders[name]) or value)
            elif kind == "number":
                number = as_number(value)
                out.append(value if number is None else number)
            else:
                out.append(value)
        data_ws.append(out)
    for cell in data_ws[1]:
        cell.font = white_bold
        cell.fill = navy_fill
    data_ws.freeze_panes = "A2"
    last_col = get_column_letter(width)
    table = Table(displayName="SourceData", ref=f"A1:{last_col}{len(rows) + 1}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    data_ws.add_table(table)
    for i, (name, kind) in enumerate(zip(header, kinds), 1):
        data_ws.column_dimensions[get_column_letter(i)].width = max(12, min(30, len(name) + 6))
        if kind == "date":
            for row in data_ws.iter_rows(min_row=2, min_col=i, max_col=i):
                row[0].number_format = "yyyy-mm-dd"
        elif kind == "number":
            for row in data_ws.iter_rows(min_row=2, min_col=i, max_col=i):
                row[0].number_format = "#,##0.00"

    # ---- Analysis tab ---------------------------------------------------------------------------
    ws = wb.create_sheet("Analysis", 0)
    ws["A1"] = "Analysis"
    ws["A1"].font = title_font
    ws["A2"] = "Every figure below reads the SourceData table on the Data tab. Paste new rows there and these update."
    ws["A2"].font = Font(italic=True, color="555555")

    line = 4
    ws.cell(line, 1, "Headline numbers").font = bold
    line += 1
    # ROWS of a table column counts every row in the table, blank cells included. A COUNTA here
    # would silently undercount the moment one row had an empty cell in that column.
    tiles = [("Rows of data", f"=ROWS({col(header[0])})", "#,##0")]
    # Date tiles first, so the cap below never drops "Latest", which is the one people look at.
    for name, kind in zip(header, kinds):
        if kind == "date":
            tiles.append((f"Earliest {name}", f"=IFERROR(MIN({col(name)}),0)", "yyyy-mm-dd"))
            tiles.append((f"Latest {name}", f"=IFERROR(MAX({col(name)}),0)", "yyyy-mm-dd"))
    for name, kind in zip(header, kinds):
        if kind == "number":
            tiles.append((f"Total {name}", f"=SUM({col(name)})", "#,##0.00"))
            tiles.append((f"Average {name}", f"=IFERROR(AVERAGE({col(name)}),0)", "#,##0.00"))
    for label, formula, fmt in tiles[:9]:
        ws.cell(line, 1, label).border = box
        ws.cell(line, 1).fill = light_fill
        cell = ws.cell(line, 2, formula)
        cell.border, cell.font, cell.number_format = box, bold, fmt
        line += 1

    # breakdown blocks, one per text column that behaves like a category
    numeric_cols = [n for n, k in zip(header, kinds) if k == "number"]
    measure = numeric_cols[0] if numeric_cols else None
    candidates = []
    for i, (name, kind) in enumerate(zip(header, kinds)):
        if kind != "text":
            continue
        counts = Counter(str(r[i]).strip() for r in rows if str(r[i]).strip())
        if 2 <= len(counts) <= MAX_CATEGORIES * 4:
            candidates.append((len(counts), name, counts))
    candidates.sort()

    chart_anchors = []
    line += 1
    for _, name, counts in candidates[:MAX_BREAKDOWNS]:
        ws.cell(line, 1, f"By {name}").font = bold
        line += 1
        head_row = line
        ws.cell(line, 1, name).font = white_bold
        ws.cell(line, 1).fill = navy_fill
        ws.cell(line, 2, "Rows").font = white_bold
        ws.cell(line, 2).fill = navy_fill
        if measure:
            ws.cell(line, 3, f"Total {measure}").font = white_bold
            ws.cell(line, 3).fill = navy_fill
        line += 1
        labels = [v for v, _ in counts.most_common(MAX_CATEGORIES)]
        first = line
        for value in labels:
            ws.cell(line, 1, value).border = box
            # Not COUNTIF. COUNTIF reads a category called ">65" or "<18" as a comparison and
            # returns 0 for it, and treats * and ? as wildcards. This compares the text exactly.
            count = ws.cell(line, 2, f"=SUMPRODUCT(--({col(name)}=$A{line}))")
            count.border, count.number_format = box, "#,##0"
            if measure:
                total = ws.cell(line, 3, f"=SUMPRODUCT(--({col(name)}=$A{line}),{col(measure)})")
                total.border, total.number_format = box, "#,##0.00"
            line += 1
        # Everything not listed above: categories beyond the first twelve, blank cells, and any
        # new category pasted in later. This is what makes the block always add up to the headline.
        ws.cell(line, 1, "Other").border = box
        ws.cell(line, 1).font = Font(italic=True)
        other = ws.cell(line, 2, f"=ROWS({col(name)})-SUM(B{first}:B{line - 1})")
        other.border, other.number_format = box, "#,##0"
        if measure:
            other_total = ws.cell(line, 3, f"=SUM({col(measure)})-SUM(C{first}:C{line - 1})")
            other_total.border, other_total.number_format = box, "#,##0.00"
        line += 1
        chart_anchors.append((name, head_row, line - 1, bool(measure)))
        line += 2

    for index, (name, head_row, last_row, has_measure) in enumerate(chart_anchors[:2]):
        chart = BarChart() if index == 0 else LineChart()
        chart.title = f"{'Total ' + measure if has_measure else 'Rows'} by {name}"
        chart.height, chart.width = 7.5, 15
        value_col = 3 if has_measure else 2
        values = Reference(ws, min_col=value_col, min_row=head_row, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=head_row + 1, max_row=last_row)
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(cats)
        chart.legend = None
        ws.add_chart(chart, f"F{5 + index * 16}")

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20

    # ---- Read me tab ----------------------------------------------------------------------------
    readme = wb.create_sheet("Read me")
    readme["A1"] = "How this workbook works"
    readme["A1"].font = title_font
    text = [
        ("Where to paste new data", ""),
        ("", "Go to the Data tab. Click cell A2. Paste your rows there. Keep the header row as it is."),
        ("", "If your new data has more rows than the table, the table grows by itself as you paste. "
             "If it has fewer, delete the leftover rows under it and the table shrinks."),
        ("", ""),
        ("How the summary updates", ""),
        ("", "Nothing on the Analysis tab points at a fixed range. Every figure reads the table by its "
             "name, SourceData, so the numbers follow whatever is in the table."),
        ("", "Excel recalculates as you paste. If you ever want to force it, press F9, or on a Mac "
             "press fn and F9."),
        ("", ""),
        ("Why there are no macros", ""),
        ("", "Macros are the usual reason a workbook behaves differently on macOS, and they make the "
             "file prompt for permission every time it opens. Everything here is ordinary Excel, so "
             "it opens and works the same on Windows, macOS and in the browser."),
        ("", ""),
        ("Adding a pivot table and a slicer", ""),
        ("", "1. Click any cell inside the table on the Data tab."),
        ("", "2. Insert, then PivotTable, then New worksheet. The source box already says SourceData, "
             "which is what keeps it pointed at the whole table however big it gets."),
        ("", "3. Drag the field you want to group by into Rows, and the number you want to measure "
             "into Values."),
        ("", "4. With the pivot selected, Insert, then Slicer, and tick the fields you want buttons for."),
        ("", "5. Right click the pivot, PivotTable Options, Data, and tick Refresh data when opening "
             "the file. After that it keeps itself current."),
        ("", ""),
        ("If a column is the wrong type", ""),
        ("", "Dates and numbers are detected from the values, not from the column name. If a column "
             "arrives as text, select it on the Data tab and set the format, then the tiles that use "
             "it will start calculating."),
        ("", "A date like 03/04/2026 was read day first, as 3 April, unless the column itself proved "
             "it was month first. If your file is month first, rebuild with --month-first."),
        ("", ""),
        ("The Other row in each breakdown", ""),
        ("", "A breakdown lists the categories that existed when the workbook was built, up to twelve. "
             "Other is everything else: further categories, blank cells, and any new category you "
             "paste in later. That is why each block always adds up to the Rows of data figure."),
    ]
    row = 3
    for heading, body in text:
        if heading:
            readme.cell(row, 1, heading).font = bold
            readme.cell(row, 1).fill = grey_fill
        if body:
            cell = readme.cell(row, 1, body)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    readme.column_dimensions["A"].width = 110

    wb.save(out_path)
    return out_path, len(rows), header, kinds, [c[1] for c in candidates[:MAX_BREAKDOWNS]], notes


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        raise SystemExit(__doc__.strip().splitlines()[-3])
    source = pathlib.Path(args[0])
    target = pathlib.Path(args[1]) if len(args) > 1 else source.with_suffix(".analysis.xlsx")
    path, n, header, kinds, blocks, notes = build(source, target, month_first="--month-first" in sys.argv)
    print(f"{path}: {n} rows, {len(header)} columns")
    print("  types     :", ", ".join(f"{h}={k}" for h, k in zip(header, kinds)))
    print("  breakdowns:", ", ".join(blocks) or "none")
    for note in notes:
        print("  note      :", note)
