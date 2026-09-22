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

Tabs produced:
  Analysis   headline tiles, one breakdown block per usable text column, native Excel charts
  Data       the rows as the SourceData table, which is where new data is pasted
  Read me    where to paste, how it refreshes, and how to add a pivot table and a slicer

Usage:
  python build_analysis_workbook.py <input.csv> [output.xlsx]

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

DATE_PATTERNS = ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d-%m-%Y")
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
MAX_CATEGORIES = 12         # rows per breakdown block


def as_date(value):
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
    for pattern in DATE_PATTERNS:
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def as_number(value):
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    text = re.sub(r"^[^\d\-.]+", "", text)      # currency symbols and the like
    if text in ("", "-", "."):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def sniff(rows, header):
    """Decide what each column is, from the values rather than from the header name."""
    kinds = []
    for i, name in enumerate(header):
        values = [r[i] for r in rows if i < len(r) and str(r[i]).strip() != ""]
        if not values:
            kinds.append("empty")
            continue
        if LABEL_HEADERS.search(str(name)):
            kinds.append("text")
            continue
        keeps_shape = sum(1 for v in values
                          if str(v).strip().startswith(("+", "0")) and len(str(v).strip()) > 5)
        if keeps_shape >= 0.3 * len(values):    # a plus or a leading zero that must survive
            kinds.append("text")
            continue
        dates = sum(1 for v in values if as_date(v) is not None)
        numbers = sum(1 for v in values if as_number(v) is not None)
        if dates >= 0.8 * len(values):
            kinds.append("date")
        elif numbers >= 0.8 * len(values):
            kinds.append("number")
        else:
            kinds.append("text")
    return kinds


def safe_name(name, used):
    """Excel table and defined names cannot contain spaces or most punctuation."""
    clean = re.sub(r"\W+", "_", str(name)).strip("_") or "column"
    if clean[0].isdigit():
        clean = "c_" + clean
    while clean.lower() in used:
        clean += "_x"
    used.add(clean.lower())
    return clean


def build(csv_path, out_path):
    with open(csv_path, newline="", encoding="utf-8-sig", errors="replace") as fh:
        reader = csv.reader(fh)
        header = next(reader)
        rows = [r for r in reader if any(str(c).strip() for c in r)]
    if not rows:
        raise SystemExit(f"{csv_path} has a header but no data rows")

    header = [h.strip() or f"column_{i}" for i, h in enumerate(header, 1)]
    width = len(header)
    rows = [(r + [""] * width)[:width] for r in rows]
    kinds = sniff(rows, header)

    wb = Workbook()

    # ---- Data tab: the only place raw rows live, and the only thing you replace -------------------
    data_ws = wb.active
    data_ws.title = "Data"
    data_ws.append(header)
    for raw in rows:
        out = []
        for value, kind in zip(raw, kinds):
            if kind == "date":
                out.append(as_date(value) or value)
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
    first_text = next((header[i] for i, k in enumerate(kinds) if k == "text"), header[0])
    tiles = [("Rows of data", f"=SUBTOTAL(103,SourceData[{first_text}])", "#,##0")]
    for name, kind in zip(header, kinds):
        if kind == "number":
            tiles.append((f"Total {name}", f"=SUM(SourceData[{name}])", "#,##0.00"))
            tiles.append((f"Average {name}", f"=IFERROR(AVERAGE(SourceData[{name}]),0)", "#,##0.00"))
        elif kind == "date":
            tiles.append((f"Earliest {name}", f"=IFERROR(MIN(SourceData[{name}]),0)", "yyyy-mm-dd"))
            tiles.append((f"Latest {name}", f"=IFERROR(MAX(SourceData[{name}]),0)", "yyyy-mm-dd"))
    for label, formula, fmt in tiles[:8]:
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
        for value in labels:
            ws.cell(line, 1, value).border = box
            count = ws.cell(line, 2, f'=COUNTIF(SourceData[{name}],$A{line})')
            count.border, count.number_format = box, "#,##0"
            if measure:
                total = ws.cell(line, 3, f'=SUMIF(SourceData[{name}],$A{line},SourceData[{measure}])')
                total.border, total.number_format = box, "#,##0.00"
            line += 1
        chart_anchors.append((name, head_row, line - 1, bool(measure)))
        line += 2

    for index, (name, head_row, last_row, has_measure) in enumerate(chart_anchors[:2]):
        chart = BarChart() if index == 0 else LineChart()
        chart.title = f"{'Total ' + measure if has_measure else 'Rows'} by {name}"
        chart.height, chart.width = 7.5, 15
        col = 3 if has_measure else 2
        values = Reference(ws, min_col=col, min_row=head_row, max_row=last_row)
        cats = Reference(ws, min_col=1, min_row=head_row + 1, max_row=last_row)
        chart.add_data(values, titles_from_data=True)
        chart.set_categories(cats)
        chart.legend = None
        ws.add_chart(chart, f"F{5 + index * 16}")

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 20

    # ---- Read me tab ----------------------------------------------------------------------------
    notes = wb.create_sheet("Read me")
    notes["A1"] = "How this workbook works"
    notes["A1"].font = title_font
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
    ]
    row = 3
    for heading, body in text:
        if heading:
            notes.cell(row, 1, heading).font = bold
            notes.cell(row, 1).fill = grey_fill
        if body:
            cell = notes.cell(row, 1, body)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        row += 1
    notes.column_dimensions["A"].width = 110

    wb.save(out_path)
    return out_path, len(rows), header, kinds, [c[1] for c in candidates[:MAX_BREAKDOWNS]]


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__.strip().splitlines()[-3])
    source = pathlib.Path(sys.argv[1])
    target = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else source.with_suffix(".analysis.xlsx")
    path, n, header, kinds, blocks = build(source, target)
    print(f"{path}: {n} rows, {len(header)} columns")
    print("  types     :", ", ".join(f"{h}={k}" for h, k in zip(header, kinds)))
    print("  breakdowns:", ", ".join(blocks) or "none")
