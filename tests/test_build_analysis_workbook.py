"""Tests for build_analysis_workbook. No pytest needed: python tests/test_build_analysis_workbook.py"""
import csv
import pathlib
import sys
import tempfile

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import build_analysis_workbook as b
from openpyxl import load_workbook


def write_csv(rows):
    d = tempfile.mkdtemp()
    p = pathlib.Path(d) / "in.csv"
    with open(p, "w", newline="", encoding="utf-8") as fh:
        csv.writer(fh).writerows(rows)
    return p


ROWS = [
    ["id", "name", "country", "phone", "signup_date", "amount"],
    ["001", "Ana", "Mauritius", "+23057000001", "2026-01-05", "120.50"],
    ["002", "Bea", "France", "+33600000002", "2026-02-11", "80.00"],
    ["003", "Cy", "Mauritius", "+23057000003", "2026-03-20", "45.25"],
    ["004", "Dee", "Senegal", "+22177000004", "2026-04-02", "200.00"],
]


def test_types_detected_from_values():
    kinds = b.sniff([r for r in ROWS[1:]], ROWS[0])
    got = dict(zip(ROWS[0], kinds))
    assert got["amount"] == "number", got
    assert got["signup_date"] == "date", got
    assert got["country"] == "text", got


def test_label_columns_stay_text():
    """An id is not a quantity and a phone is not a number: summing one is meaningless and
    storing the other as a number eats the leading plus."""
    kinds = dict(zip(ROWS[0], b.sniff([r for r in ROWS[1:]], ROWS[0])))
    assert kinds["id"] == "text", kinds
    assert kinds["phone"] == "text", kinds


def test_money_column_is_not_mistaken_for_a_label():
    header = ["last_order_amount", "order_no"]
    rows = [["10.00", "A1"], ["20.00", "A2"], ["30.00", "A3"]]
    kinds = dict(zip(header, b.sniff(rows, header)))
    assert kinds["last_order_amount"] == "number", kinds
    assert kinds["order_no"] == "text", kinds


def test_iso_datetime_is_a_date():
    header = ["scraped_at"]
    rows = [["2026-09-18T12:00:00Z"], ["2026-09-19T09:30:00Z"], ["2026-09-20T23:59:59Z"]]
    assert b.sniff(rows, header) == ["date"]


def test_workbook_structure_and_formulas():
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv(ROWS), out)
    wb = load_workbook(out)
    assert wb.sheetnames == ["Analysis", "Data", "Read me"], wb.sheetnames

    data = wb["Data"]
    table = list(data.tables.values())[0]
    assert table.displayName == "SourceData", table.displayName

    analysis = wb["Analysis"]
    formulas = [c.value for row in analysis.iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith("=")]
    assert formulas, "no formulas written"
    # every formula must go through the table by name, never a fixed range
    assert all("SourceData[" in f for f in formulas), [f for f in formulas if "SourceData[" not in f]


def test_read_me_tab_explains_where_to_paste():
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv(ROWS), out)
    text = " ".join(str(c.value) for row in load_workbook(out)["Read me"].iter_rows()
                    for c in row if c.value)
    for phrase in ("Data tab", "SourceData", "pivot", "macros"):
        assert phrase.lower() in text.lower(), phrase


def test_empty_csv_is_rejected_clearly():
    try:
        b.build(write_csv([ROWS[0]]), pathlib.Path(tempfile.mkdtemp()) / "out.xlsx")
    except SystemExit as e:
        assert "no data rows" in str(e)
    else:
        raise AssertionError("expected a SystemExit for a header-only file")


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn(); print("ok  ", name)
        except Exception as e:
            failed += 1; print("FAIL", name, repr(e))
    print(f"{len(tests) - failed} passed, {failed} failed")
    sys.exit(1 if failed else 0)
