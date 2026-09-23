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


# ---- from the adversarial audit of 23 September ------------------------------------------------
# Each of these failed against the first published version. They stay so it cannot regress.


def formulas_of(out):
    ws = load_workbook(out)["Analysis"]
    return [c.value for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and c.value.startswith("=")]


def test_duplicate_headers_are_made_unique():
    """Excel refuses a table with two columns of the same name and strips the table on open,
    which silently removes the one thing this workbook is built on."""
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    _, _, header, _, _, notes = b.build(write_csv([["name", "name", "NAME"], ["a", "b", "c"], ["d", "e", "f"]]), out)
    assert header == ["name", "name_2", "NAME_3"], header
    assert len(set(h.lower() for h in header)) == 3
    assert any("name_2" in n for n in notes), notes
    cols = [c.name for c in list(load_workbook(out)["Data"].tables.values())[0].tableColumns]
    assert len(set(cols)) == len(cols), cols


def test_padded_day_and_month_dates_are_dates():
    """03/04/2026 starts with a zero. The leading-zero rule used to win and the column stayed text,
    so a whole date column lost its Earliest and Latest tiles without a word."""
    assert b.sniff([["03/04/2026"], ["05/06/2026"], ["07/08/2026"]], ["when"]) == ["date"]


def test_small_decimals_are_numbers_not_labels():
    """0.8500 also starts with a zero and is longer than five characters. It is a number."""
    assert b.sniff([["0.8500"], ["0.1234"], ["0.9999"]], ["rate"]) == ["number"]


def test_postcode_and_phone_still_stay_text():
    assert b.sniff([["012345"], ["023456"], ["034567"]], ["zone"]) == ["text"]
    assert b.sniff([["+23057000001"], ["+33600000002"]], ["contact"]) == ["text"]


def test_date_order_is_decided_once_per_column():
    """A column with 12/25/2026 in it is month first everywhere, so 03/04/2026 in the same column is
    4 March. The old code read each cell on its own and put rows in different months."""
    assert b.date_order(["03/04/2026", "12/25/2026"]) == "mdy"
    assert b.date_order(["03/04/2026", "25/12/2026"]) == "dmy"
    assert b.date_order(["03/04/2026", "05/06/2026"]) == "dmy"
    assert b.date_order(["03/04/2026", "05/06/2026"], default="mdy") == "mdy"
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv([["when", "v"], ["03/04/2026", "1"], ["12/25/2026", "2"]]), out)
    cells = [c.value for c in load_workbook(out)["Data"]["A"][1:]]
    assert [str(c)[:10] for c in cells] == ["2026-03-04", "2026-12-25"], cells


def test_a_self_contradicting_date_column_is_left_as_text_and_reported():
    """25/12/2026 and 12/25/2026 in one column cannot both be right. Guessing would put some rows
    in the wrong month, so the column stays text and the build says why."""
    rows = [["when"], ["25/12/2026"], ["12/25/2026"], ["01/02/2026"]]
    assert b.sniff(rows[1:], rows[0]) == ["text"]
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    notes = b.build(write_csv([["when", "v"], ["25/12/2026", "1"], ["12/25/2026", "2"]]), out)[-1]
    assert any("mixes" in n for n in notes), notes


def test_numbers_in_the_formats_people_actually_export():
    got = [b.as_number(x) for x in ["$1,234.50", "€ 99", "1 000", "(1,200.00)", "-$5", "12 USD"]]
    assert got == [1234.5, 99.0, 1000.0, -1200.0, -5.0, 12.0], got
    # a percent is not stripped to a number a hundred times too big, and words are not numbers
    assert [b.as_number(x) for x in ["12%", "12 apples", "abc", ""]] == [None, None, None, None]
    # letters glued to digits are codes: the first version read SKU123 as 123 and summed it
    assert [b.as_number(x) for x in ["SKU123", "A100", "c14"]] == [None, None, None]
    assert b.sniff([["A100"], ["B200"], ["C300"]], ["code"]) == ["text"]


def test_breakdowns_never_use_countif_or_sumif():
    """COUNTIF reads a category called ">65" as a comparison and returns 0. Age bands, size bands
    and price bands are all written that way. The count must compare the text exactly."""
    rows = [["band", "amount"], [">65", "10"], ["<18", "20"], [">65", "30"], ["18-65", "40"]]
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv(rows), out)
    formulas = formulas_of(out)
    assert not any("COUNTIF" in f or "SUMIF(" in f for f in formulas), formulas
    assert any(f.startswith("=SUMPRODUCT(--(SourceData[band]=$A") for f in formulas), formulas


def test_row_count_does_not_depend_on_blank_cells():
    """The first text column is allowed to have empty cells. Counting its non-empty cells used to
    report fewer rows than the table had."""
    rows = [["name", "amount"], ["Ana", "1"], ["", "2"], ["Cy", "3"]]
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv(rows), out)
    formulas = formulas_of(out)
    assert "=ROWS(SourceData[name])" in formulas, formulas
    assert not any("SUBTOTAL" in f or "COUNTA" in f for f in formulas), formulas


def test_special_characters_in_headers_are_escaped_for_excel():
    """Excel requires a quote before $ - . , # [ ] and a few others inside SourceData[...]."""
    assert b.col("amount ($)") == "SourceData[amount ('$)]"
    assert b.col("a-b") == "SourceData[a'-b]"
    assert b.col("q[1]") == "SourceData[q'[1']]"
    assert b.col("plain name") == "SourceData[plain name]"
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv([["amount ($)", "k"], ["5", "a"], ["7", "b"]]), out)
    assert "=SUM(SourceData[amount ('$)])" in formulas_of(out)


def test_every_breakdown_ends_with_an_other_row():
    """Twelve categories are listed. Without an Other row the block stops adding up to the
    headline as soon as there is a thirteenth, or a new one is pasted in later."""
    rows = [["cat", "amount"]] + [[f"c{i % 15}", "1"] for i in range(60)]
    out = pathlib.Path(tempfile.mkdtemp()) / "out.xlsx"
    b.build(write_csv(rows), out)
    ws = load_workbook(out)["Analysis"]
    labels = [ws.cell(r, 1).value for r in range(1, ws.max_row + 1)]
    assert "Other" in labels, labels
    other = ws.cell(labels.index("Other") + 1, 2).value
    assert other.startswith("=ROWS(SourceData[cat])-SUM(B"), other


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
