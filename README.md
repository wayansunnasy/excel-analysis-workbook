# Reusable Excel analysis workbook

Point this at any CSV and it builds an Excel workbook you can paste new data into. The summary
updates itself. No macros, no add-ins, and it behaves the same on Windows and macOS.

```
python build_analysis_workbook.py sales.csv
python build_analysis_workbook.py sales.csv --month-first
```

## The problem it solves

Most spreadsheet "dashboards" address a fixed block of cells. Paste 700 rows into a sheet built
for 500 and the summary silently ignores the last 200. Paste 300 and it counts empty rows as real
ones. Nobody notices, because the numbers still look like numbers.

Here nothing addresses a fixed range. The rows go into a real Excel Table named `SourceData`, and
every tile, breakdown and chart refers to it by name:

```
=SUM(SourceData[amount])
=SUMPRODUCT(--(SourceData[country]=$A12))
```

An Excel Table grows and shrinks with whatever you paste, so the formulas follow the data instead
of pointing at dead cells.

## Proof that it actually follows the data

Built from the 500 row file in `examples/`, recalculated with LibreOffice, then 50 rows were pasted
under the table and it was recalculated again. The pasted rows carried a price of 100.00, a stock
count of 1, a date of 2026-09-23 and a category that did not exist before.

| Figure | 500 rows | After pasting 50 more |
|---|---|---|
| Rows of data | 500 | 550 |
| Total price | 17,520.20 | 22,520.20 |
| Total stock_count | 6,833 | 6,883 |
| Latest scraped_at_utc | 2026-09-17 | 2026-09-23 |
| Category breakdown, Other row | 131 | 181 |

Not one formula was edited between those two columns. The total rose by exactly 5,000.00, the
stock count by exactly 50, and the 50 rows of the new category landed in the breakdown's `Other`
row, so the block still adds up to the headline. It did both times: the twelve category rows plus
`Other` sum to 500 before and 550 after. No formula returned an error in either run.

The whole thing is one command if you have LibreOffice installed:

```
python tests/paste_proof.py
```

## What you get

Three tabs.

- `Analysis` headline tiles, a breakdown block for each column that behaves like a category, and
  native Excel charts. Not pictures of charts.
- `Data` your rows as the `SourceData` table. This is the only tab you replace.
- `Read me` where to paste, how it refreshes, why there are no macros, and the five steps to add
  a pivot table and a slicer.

Every breakdown ends with an `Other` row. It holds the categories beyond the first twelve, blank
cells, and anything new you paste in later, which is what keeps each block adding up to the row
count instead of quietly drifting away from it.

## Column types are read from the values, not the header

With one deliberate exception. A header that names a label rather than a measurement stays as
text: `id`, `phone`, `postcode`, anything ending `_no` or `_code`, and any column whose values
mostly begin with a plus or a leading zero.

Storing a phone number as a number eats the leading plus. Summing an id means nothing. This is the
single most common way a "cleaned" spreadsheet arrives quietly broken.

Note that `last_order_amount` is money and `order_no` is a label, so the match is against the
whole header rather than any word inside it. There is a test for exactly that. Letters glued to
digits are codes too: `SKU123` is not 123.

Numbers are read the way a person reads them: `1,234.50`, `1 000`, `$99`, `12 USD` and
`(1,200.00)` are all numbers. `12%` is left as text on purpose, because reading it as 12 would
be wrong by a factor of a hundred.

**Dates.** `03/04/2026` is ambiguous. Each date column is decided once, as a whole: if any value in
it can only be day first, such as `25/12/2026`, the column is day first; if any can only be month
first, the column is month first; and if it contains both, it is left as text and the build says
so, because reading it either way would put some rows in the wrong month. A column that never
settles the question is read day first unless you pass `--month-first`.

**Headers.** Two columns with the same name would make Excel strip the table on open, so the
second becomes `name_2` and the build tells you. Characters that Excel needs escaped inside a
table reference, such as `$`, `-` and `.`, are escaped.

## What it does not do

`openpyxl` cannot create pivot tables or slicers, so those are added by hand in Excel. The
`Read me` tab carries the five step instruction, including ticking "Refresh data when opening the
file" so the pivot keeps itself current.

## How this was built

Specification first, then constraints, then acceptance tests, then code.

The constraints were fixed before a line existed: no fixed ranges anywhere, no macros, identical
behaviour on Windows and macOS, and label columns never coerced to numbers. Each one became a test.
The implementation was then produced with an AI-assisted workflow and held to those tests rather
than to an opinion about whether it looked right.

The first version passed its seven tests and a LibreOffice recalculation, and was published. A
separate adversarial pass two days later, run against the published code with hostile input rather
than friendly input, found eight defects the friendly tests had not reached:

- a padded date such as `03/04/2026` was classed as a label because it starts with a zero, so the
  column lost its date tiles without a word
- each cell in a date column was read on its own, so `03/04/2026` and `12/25/2026` in the same
  column landed in different months
- `SKU123` was read as the number 123 and summed
- a category called `>65` or `<18` counted as zero, because `COUNTIF` reads it as a comparison
- the row count counted the non-empty cells of the first text column, so one blank cell undercounted it
- two columns with the same header produced a table Excel would strip on open
- a header such as `amount ($)` produced a reference Excel cannot parse
- a thirteenth category, or a new one pasted in later, vanished from the breakdown

Every one of them is now a test, and the fixes are in this version. That is the part of the method
worth copying: the friendly tests prove the design, the hostile pass proves the code, and nothing
is called finished until both have been run.

## Install and test

```
pip install openpyxl
python tests/test_build_analysis_workbook.py
```

Eighteen tests, no pytest required. The strongest one asserts that every formula written to the
Analysis tab goes through `SourceData[...]`, so a fixed range can never creep back in.

## Licence

MIT.
