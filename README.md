# Reusable Excel analysis workbook

Point this at any CSV and it builds an Excel workbook you can paste new data into. The summary
updates itself. No macros, no add-ins, and it behaves the same on Windows and macOS.

```
python build_analysis_workbook.py sales.csv
```

## The problem it solves

Most spreadsheet "dashboards" address a fixed block of cells. Paste 700 rows into a sheet built
for 500 and the summary silently ignores the last 200. Paste 300 and it counts empty rows as real
ones. Nobody notices, because the numbers still look like numbers.

Here nothing addresses a fixed range. The rows go into a real Excel Table named `SourceData`, and
every tile, breakdown and chart refers to it by name:

```
=SUM(SourceData[amount])
=COUNTIF(SourceData[country],$A12)
```

An Excel Table grows and shrinks with whatever you paste, so the formulas follow the data instead
of pointing at dead cells.

## Proof that it actually follows the data

Built from a 735 row file, then recalculated with LibreOffice:

| Figure | 735 rows | After pasting 50 more |
|---|---|---|
| Rows of data | 735 | 785 |
| Total amount | 336,731.99 | 341,731.99 |
| Latest date | 2025-08-22 | 2026-01-15 |
| Rows where status is active | 366 | 416 |

Not one formula was edited between those two columns. The 50 added rows carried 100.00 each, and
the total rose by exactly 5,000.00. The breakdowns also reconcile to the headline: 366 + 242 + 127
equals the 735 row count.

## What you get

Three tabs.

- `Analysis` headline tiles, a breakdown block for each column that behaves like a category, and
  native Excel charts. Not pictures of charts.
- `Data` your rows as the `SourceData` table. This is the only tab you replace.
- `Read me` where to paste, how it refreshes, why there are no macros, and the five steps to add
  a pivot table and a slicer.

## Column types are read from the values, not the header

With one deliberate exception. A header that names a label rather than a measurement stays as
text: `id`, `phone`, `postcode`, anything ending `_no` or `_code`, and any column whose values
mostly begin with a plus or a leading zero.

Storing a phone number as a number eats the leading plus. Summing an id means nothing. This is the
single most common way a "cleaned" spreadsheet arrives quietly broken.

Note that `last_order_amount` is money and `order_no` is a label, so the match is against the
whole header rather than any word inside it. There is a test for exactly that.

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

Nothing was published until all seven tests passed and a LibreOffice recalculation confirmed the
table in the section above: 50 rows added, every figure moved, no formula edited.

That order is the point. The interesting part of building software this way is not the generation,
it is deciding what "correct" means before anything is generated, and then refusing to ship until
it is demonstrated.

## Install and test

```
pip install openpyxl
python tests/test_build_analysis_workbook.py
```

Seven tests, no pytest required. The strongest one asserts that every formula written to the
Analysis tab goes through `SourceData[...]`, so a fixed range can never creep back in.

## Licence

MIT.
