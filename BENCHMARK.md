# Benchmark

Written by `benchmark/run_benchmark.py` from a real run. Do not edit by hand.

- Data: `benchmark/generate.py --rows 5000 --seed 42` (5000 clean customer rows, then damaged: 5297 rows and 52970 cells in the messy file)
- Config: defaults, no config file, unless a section says otherwise
- Python 3.14.3 on Windows AMD64

## Result by error type, default settings

Recall: of the errors planted, how many ended up right. Precision: of the times the tool took that action, how many were right. "Right" for a fix means the cell equals the truth table. "Right" for a flag means the cell was left as typed and listed in `Issues`. By default a fuzzy duplicate is never removed: it is listed in `Issues` with the row it resembles.

| Error type | Expected | Planted | Handled | Recall | Tool actions | Correct | Precision |
|---|---|---:|---:|---:|---:|---:|---:|
| Extra whitespace | fix | 762 | 762 | 100.0% | 762 | 762 | 100.0% |
| Placeholder text ("N/A", "-", "null") | fix | 692 | 692 | 100.0% | 692 | 692 | 100.0% |
| Name in ALL CAPS or lower case | fix | 307 | 295 | 96.1% | 307 | 295 | 96.1% |
| Email with capitals | fix | 228 | 228 | 100.0% | 228 | 228 | 100.0% |
| Phone in a non-standard format | fix | 545 | 545 | 100.0% | 545 | 545 | 100.0% |
| Date in a non-ISO format, one reading | fix | 1986 | 1986 | 100.0% | 1986 | 1986 | 100.0% |
| Date readable two ways, column is all month-first | fix | 206 | 206 | 100.0% | 206 | 206 | 100.0% |
| Currency text ("$1,200.50", "1.200,50") | fix | 493 | 493 | 100.0% | 493 | 493 | 100.0% |
| State name or lower-case code | fix | 408 | 408 | 100.0% | 408 | 408 | 100.0% |
| Invalid email | flag, leave as typed | 93 | 93 | 100.0% | 93 | 93 | 100.0% |
| Invalid phone | flag, leave as typed | 92 | 92 | 100.0% | 92 | 92 | 100.0% |
| Date readable two ways, column mixes orders | flag, leave as typed | 183 | 183 | 100.0% | 183 | 183 | 100.0% |
| Not a real date ("02/30/2024", "Q1 2024") | flag, leave as typed | 94 | 94 | 100.0% | 94 | 94 | 100.0% |
| Number readable two ways ("1,500" / "1.500") | flag, leave as typed | 98 | 98 | 100.0% | 98 | 98 | 100.0% |
| Misspelled state | flag, leave as typed | 51 | 51 | 100.0% | 51 | 51 | 100.0% |
| Exact duplicate row | remove, log in Removed | 150 | 150 | 100.0% | 150 | 150 | 100.0% |
| Fuzzy duplicate row, listed for review | flag, keep the row | 122 | 122 | 100.0% | 127 | 122 | 96.1% |
| Empty row | remove, log in Removed | 25 | 25 | 100.0% | 25 | 25 | 100.0% |

## Whole-file checks, default settings

| Check | Result |
|---|---:|
| Cells damaged on purpose | 6238 |
| Cells logged in `Changes` | 5627 |
| Changes that left a cell different from the truth | 12 |
| Cells still different from the truth after cleaning | 12 |
| Rows removed that were not a planted duplicate or empty row | 0 |
| Every input row is in `Cleaned` or `Removed` | yes |
| Control totals balance on every numeric column | yes |
| Flags of a type the generator did not plant | 0  |

## With fuzzy removal turned on (`--fuzzy-remove`)

Opt-in. A match scoring 95 or more is removed and the kept row takes the removed row's value wherever its own cell is empty. False removals and missed duplicates are counted separately.

| Error type | Expected | Planted | Handled | Recall | Tool actions | Correct | Precision |
|---|---|---:|---:|---:|---:|---:|---:|
| Fuzzy duplicate row, removed | remove, log in Removed | 122 | 121 | 99.2% | 121 | 121 | 100.0% |
| Fuzzy duplicate row, removed or listed for review | remove or flag | 122 | 122 | 100.0% | 127 | 122 | 96.1% |

| Check | Result |
|---|---:|
| Real customer rows removed by mistake (false removals) | 0 |
| Planted fuzzy duplicates neither removed nor listed (missed) | 0 |
| Cells filled into a kept row from a removed row | 0 |
| Cells different from the truth after cleaning | 12 |
| Control totals balance on every numeric column | yes |

| Planted slip | Removed | Listed for review | Missed |
|---|---:|---:|---:|
| One typo in the name | 81 | 1 | 0 |
| One typo in the email | 40 | 0 | 0 |

## Speed

End to end: read the CSV, clean, write the XLSX. Median of 3 runs, one process.

| Run | Rows | Seconds | Rows per second |
|---|---:|---:|---:|
| Defaults (fuzzy duplicate matching on) | 5297 | 3.06 | 1,729 |
| `--no-fuzzy` | 5297 | 1.13 | 4,691 |
| Defaults, larger file (1 run) | 21192 | 41.80 | 507 |
| `--no-fuzzy`, larger file (1 run) | 21192 | 3.71 | 5,705 |

Fuzzy matching compares each row with the rows kept before it, so its cost grows with the square of the row count. A rows-per-second figure with fuzzy on holds for that file size only.

## Other seeds

The same test on data the rules were not tuned against. Seeds 1 to 5 were looked at while building the fuzzy rules, so they are not used here. The result columns up to 'Fuzzy listed, precision' are from default settings, the last two from `--fuzzy-remove`.

| Seed | Rows | Wrong changes | Cells wrong after | Name case recall | Fuzzy listed, recall | Fuzzy listed, precision | Removal on: rows wrongly removed | Removal on: fuzzy removed, recall |
|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| 101 | 5298 | 13 | 13 | 95.8% | 100.0% | 98.4% | 0 | 99.2% |
| 102 | 5298 | 9 | 9 | 97.0% | 100.0% | 97.6% | 0 | 99.2% |
| 103 | 5300 | 7 | 7 | 97.7% | 100.0% | 98.4% | 0 | 99.2% |
| 104 | 5299 | 9 | 9 | 97.1% | 100.0% | 96.1% | 0 | 96.8% |
| 105 | 5298 | 8 | 8 | 97.4% | 100.0% | 95.3% | 0 | 99.2% |

## Column types the tool inferred

| Column | Type (decided by) |
|---|---|
| customer_id | id (header) |
| full_name | name (header) |
| email | email (header) |
| phone | phone (header) |
| city | text (default) |
| state | state (header) |
| signup_date | date (header) |
| last_order_date | date (header) |
| lifetime_value | currency (header) |
| orders | number (values) |

## What this does not show

The generator plants the formats listed in `benchmark/generate.py`. A real file can hold formats that are not on that list, and those are not scored here. `LIMITATIONS.md` lists the known gaps. The surnames that cannot be rebuilt from ALL CAPS (MacDonald, de la Cruz, LeBlanc) are in the data on purpose, which is why the name row is below 100%. Every planted fuzzy duplicate has both key columns filled: rows with a blank key are not fuzzy-matched at all, so that gap is not measured here. The data is a CSV, so the checks that only apply to Excel files (merged cells, hidden rows, formulas with no stored value, IDs stored as numbers) are covered by the test suite, not by this benchmark. These numbers describe this generated data. On a client file the same figures have to be measured on that file.

## Reproduce

```
python benchmark/run_benchmark.py --rows 5000 --seed 42
```

Counts repeat exactly for the same seed. Timings vary by machine.
