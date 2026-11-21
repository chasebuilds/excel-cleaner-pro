# Limitations

What this tool does not handle, or handles in a way you should know about.

## File structure

- The header row is found by a rule: the first row at least half filled with text. A title row that
  is itself half filled with text will be taken as the header. Set `--header-row N` when the
  `Summary` notes show the wrong row.
- Multi-row headers are reported, not combined. Group labels above the header go to `Removed`.
- Merged cells are flagged, not filled down. The value stays in the first cell of the range and the
  others stay empty.
- Total lines are flagged when a cell says `Total`, `Subtotal`, `Grand total` or `Sum`. A total row
  with no such label is not detected. Flagged rows are kept, so control totals include them.
- Hidden sheets, rows and columns are read like any other and listed in the notes. The tool does
  not decide whether hidden data is in scope.
- XLSX formulas are read as the last value Excel calculated. A formula with no stored value is
  flagged `formula_no_value` and the cell stays empty. The tool does not calculate formulas.
- Cell colors, comments, column widths, charts, and macros in the input are not carried over.
- The whole file is held in memory and an XLSX file is loaded twice (values, then formulas).
  Tested to 21,000 rows by 10 columns.
- More than 1,048,575 data rows in any output sheet is refused with an error. Split the input.
- Old `.xls` files are not supported. Save as `.xlsx` first.
- Output is XLSX only. No CSV output.

## Identifiers Excel has already damaged

- These are flagged, never repaired: ID columns stored as numbers, ZIP columns with 3 or 4 digit
  values, numbers of 16 or more digits, scientific notation text (`1.23E+15`), and codes that look
  like Excel date conversions (`2-Sep`).
- Each is one `Issues` line per column, with a count and the first five rows, not one line per cell.
- A column is treated as an ID only when its header says so (id, sku, code, zip, postal, number, no,
  acct) or its values have leading zeros. A code column with another header is checked only for
  scientific notation and date-like codes.
- Long codes stored as text are flagged as truncated only when every one of them ends in 0.

## Encoding

- CSV encoding cannot be detected with certainty. The tool tries UTF-8, then falls back to
  Windows-1252, and prints which one it used on `Summary`. Other encodings (UTF-16, Shift JIS,
  Mac Roman) will read wrong.
- Text that looks like UTF-8 read through a Western code page is flagged `possible_mojibake`.
  It is not repaired. Other kinds of encoding damage are not detected.
- Text is normalized to Unicode NFC (logged as `unicode_nfc`). Look-alike characters from other
  alphabets and full-width characters are left as they are.

## Names

- Case is fixed only when the whole value is ALL CAPS or all lower case. Mixed case is left alone.
- Capitals inside a surname cannot be rebuilt from ALL CAPS. `MACDONALD` becomes `Macdonald`,
  `DE LA CRUZ` becomes `De La Cruz`, `LEBLANC` becomes `Leblanc`. The benchmark measures this:
  about 96% of case errors come back right. Every change is in `Changes` under `name_case`.
- A column is treated as a person's name only when its header says so (`name`, `contact`, `surname`)
  and does not also say company, product, user, city and so on. Company names are never re-cased.
  A header like `Customer` with no "name" in it is treated as plain text unless you set its type in
  the config.

## Dates

- A column is read as dates when its header says so (`Date`, `DOB`), or with no such header when
  at least 80% of filled cells look like dates and 60% are readable dates. A messier column than
  that stays text; set its type in the config (`"Last Order" = "date"`).
- Read: ISO, `Y/M/D`, numeric with `/`, `-` or `.`, and English month names (`March 5, 2024`,
  `5-Mar-2024`, `05 March 2024`). Not read: other languages, `20240305`, week numbers, quarters,
  month and year alone, Excel serial numbers stored as text. These are flagged `unparsed_date`.
- A date that could be day-first or month-first is converted only when the rest of the column
  proves the order (at least 3 values that can only be read one way, none the other way). That is
  an inference from the column, not a certainty. It is logged as `date_iso_inferred_order` and noted
  in `Summary`. Use `--date-order flag` to turn it off.
- Two-digit years use a pivot: 00 to 49 is 2000s, 50 to 99 is 1900s. Logged as
  `date_iso_2digit_year`. Change `two_digit_year_pivot` if that is wrong for birth dates.
- Dates are written as ISO text, not Excel date cells. Time zones are dropped.
- A real Excel date cell is turned into ISO text without a `Changes` entry, because the day is the
  same.

## Numbers and currency

- Read: `$`, `€`, `£`, `¥`, the codes USD EUR GBP CAD AUD, thousands separators in US or European
  style, decimal comma, accounting negatives in parentheses.
- Not read: percentages, fractions, scientific notation, units (`12 kg`), written-out numbers. These
  are flagged `unparsed_number`.
- The currency symbol is dropped. A column that mixes dollars and euros becomes one column of plain
  numbers with no record of which was which, other than the before value in `Changes`.
- `1.250` with no other clue is read as 1.25. It is flagged as ambiguous only when the same column
  also holds a clearly European value such as `1.200,50`. In a column with no such value, a
  European whole number written `1.250` will be read wrong and will not be flagged.
- A plain number stored as text (`1250.00`) becomes a real number without a `Changes` entry.
- Control totals show that no amount went missing between reading and output. They are computed
  after parsing, so they do not prove a value was parsed correctly. Cells that could not be read as
  numbers are counted, not summed.

## Phones, states, emails

- Phones: United States and Canada only. Other countries are flagged `non_us_phone` and left as
  typed. No check that the number is in service.
- States: US states, DC and five territories. No Canadian provinces, no countries.
- Emails: syntax check only. No check that the domain or mailbox exists. Typos in a valid-looking
  address (`gmial.com`) are not caught.

## Placeholders

- `na`, `none`, `null`, `-` and similar are turned into empty cells by default. If a column uses one
  of them as a real value (NA for Namibia, say), remove it from `placeholders.tokens`. `Summary`
  lists how many of each were blanked in each column so this is easy to spot.

## Duplicates

- Exact duplicates are compared after cleaning, across all columns. Two rows that differ in any
  column, including an ID column, are not exact duplicates.
- Fuzzy matches are only listed by default. Removal is opt-in (`--fuzzy-remove`).
- Fuzzy matching runs only when the tool finds at least two name or email columns. For other
  tables, set `dedupe.fuzzy_keys`.
- Rows with a blank key column are never fuzzy-matched. `Summary` says how many were skipped.
- With removal on, the first row seen is kept, not the most complete one. Its blanks are filled
  from the removed row. Where both rows have a value, the kept row's value wins and the other stays
  in `Removed`. There is no rule to prefer the newer or longer value.
- With three or more rows for one customer, each later row is compared with the kept row only.
- Speed: each row is compared with the rows kept before it. Measured: 5,297 rows in about 3 seconds,
  21,192 rows in about 53 seconds. Past roughly 50,000 rows, run with `--no-fuzzy` or expect minutes.
- The benchmark plants one-typo duplicates. Nicknames (Bob and Robert), swapped fields, and married
  names are not tested and will mostly be missed.

## Validation rules

- Five rule kinds: unique, required, allowed values, range, pattern. No rules across columns
  (end date after start date), no rules across sheets.
- Rules run on the cleaned data. A failed rule is reported and nothing is changed.

## Not built

- A `DELIVERY.md` handover file per job.
- Optional repair of encoding damage.
- Review workflow: there is no way to feed decisions on the `Issues` sheet back in.

## Benchmark scope

The benchmark scores the formats its generator plants, on a CSV. It does not prove the tool handles
formats that are not in `benchmark/generate.py`, and it does not exercise the Excel-only checks
(merged cells, hidden rows, formulas, IDs stored as numbers), which are covered by the test suite.
Its numbers describe generated data and do not transfer to a client file. Tests pass on Python
3.12, 3.13 and 3.14 on Windows. Python 3.11, macOS and Linux are covered only by the CI file, which
has not run yet.
