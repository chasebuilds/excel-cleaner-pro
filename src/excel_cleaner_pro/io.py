"""Write the audited workbook."""

from __future__ import annotations

from collections import Counter
from collections.abc import Set
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .config import Config
from .model import CleanResult, SheetResult
from .reader import read_tables  # noqa: F401  (kept importable from here)

_HEADER_FONT = Font(bold=True, color="FFFFFF")
_HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
# one header row plus this many data rows is all an Excel sheet can hold
EXCEL_MAX_DATA_ROWS = 1_048_575


def _write_sheet(
    ws: Worksheet, headers: list[str], rows: list[list[Any]], number_cols: Set[int] = frozenset()
) -> None:
    if len(rows) > EXCEL_MAX_DATA_ROWS:
        raise ValueError(
            f"sheet {ws.title!r} would need {len(rows):,} rows; an Excel sheet holds "
            f"{EXCEL_MAX_DATA_ROWS:,}. Split the input file and run each part."
        )
    ws.append(headers)
    for cell in ws[1]:
        cell.font, cell.fill = _HEADER_FONT, _HEADER_FILL
        cell.alignment = Alignment(vertical="center")
    widths = [len(str(h)) for h in headers]
    for r, row in enumerate(rows, start=2):
        for c, value in enumerate(row, start=1):
            if value is None:
                continue
            if isinstance(value, list | tuple | dict):
                value = str(value)
            cell = ws.cell(row=r, column=c, value=value)
            # openpyxl saves text that starts with = as a live formula unless told otherwise
            if isinstance(value, str) and value.startswith("="):
                cell.data_type = "s"
            if c - 1 in number_cols and isinstance(value, float):
                cell.number_format = "#,##0.00"
            if r <= 200:
                widths[c - 1] = max(widths[c - 1], len(str(value)))
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 8), 60)
    ws.freeze_panes = "A2"
    if headers and rows:
        ws.auto_filter.ref = ws.dimensions


def _write_lines(ws: Worksheet, rows: list[list[Any]], first_width: int = 34) -> None:
    for row in rows:
        ws.append(row)
    ws.column_dimensions["A"].width = first_width
    for letter in "BCDEFGHI":
        ws.column_dimensions[letter].width = 18
    for cells in ws.iter_rows(min_col=1, max_col=1):
        cell = cells[0]
        if isinstance(cell.value, str) and cell.value.startswith("## "):
            cell.value = cell.value[3:]
            cell.font = Font(bold=True)


def _unique_title(base: str, used: set[str]) -> str:
    bad = set("[]:*?/\\")
    title = "".join("_" if ch in bad else ch for ch in base)[:31]
    n = 2
    while title.lower() in used:
        tail = f" ({n})"
        title = title[: 31 - len(tail)] + tail
        n += 1
    used.add(title.lower())
    return title


def _flatten(raw: dict[str, Any], prefix: str = "") -> list[list[Any]]:
    out: list[list[Any]] = []
    for key, value in raw.items():
        if isinstance(value, dict):
            out += _flatten(value, f"{prefix}{key}.")
        else:
            out.append(
                [f"{prefix}{key}", value if isinstance(value, int | float | str) else str(value)]
            )
    return out


def summary_rows(result: CleanResult, cfg: Config) -> list[list[Any]]:
    removed = Counter(r.reason for r in result.removed)
    rules = Counter(c.rule for c in result.changes)
    issues = Counter(i.issue for i in result.issues)
    unaccounted = result.rows_in - result.rows_out - len(result.removed)
    rows: list[list[Any]] = [
        ["## Reconciliation"],
        ["Source file", result.source],
        ["Rows read (header not counted)", result.rows_in],
        ["Rows in Cleaned", result.rows_out],
        ["Rows in Removed", len(result.removed)],
        ["Read - Cleaned - Removed (must be 0)", unaccounted],
        ["Cells changed", len(result.changes)],
        ["Items flagged for review", len(result.issues)],
        [],
        ["## Control totals, numeric columns"],
        ["sheet", "column", "total before rows were removed", "total in Cleaned",
         "total in Removed", "moved into Cleaned by survivor fill", "balance (must be 0)",
         "cells that are not numbers"],
    ]  # fmt: skip
    for sheet in result.sheets:
        for t in sheet.control_totals:
            rows.append(
                [sheet.name, t.column, t.total_before_removal, t.in_cleaned, t.in_removed,
                 t.filled_from_removed, t.balance, t.not_a_number]
            )  # fmt: skip
    rows += [
        [],
        ["## Rows removed, by reason"],
        *([reason, n] for reason, n in sorted(removed.items())),
        [],
        ["## Cells changed, by rule"],
        *([rule, n] for rule, n in sorted(rules.items())),
        [],
        ["## Flagged, by issue"],
        *([issue, n] for issue, n in sorted(issues.items())),
        [],
        ["## Placeholder text turned into empty cells"],
        ["sheet", "column", "text", "cells"],
    ]
    for sheet in result.sheets:
        for (column, token), n in sorted(sheet.placeholders.items()):
            rows.append([sheet.name, column, token, n])
    for sheet in result.sheets:
        kept = len(sheet.rows)
        rows += [
            [],
            [f"## Columns in sheet: {sheet.name}"],
            ["column", "type", "type decided by", "missing", "missing %", "changes", "flagged"],
        ]
        for col in sheet.columns:
            pct = round(100 * col.missing / kept, 1) if kept else 0.0
            rows.append(
                [col.name, col.type, col.type_source, col.missing, pct, col.changes, col.issues]
            )
        if sheet.notes:
            rows += [[], [f"## Notes for sheet: {sheet.name}"], *([note] for note in sheet.notes)]
    rows += [[], ["## Run record"], *([k, v] for k, v in result.manifest.items())]
    rows += [[], ["## Config used"], *_flatten(cfg.raw)]
    return rows


def _readme_rows(result: CleanResult, cleaned_titles: list[str]) -> list[list[Any]]:
    validation = any(s.validation for s in result.sheets)
    failed = sum(1 for s in result.sheets for r in s.validation if r.status != "pass")
    lines = [
        ["## Read me first"],
        [f"Source file: {result.source}"],
        [f"{result.rows_in} rows read. Rows in the cleaned data: {result.rows_out}. "
         f"Rows removed: {len(result.removed)}, each listed in full on the Removed sheet."],
        [f"{len(result.changes)} cells were changed. Each one is on the Changes sheet with its "
         "value before and after."],
        [f"{len(result.issues)} items need a person to look at them. They are on the Issues sheet "
         "and were not changed."],
        [],
        ["## What each sheet is"],
        [", ".join(cleaned_titles), "The cleaned data. The source_row column is the row number "
         "in the file you sent."],
        ["Changes", "One line per edit: row, column, before, after, and the rule that made it."],
        ["Removed", "Every row that was taken out, with the reason and the row it duplicated."],
        ["Issues", "Values the tool would not guess at. Start here."],
    ]  # fmt: skip
    if validation:
        lines.append(
            ["Validation", f"Rules you asked for, pass or fail. {failed} rule(s) did not pass."]
        )
    lines += [
        ["Data Dictionary", "One line per column: type, how many are empty, range, examples."],
        ["Summary", "Counts, control totals, notes about the file, and the settings used."],
        [],
        ["## How to check this work"],
        ["On Summary, 'Read - Cleaned - Removed' must be 0 and every control total must balance."],
        ["Pick any line on Changes and compare it with the same row in your original file."],
        ["Your original file was not modified."],
    ]
    return lines


def _dictionary_rows(sheet: SheetResult) -> list[list[Any]]:
    out = []
    rules_by_col: dict[str, set[str]] = {}
    for change in sheet.changes:
        if change.rule != "header":
            rules_by_col.setdefault(change.column, set()).add(change.rule)
    total = len(sheet.rows)
    for col, report in enumerate(sheet.columns):
        values = [row[col] for row in sheet.rows if row[col] is not None]
        numbers = [v for v in values if isinstance(v, int | float) and not isinstance(v, bool)]
        low: float | str | None = None
        high: float | str | None = None
        if numbers and len(numbers) >= len(values) / 2:
            low, high = min(numbers), max(numbers)
        elif report.type == "date":
            # ISO text sorts in date order; values that failed to parse are left out
            dates = [v for v in values if isinstance(v, str) and v[:4].isdigit() and v[4:5] == "-"]
            if dates:
                low, high = min(dates), max(dates)
        distinct = list(dict.fromkeys(values))
        out.append(
            [
                sheet.name,
                report.name,
                report.type,
                report.type_source,
                len(values),
                report.missing,
                round(100 * report.missing / total, 1) if total else 0.0,
                len(distinct),
                low,
                high,
                " | ".join(str(v) for v in distinct[:3]),
                report.changes,
                report.issues,
                ", ".join(sorted(rules_by_col.get(report.name, ()))),
            ]
        )
    return out


def write_workbook(result: CleanResult, path: str | Path, cfg: Config) -> None:
    book = Workbook()
    readme = book.worksheets[0]
    readme.title = "Read me first"
    used: set[str] = {
        "read me first", "changes", "removed", "issues", "validation", "data dictionary", "summary",
    }  # fmt: skip
    single = len(result.sheets) == 1

    titles = []
    for sheet in result.sheets:
        title = "Cleaned" if single else _unique_title(f"Cleaned - {sheet.name}", used)
        titles.append(title)
        headers, rows = sheet.headers, sheet.rows
        offset = 0
        if cfg.source_row_column:
            offset = 1
            headers = ["source_row", *headers]
            rows = [[src, *row] for src, row in zip(sheet.source_rows, rows, strict=True)]
        money = {i + offset for i, col in enumerate(sheet.columns) if col.type == "currency"}
        _write_sheet(book.create_sheet(title), headers, rows, money)
    if not result.sheets:
        titles.append("Cleaned")
        book.create_sheet("Cleaned")
    _write_lines(readme, _readme_rows(result, titles), first_width=30)
    readme.column_dimensions["B"].width = 90

    _write_sheet(
        book.create_sheet("Changes"),
        ["sheet", "row", "column", "before", "after", "rule", "note"],
        [[c.sheet, c.row, c.column, c.before, c.after, c.rule, c.note] for c in result.changes],
    )
    widest: SheetResult | None = max(result.sheets, key=lambda s: len(s.headers), default=None)
    value_headers = (
        result.sheets[0].headers
        if single
        else [f"value_{i + 1}" for i in range(len(widest.headers) if widest else 0)]
    )
    _write_sheet(
        book.create_sheet("Removed"),
        ["sheet", "row", "reason", "duplicate_of_row", "score", *value_headers],
        [
            [r.sheet, r.row, r.reason, r.duplicate_of_row, r.score, *r.values]
            for r in result.removed
        ],
    )
    _write_sheet(
        book.create_sheet("Issues"),
        ["sheet", "row", "column", "value", "issue", "detail", "related_row"],
        [
            [i.sheet, i.row, i.column, i.value, i.issue, i.detail, i.related_row]
            for i in result.issues
        ],
    )
    if any(s.validation for s in result.sheets):
        rows = []
        for sheet in result.sheets:
            for rule in sheet.validation:
                rows.append([sheet.name, rule.rule, rule.column, rule.status, len(rule.failures)])
                rows += [
                    [sheet.name, rule.rule, rule.column, "failing row", None, row, value]
                    for row, value in rule.failures
                ]
        _write_sheet(
            book.create_sheet("Validation"),
            ["sheet", "rule", "column", "status", "failing rows", "row", "value"],
            rows,
        )
    _write_sheet(
        book.create_sheet("Data Dictionary"),
        ["sheet", "column", "type", "type decided by", "filled", "missing", "missing %",
         "distinct values", "min", "max", "examples", "cells changed", "cells flagged",
         "rules applied"],
        [row for sheet in result.sheets for row in _dictionary_rows(sheet)],
    )  # fmt: skip
    _write_lines(book.create_sheet("Summary"), summary_rows(result, cfg))
    book.save(path)
