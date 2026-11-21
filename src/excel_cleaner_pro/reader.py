"""Read CSV and XLSX into tables, and report what the file itself hides from a script.

Nothing here repairs anything. Merged cells, hidden rows, formulas with no stored value
and rows above the header are reported so a person can decide.
"""

from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.datetime import CALENDAR_MAC_1904
from openpyxl.worksheet.worksheet import Worksheet

from .config import Config
from .model import ReadIssue, Table

_DELIMITER_NAMES = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}


def _filled(row: list[Any]) -> list[Any]:
    return [v for v in row if v is not None and str(v).strip() != ""]


def locate_header(rows: list[list[Any]], cfg: Config) -> int:
    """0-based index of the header row: the first row that is at least half filled with text."""
    if cfg.header_row:
        if cfg.header_row > len(rows):
            raise ValueError(
                f"input.header_row is {cfg.header_row} but the sheet has {len(rows)} rows"
            )
        return cfg.header_row - 1
    width = max((len(_filled(r)) for r in rows[:200]), default=0)
    need = max(1, math.ceil(width / 2))
    for i, row in enumerate(rows[:25]):
        cells = _filled(row)
        if len(cells) >= need and all(isinstance(v, str) for v in cells):
            return i
    return 0


def build_table(name: str, rows: list[list[Any]], cfg: Config) -> Table:
    if not rows:
        return Table(name, [], [])
    at = locate_header(rows, cfg)
    table = Table(name, rows[at], rows[at + 1 :], header_row=at + 1)
    table.above_header = [(i + 1, rows[i]) for i in range(at)]
    if at:
        how = (
            "set in the config"
            if cfg.header_row
            else "the first row at least half filled with text"
        )
        table.notes.append(
            f"header row is row {at + 1} ({how}); the {at} row(s) above it are in the Removed "
            "sheet with the reason 'above header row'"
        )
        grouped = [i + 1 for i in range(at) if len(_filled(rows[i])) >= 2]
        if grouped:
            table.notes.append(
                f"possible multi-row header: row(s) {grouped} above the header hold more than one "
                "label; they were not merged into the column names"
            )
    return table


def read_csv(path: Path, cfg: Config) -> list[Table]:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8-sig")
        encoding = "UTF-8 with byte order mark" if raw.startswith(b"\xef\xbb\xbf") else "UTF-8"
    except UnicodeDecodeError:
        # Excel on Windows saves plain "CSV" in the legacy code page. Any byte string
        # decodes as cp1252-or-latin-1, so this is a guess and the Summary says so.
        try:
            text = raw.decode("cp1252")
            encoding = "Windows-1252 (guess: the file is not valid UTF-8)"
        except UnicodeDecodeError:
            text = raw.decode("latin-1")
            encoding = "Latin-1 (guess: the file is not valid UTF-8 or Windows-1252)"
    try:
        dialect = csv.Sniffer().sniff(text[:8192], delimiters=",;\t|")
        delimiter = dialect.delimiter
    except csv.Error:
        dialect, delimiter = csv.excel, ","
    rows = list(csv.reader(text.splitlines(keepends=True), dialect))
    table = build_table(path.stem, rows, cfg)
    table.meta = {
        "Encoding assumed": encoding,
        "Delimiter assumed": _DELIMITER_NAMES.get(delimiter, repr(delimiter)),
    }
    return [table]


def _inspect(ws: Worksheet, ws_formulas: Worksheet, table: Table) -> None:
    """Add notes and issues for structure that a values-only read would hide."""
    if ws.sheet_state != "visible":
        table.notes.append("this sheet is hidden in the workbook; it was read like any other")

    for rng in ws.merged_cells.ranges:
        cells = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
        value = ws.cell(rng.min_row, rng.min_col).value
        what = (
            f"merged range {rng.coord}: Excel keeps the value in the first cell only, the other "
            f"{cells - 1} cell(s) read as empty and were not filled"
        )
        if rng.min_row <= table.header_row:
            table.notes.append(f"header area has {what}")
        else:
            table.issues.append(
                ReadIssue(rng.min_row, rng.min_col - 1, value, "merged_cells", what)
            )

    hidden_rows = sorted(r for r, dim in ws.row_dimensions.items() if dim.hidden)
    if hidden_rows:
        table.notes.append(
            f"{len(hidden_rows)} hidden row(s) were read like any other (first: {hidden_rows[:10]})"
        )
    hidden_cols = sorted(c for c, dim in ws.column_dimensions.items() if dim.hidden)
    if hidden_cols:
        table.notes.append(f"hidden column(s) were read like any other: {hidden_cols}")

    formulas = 0
    for row in ws_formulas.iter_rows():
        for cell in row:
            if cell.data_type != "f" or cell.row is None or cell.column is None:
                continue
            formulas += 1
            if ws.cell(cell.row, cell.column).value is None:
                table.issues.append(
                    ReadIssue(
                        cell.row,
                        cell.column - 1,
                        None,
                        "formula_no_value",
                        f"formula {str(cell.value)[:80]} has no stored result in the file, so the "
                        "cell reads as empty. Open and save the file in Excel, then run again.",
                    )
                )
    if formulas:
        table.notes.append(
            f"{formulas} formula cell(s) were read as the last value Excel calculated, "
            "not recalculated"
        )


def read_xlsx(path: Path, cfg: Config) -> list[Table]:
    # two loads: one for the values Excel last calculated, one for the formula text
    book = load_workbook(path, data_only=True)
    formulas = load_workbook(path, data_only=False)
    tables, empty = [], []
    for ws in book.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        while rows and all(v is None for v in rows[-1]):
            rows.pop()
        if not rows:
            empty.append(ws.title)
            continue
        table = build_table(ws.title, rows, cfg)
        _inspect(ws, formulas[ws.title], table)
        tables.append(table)
    meta = {
        "Date system": "1904 (Mac)" if book.epoch == CALENDAR_MAC_1904 else "1900",
        "Sheets read": ", ".join(t.name for t in tables) or "none",
    }
    if empty:
        meta["Sheets skipped (empty)"] = ", ".join(empty)
    for table in tables:
        table.meta = meta
    return tables


def read_tables(path: str | Path, cfg: Config) -> list[Table]:
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in (".csv", ".tsv", ".txt"):
        return read_csv(path, cfg)
    if suffix in (".xlsx", ".xlsm"):
        return read_xlsx(path, cfg)
    raise ValueError(f"unsupported file type {suffix!r}: use .csv or .xlsx")
