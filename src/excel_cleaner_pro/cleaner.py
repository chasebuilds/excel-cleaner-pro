"""The cleaning pipeline for one table. Every edit lands in changes, issues or removed."""

from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from datetime import date, datetime
from typing import Any

from . import rules
from .config import Config, load_config
from .dedupe import handle_fuzzy, remove_exact
from .infer import infer_type
from .model import Change, ColumnReport, ControlTotal, Issue, Removed, SheetResult, Table
from .validate import run_rules

# UTF-8 bytes shown through a Western code page: A-tilde plus a symbol where an accented
# letter should be, or a-circumflex plus a euro sign where a curly quote should be
_MOJIBAKE = re.compile("Ã[\u0080-¿]|â€|Â[ -¿]|�")
_TOTAL_LABEL = re.compile(r"^(grand |sub)?totals?:?$|^sum:?$", re.IGNORECASE)
_SCIENTIFIC = re.compile(r"^\d(\.\d+)?E\+?\d+$", re.IGNORECASE)
_MONTH = "jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec"
_DATE_LIKE_CODE = re.compile(rf"^(\d{{1,2}}-({_MONTH})|({_MONTH})-\d{{2,4}})$", re.IGNORECASE)
_SIGNED_NUMBER = re.compile(r"^[+-][\d\s().,-]+$")
_ZIP_HEADER = re.compile(r"(zip|postal)", re.IGNORECASE)


def _clean_headers(table: Table, sheet: SheetResult, width: int) -> list[str]:
    headers: list[str] = []
    seen: Counter[str] = Counter()
    for i in range(width):
        raw = table.headers[i] if i < len(table.headers) else None
        name = rules.clean_whitespace(str(raw)) if raw is not None else ""
        if not name:
            name = f"column_{i + 1}"
        seen[name] += 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        if name != raw:
            sheet.changes.append(Change(sheet.name, table.header_row, name, raw, name, "header"))
        headers.append(name)
    return headers


def _base_clean(value: Any, cfg: Config) -> tuple[Any, list[tuple[Any, Any, str]]]:
    """Unicode form, whitespace and placeholders. Returns the value and each step taken."""
    if not isinstance(value, str):
        return value, []
    steps: list[tuple[Any, Any, str]] = []
    text = value
    if cfg.rules["unicode_nfc"] and not unicodedata.is_normalized("NFC", text):
        # NFC only joins a letter with its accent. NFKC would also rewrite characters
        # such as a ligature in a name or a superscript in a unit, so it is not used.
        fixed = unicodedata.normalize("NFC", text)
        steps.append((text, fixed, "unicode_nfc"))
        text = fixed
    if cfg.rules["whitespace"]:
        fixed = rules.clean_whitespace(text)
        if fixed != text:
            steps.append((text, fixed or None, "whitespace"))
            text = fixed
    if text == "":
        return None, steps
    if cfg.rules["placeholders"] and text.casefold() in cfg.placeholder_tokens:
        steps.append((text, None, "placeholder"))
        return None, steps
    return text, steps


def _native_date(value: Any) -> Any:
    """A real Excel date becomes ISO text. Same day, so it is not logged as a change."""
    if isinstance(value, datetime):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.date().isoformat()
        return value.replace(microsecond=0).isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def _is_number(value: Any) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool)


class _ColumnCleaner:
    def __init__(self, sheet: SheetResult, cfg: Config, col: int, report: ColumnReport):
        self.sheet, self.cfg, self.col, self.report = sheet, cfg, col, report

    def change(self, i: int, after: Any, rule: str) -> None:
        row = self.sheet.rows[i]
        self.sheet.changes.append(
            Change(self.sheet.name, self.sheet.source_rows[i], self.report.name, row[self.col],
                   after, rule)
        )  # fmt: skip
        row[self.col] = after

    def issue(self, i: int, issue: str, detail: str = "") -> None:
        self.sheet.issues.append(
            Issue(self.sheet.name, self.sheet.source_rows[i], self.report.name,
                  self.sheet.rows[i][self.col], issue, detail)
        )  # fmt: skip

    def column_issue(self, hits: list[int], issue: str, why: str) -> None:
        """One Issues line for a fault that affects many cells of the column the same way."""
        if not hits:
            return
        rows = ", ".join(str(self.sheet.source_rows[i]) for i in hits[:5])
        more = ", ..." if len(hits) > 5 else ""
        self.issue(
            hits[0], issue, f"{len(hits)} value(s) in this column (rows {rows}{more}). {why}"
        )

    def cells(self):
        for i, row in enumerate(self.sheet.rows):
            if row[self.col] is not None:
                yield i, row[self.col]

    def scan_text(self) -> None:
        for i, value in self.cells():
            if not isinstance(value, str):
                continue
            if _MOJIBAKE.search(value):
                self.issue(
                    i, "possible_mojibake",
                    "looks like text saved in one encoding and read in another; not repaired",
                )  # fmt: skip
            elif _TOTAL_LABEL.match(value):
                self.issue(
                    i, "possible_total_row",
                    "this row looks like a total line, not a record; it was kept",
                )  # fmt: skip

    # ---- one method per column type

    def name(self) -> None:
        for i, value in self.cells():
            if isinstance(value, str) and (fixed := rules.clean_name(value)) != value:
                self.change(i, fixed, "name_case")

    def email(self) -> None:
        for i, value in self.cells():
            text = str(value)
            if not rules.is_valid_email(text):
                self.issue(i, "invalid_email", "left as typed")
            elif text != text.lower():
                self.change(i, text.lower(), "email_lowercase")

    def phone(self) -> None:
        for i, value in self.cells():
            # Excel stores an unformatted phone number as a float
            text = (
                str(int(value)) if isinstance(value, float) and value.is_integer() else str(value)
            )
            result = rules.clean_phone(text, self.cfg.phone_format)
            if result.issue:
                self.issue(i, result.issue, result.detail)
            elif result.value != value:
                self.change(i, result.value, "phone_format")

    def state(self) -> None:
        for i, value in self.cells():
            result = rules.clean_state(str(value))
            if result.value is None:
                self.issue(i, "unknown_state", result.detail)
            elif result.value != value:
                self.change(i, result.value, "state_code")

    def date(self) -> None:
        parsed: dict[int, rules.DateResult] = {}
        for i, value in self.cells():
            if isinstance(value, str):
                parsed[i] = rules.parse_date(value, self.cfg.two_digit_year_pivot)
            else:
                self.sheet.rows[i][self.col] = _native_date(value)

        proof = Counter(p.evidence for p in parsed.values() if p.evidence)
        order = self.cfg.date_order
        if order == "infer":
            need = self.cfg.date_infer_min_evidence
            if proof["mdy"] >= need and not proof["dmy"]:
                order = "mdy"
            elif proof["dmy"] >= need and not proof["mdy"]:
                order = "dmy"
            else:
                order = "flag"
            inferred = order != "flag"
        else:
            inferred = False

        ambiguous = sum(1 for p in parsed.values() if p.kind == "ambiguous")
        if ambiguous and inferred:
            first = "month" if order == "mdy" else "day"
            self.sheet.notes.append(
                f"{self.report.name}: {ambiguous} date(s) that could be read two ways were read "
                f"{first}-first, because {proof[order]} other value(s) in the column can only be "
                f"{first}-first and none can only be the other way (rule date_iso_inferred_order)"
            )
        seen = f"column has {proof['mdy']} month-first and {proof['dmy']} day-first proofs"

        for i, p in parsed.items():
            value = self.sheet.rows[i][self.col]
            if p.kind == "bad":
                self.issue(i, "unparsed_date", "not a date format this tool reads")
            elif p.kind == "ambiguous" and order == "flag":
                self.issue(
                    i,
                    "ambiguous_date",
                    f"could be {p.mdy} (month first) or {p.dmy} (day first); {seen}",
                )
            else:
                iso = p.iso or (p.mdy if order == "mdy" else p.dmy)
                if iso == value:
                    continue
                if p.kind == "ambiguous" and inferred:
                    rule = "date_iso_inferred_order"
                elif p.two_digit_year:
                    rule = "date_iso_2digit_year"
                else:
                    rule = "date_iso"
                self.change(i, iso, rule)

    def number(self) -> None:
        as_currency = self.report.type == "currency"
        parsed = {i: rules.parse_number(v) for i, v in self.cells() if isinstance(v, str)}
        proof = Counter(p.evidence for p in parsed.values() if p.evidence)
        # both separator styles in one column: no single reading can be trusted
        mixed = proof["us"] and proof["eu"]
        reading = None if mixed else ("eu" if proof["eu"] else "us")

        for i, p in parsed.items():
            if p.kind == "bad":
                self.issue(i, "unparsed_number", "not a number format this tool reads")
                continue
            if p.kind == "ambiguous" and reading is None:
                self.issue(
                    i,
                    "ambiguous_number",
                    f"could be {p.us} or {p.eu}; column mixes {proof['us']} dot-decimal and "
                    f"{proof['eu']} comma-decimal values",
                )
                continue
            value = p.value if p.kind == "ok" else (p.us if reading == "us" else p.eu)
            if as_currency:
                value = round(float(value), 2)
            # a bare literal such as 1250.00 only changes type, not meaning, so it is not logged
            if p.plain and not (p.kind == "ambiguous" and reading == "eu"):
                self.sheet.rows[i][self.col] = value
            else:
                self.change(i, value, "number_parse")

    currency = number

    def excel_damage(self) -> None:
        """Marks Excel leaves on a code after it has already changed it. Flag only."""
        cells = list(self.cells())
        scientific = [i for i, v in cells if isinstance(v, str) and _SCIENTIFIC.match(v)]
        dated = [
            i for i, v in cells
            if isinstance(v, date | datetime) or (isinstance(v, str) and _DATE_LIKE_CODE.match(v))
        ]  # fmt: skip
        # in free text a lone match proves nothing, so text columns need a clear minority
        if self.report.type == "text" and len(scientific) + len(dated) > 0.2 * len(cells):
            return
        self.column_issue(
            scientific, "scientific_notation",
            "Excel shows long numbers this way and the hidden digits may be gone. "
            "Check against the source system.",
        )  # fmt: skip
        self.column_issue(
            dated, "possible_date_conversion",
            "Excel turns a code such as SEPT2 or 2-9 into a date on entry. "
            "The original text cannot be recovered from this file.",
        )  # fmt: skip

    text = excel_damage

    def id(self) -> None:
        cells = list(self.cells())
        self.column_issue(
            [i for i, v in cells if _is_number(v)], "id_stored_as_number",
            "The file holds these as numbers, not text, so any leading zeros were lost before "
            "this tool read them.",
        )  # fmt: skip
        digits = {
            i: str(int(v)) if _is_number(v) and float(v).is_integer() else str(v) for i, v in cells
        }
        if _ZIP_HEADER.search(self.report.name):
            self.column_issue(
                [i for i, d in digits.items() if d.isdigit() and len(d) in (3, 4)],
                "possible_lost_leading_zero",
                "A US ZIP code has 5 digits. Not padded, because the tool cannot tell a lost "
                "zero from a wrong value.",
            )
        # A number of 16 or more digits cannot survive Excel. Long codes kept as text can,
        # so text is flagged only when every long value ends in 0, the mark truncation leaves.
        long_numbers = [i for i, v in cells if _is_number(v) and abs(v) >= 1e15]
        long_text = [i for i, v in cells if isinstance(v, str) and v.isdigit() and len(v) >= 16]
        if len(long_text) >= 2 and all(
            self.sheet.rows[i][self.col].endswith("0") for i in long_text
        ):
            long_numbers += long_text
        self.column_issue(
            sorted(long_numbers),
            "possible_truncated_digits",
            "Excel keeps 15 digits of a number and writes zeros after that.",
        )
        self.excel_damage()


def _control_totals(sheet: SheetResult, before: dict[int, float]) -> list[ControlTotal]:
    totals = []
    for col, start in before.items():
        name = sheet.headers[col]
        kept = [row[col] for row in sheet.rows]
        gone = [
            r.values[col]
            for r in sheet.removed
            if r.duplicate_of_row is not None and _is_number(r.values[col])
        ]
        filled = [
            c.after
            for c in sheet.changes
            if c.rule == "survivor_fill" and c.column == name and _is_number(c.after)
        ]
        totals.append(
            ControlTotal(
                name,
                round(start, 6),
                round(math.fsum(v for v in kept if _is_number(v)), 6),
                round(math.fsum(gone), 6),
                round(math.fsum(filled), 6),
                sum(1 for v in kept if v is not None and not _is_number(v)),
            )
        )
    return totals


def clean_table(table: Table, cfg: Config | None = None) -> SheetResult:
    cfg = cfg or load_config()
    width = max([len(table.headers), *(len(r) for r in table.rows)], default=0)
    sheet = SheetResult(
        table.name,
        [],
        rows_in=len(table.rows) + len(table.above_header),
        header_row=table.header_row,
        notes=list(table.notes),
        meta=dict(table.meta),
    )
    sheet.headers = _clean_headers(table, sheet, width)
    for row_number, values in table.above_header:
        sheet.removed.append(
            Removed(sheet.name, row_number, "above header row", None, None, values)
        )

    for offset, raw in enumerate(table.rows):
        src = table.header_row + 1 + offset
        padded = list(raw) + [None] * (width - len(raw))
        cleaned = [_base_clean(v, cfg) for v in padded]
        if all(v is None for v, _ in cleaned):
            sheet.removed.append(Removed(sheet.name, src, "empty row", None, None, padded))
            continue
        for header, (_, steps) in zip(sheet.headers, cleaned, strict=True):
            for before, after, rule in steps:
                sheet.changes.append(Change(sheet.name, src, header, before, after, rule))
                if rule == "placeholder":
                    sheet.placeholders[header, before.casefold()] += 1
        sheet.rows.append([v for v, _ in cleaned])
        sheet.source_rows.append(src)

    for found in table.issues:
        column = sheet.headers[found.col] if found.col is not None and found.col < width else ""
        sheet.issues.append(
            Issue(sheet.name, found.row, column, found.value, found.issue, found.detail)
        )

    for col, header in enumerate(sheet.headers):
        col_type, source = infer_type(header, [r[col] for r in sheet.rows], cfg)
        report = ColumnReport(header, col_type, source)
        sheet.columns.append(report)
        cleaner = _ColumnCleaner(sheet, cfg, col, report)
        cleaner.scan_text()
        rule_switch = {"name": "name_case", "date": "dates", "number": "numbers",
                       "currency": "numbers", "state": "states"}  # fmt: skip
        rule_switch = rule_switch.get(col_type, col_type)
        if col_type in ("text", "id") or cfg.rules.get(rule_switch, False):
            getattr(cleaner, col_type)()

    numeric = [i for i, c in enumerate(sheet.columns) if c.type in ("number", "currency")]
    before = {i: math.fsum(r[i] for r in sheet.rows if _is_number(r[i])) for i in numeric}

    if cfg.dedupe_exact:
        remove_exact(sheet)
    if cfg.dedupe_fuzzy:
        handle_fuzzy(sheet, cfg)

    sheet.control_totals = _control_totals(sheet, before)
    risky = sum(
        1
        for row in sheet.rows
        for v in row
        if isinstance(v, str) and v[0] in "=+-@" and not _SIGNED_NUMBER.match(v)
    )
    if risky:
        sheet.notes.append(
            f"{risky} text value(s) start with =, +, - or @. They are stored as text in this "
            "workbook. Saved as CSV and reopened, a spreadsheet program may run them as formulas."
        )
    sheet.validation = run_rules(sheet, cfg)

    changed = Counter(c.column for c in sheet.changes if c.rule != "header")
    flagged = Counter(i.column for i in sheet.issues)
    for col, report in enumerate(sheet.columns):
        report.missing = sum(1 for r in sheet.rows if r[col] is None)
        report.changes = changed[report.name]
        report.issues = flagged[report.name]
    sheet.changes.sort(key=lambda c: c.row)
    sheet.issues.sort(key=lambda i: i.row)
    sheet.removed.sort(key=lambda r: r.row)
    return sheet
