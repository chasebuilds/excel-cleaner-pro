"""Plain records passed between the reader, the cleaner and the writer."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class ReadIssue:
    """Something the reader saw in the file itself, before any cleaning. col is 0-based."""

    row: int
    col: int | None
    value: Any
    issue: str
    detail: str


@dataclass
class Table:
    """One input sheet. Row numbers are spreadsheet rows; header_row says where the names are."""

    name: str
    headers: list[Any]
    rows: list[list[Any]]
    header_row: int = 1
    above_header: list[tuple[int, list[Any]]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    issues: list[ReadIssue] = field(default_factory=list)
    meta: dict[str, str] = field(default_factory=dict)


@dataclass(frozen=True)
class Change:
    sheet: str
    row: int
    column: str
    before: Any
    after: Any
    rule: str
    note: str = ""


@dataclass(frozen=True)
class Issue:
    sheet: str
    row: int
    column: str
    value: Any
    issue: str
    detail: str = ""
    related_row: int | None = None


@dataclass(frozen=True)
class Removed:
    sheet: str
    row: int
    reason: str
    duplicate_of_row: int | None
    score: float | None
    values: list[Any]


@dataclass
class ColumnReport:
    name: str
    type: str
    type_source: str
    missing: int = 0
    changes: int = 0
    issues: int = 0


@dataclass(frozen=True)
class ControlTotal:
    """Sum of one numeric column at each stage. balance is zero when nothing went astray."""

    column: str
    total_before_removal: float
    in_cleaned: float
    in_removed: float
    filled_from_removed: float
    not_a_number: int

    @property
    def balance(self) -> float:
        moved = self.in_cleaned + self.in_removed - self.filled_from_removed
        return round(moved - self.total_before_removal, 6)


@dataclass(frozen=True)
class RuleResult:
    rule: str
    column: str
    status: str  # pass, fail, column not found
    failures: list[tuple[int, Any]]  # (source row, value)


@dataclass
class SheetResult:
    name: str
    headers: list[str]
    rows: list[list[Any]] = field(default_factory=list)
    source_rows: list[int] = field(default_factory=list)
    rows_in: int = 0
    header_row: int = 1
    changes: list[Change] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    removed: list[Removed] = field(default_factory=list)
    columns: list[ColumnReport] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    meta: dict[str, str] = field(default_factory=dict)
    placeholders: Counter = field(default_factory=Counter)  # (column, token) -> cells blanked
    control_totals: list[ControlTotal] = field(default_factory=list)
    validation: list[RuleResult] = field(default_factory=list)


@dataclass
class CleanResult:
    source: str
    sheets: list[SheetResult]
    manifest: dict[str, Any] = field(default_factory=dict)

    @property
    def changes(self) -> list[Change]:
        return [c for s in self.sheets for c in s.changes]

    @property
    def issues(self) -> list[Issue]:
        return [i for s in self.sheets for i in s.issues]

    @property
    def removed(self) -> list[Removed]:
        return [r for s in self.sheets for r in s.removed]

    @property
    def rows_in(self) -> int:
        return sum(s.rows_in for s in self.sheets)

    @property
    def rows_out(self) -> int:
        return sum(len(s.rows) for s in self.sheets)
