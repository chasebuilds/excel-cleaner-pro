"""Declared rules checked against the cleaned data. A failed rule reports; it never edits."""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Callable
from typing import Any

from .config import Config
from .model import RuleResult, SheetResult


def _check(sheet: SheetResult, rule: str, column: str, bad: Callable[..., bool]) -> RuleResult:
    if column not in sheet.headers:
        return RuleResult(rule, column, "column not found", [])
    col = sheet.headers.index(column)
    values = [row[col] for row in sheet.rows]
    failures = [
        (src, value)
        for src, value in zip(sheet.source_rows, values, strict=True)
        if bad(value, values)
    ]
    return RuleResult(rule, column, "fail" if failures else "pass", failures)


def run_rules(sheet: SheetResult, cfg: Config) -> list[RuleResult]:
    rules = cfg.validate
    results: list[RuleResult] = []

    for column in rules["unique"]:
        counts: Counter[Any] = Counter()
        if column in sheet.headers:
            col = sheet.headers.index(column)
            counts.update(row[col] for row in sheet.rows if row[col] is not None)
        results.append(
            _check(sheet, "unique", column, lambda v, _, c=counts: v is not None and c[v] > 1)
        )
    for column in rules["required"]:
        results.append(_check(sheet, "required", column, lambda v, _: v is None))
    for column, allowed in rules["allowed"].items():
        ok = set(allowed)
        results.append(
            _check(
                sheet, f"allowed: {sorted(map(str, ok))}", column,
                lambda v, _, ok=ok: v is not None and v not in ok,
            )
        )  # fmt: skip
    for column, (low, high) in rules["range"].items():

        def outside(value: Any, _: Any, low: Any = low, high: Any = high) -> bool:
            if value is None:
                return False
            # text in a numeric column is out of range by definition: it cannot be compared
            if isinstance(value, bool) or not isinstance(value, type(low) | int | float):
                return True
            try:
                return not low <= value <= high
            except TypeError:
                return True

        results.append(_check(sheet, f"range: {low} to {high}", column, outside))
    for column, pattern in rules["pattern"].items():
        regex = re.compile(pattern)
        results.append(
            _check(
                sheet, f"pattern: {pattern}", column,
                lambda v, _, regex=regex: v is not None and not regex.fullmatch(str(v)),
            )
        )  # fmt: skip
    return results
