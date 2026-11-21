"""Machine-readable run summary: the counts from the Summary sheet as JSON."""

from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from .model import Change, CleanResult, SheetResult
from .version import __version__

REPORT_FORMAT = 1


def _by_rule(changes: list[Change]) -> dict[str, dict[str, Any]]:
    columns: defaultdict[str, Counter[str]] = defaultdict(Counter)
    for change in changes:
        columns[change.rule][change.column] += 1
    return {
        rule: {"cells": sum(cols.values()), "columns": dict(cols.most_common())}
        for rule, cols in sorted(columns.items(), key=lambda kv: -sum(kv[1].values()))
    }


def _counts(values: list[str]) -> dict[str, int]:
    return dict(Counter(values).most_common())


def _sheet(sheet: SheetResult) -> dict[str, Any]:
    return {
        "name": sheet.name,
        "rows_in": sheet.rows_in,
        "rows_out": len(sheet.rows),
        "rows_removed": len(sheet.removed),
        "changes_by_rule": _by_rule(sheet.changes),
        "issues_by_type": _counts([i.issue for i in sheet.issues]),
        "removed_by_reason": _counts([r.reason for r in sheet.removed]),
        "column_types": {c.name: c.type for c in sheet.columns},
        "control_totals": [
            {
                "column": t.column,
                "total_before_removal": t.total_before_removal,
                "in_cleaned": t.in_cleaned,
                "in_removed": t.in_removed,
                "filled_from_removed": t.filled_from_removed,
                "balance": t.balance,
            }
            for t in sheet.control_totals
        ],
        "validation": [
            {"rule": v.rule, "column": v.column, "status": v.status, "failures": len(v.failures)}
            for v in sheet.validation
        ],
    }


def build_report(result: CleanResult) -> dict[str, Any]:
    """Counts by rule, issue type and removal reason, for the whole run and per sheet."""
    return {
        "report_format": REPORT_FORMAT,
        "tool": f"excel-cleaner-pro {__version__}",
        "source": result.source,
        "input_sha256": result.manifest.get("Input SHA-256"),
        "rows_in": result.rows_in,
        "rows_out": result.rows_out,
        "rows_removed": len(result.removed),
        "rows_unaccounted": result.rows_in - result.rows_out - len(result.removed),
        "cells_changed": len(result.changes),
        "issues": len(result.issues),
        "changes_by_rule": _by_rule(result.changes),
        "issues_by_type": _counts([i.issue for i in result.issues]),
        "removed_by_reason": _counts([r.reason for r in result.removed]),
        "sheets": [_sheet(s) for s in result.sheets],
    }


def write_report_json(result: CleanResult, path: str | Path) -> None:
    text = json.dumps(build_report(result), indent=2, ensure_ascii=False, default=str)
    Path(path).write_text(text + "\n", encoding="utf-8")
