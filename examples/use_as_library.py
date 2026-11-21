"""Use excel-cleaner-pro from Python instead of the command line.

    python examples/use_as_library.py [INPUT] [OUTPUT.xlsx]

With no arguments it cleans samples/messy_customers.csv and writes the workbook to a temporary
folder. It shows three things: cleaning a file with custom settings, reading the audit trail
from the result, and cleaning rows that are already in memory.
"""

from __future__ import annotations

import sys
import tempfile
from collections import Counter
from pathlib import Path

from excel_cleaner_pro import Table, clean_file, clean_tables, load_config

SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "messy_customers.csv"


def clean_with_settings(src: Path, out: Path) -> dict[str, object]:
    # Settings use the same names as the TOML config file (see --print-config).
    cfg = load_config(overrides={"phone": {"format": "dashes"}, "dates": {"order": "infer"}})
    result = clean_file(src, out, cfg)

    print(f"{src.name}: {result.rows_in} rows in, {result.rows_out} rows out, wrote {out.name}")
    by_rule = Counter(c.rule for c in result.changes)
    for rule, n in by_rule.most_common():
        print(f"  {n:4d} cells changed by {rule}")

    first = result.changes[0]
    print(f"  e.g. row {first.row}, {first.column}: {first.before!r} -> {first.after!r}")
    for issue in result.issues[:3]:
        print(f"  to review: row {issue.row}, {issue.column}: {issue.issue} ({issue.value!r})")
    for gone in result.removed:
        print(f"  removed row {gone.row}: {gone.reason}")

    return {
        "rows_in": result.rows_in,
        "rows_out": result.rows_out,
        "changes_by_rule": dict(by_rule),
        "issues": len(result.issues),
        "removed": len(result.removed),
    }


def clean_in_memory() -> list[list[object]]:
    # Rows from a database query or an API can be cleaned without writing a file first.
    table = Table(
        name="crm_export",
        headers=["Name", "Email", "Phone", "State"],
        rows=[
            ["  ANA  LOPEZ ", "Ana@Example.com", "502.555.0199", "kentucky"],
            ["N/A", "not an email", "12345", "Kentuky"],
        ],
    )
    sheet = clean_tables([table], load_config(overrides={"dedupe": {"fuzzy": False}})).sheets[0]
    print("in memory:", sheet.rows)
    return sheet.rows


def main(argv: list[str] | None = None) -> dict[str, object]:
    args = sys.argv[1:] if argv is None else argv
    src = Path(args[0]) if args else SAMPLE
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(args[1]) if len(args) > 1 else Path(tmp) / f"{src.stem}_cleaned.xlsx"
        summary = clean_with_settings(src, out)
        summary["written"] = out.exists()
    summary["in_memory_rows"] = clean_in_memory()
    return summary


if __name__ == "__main__":
    main()
