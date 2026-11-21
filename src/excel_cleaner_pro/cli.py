"""Command line: excel-cleaner INPUT -o OUTPUT.xlsx"""

from __future__ import annotations

import argparse
import sys
from collections import Counter
from pathlib import Path

from .api import clean_file
from .config import ConfigError, default_config_text, load_config, option_overrides
from .report import write_report_json
from .version import __version__


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-cleaner",
        description="Clean a messy CSV or XLSX file into a workbook with sheets "
        "Cleaned, Changes, Removed, Issues and Summary.",
    )
    parser.add_argument("input", nargs="?", help="CSV or XLSX file to clean")
    parser.add_argument("-o", "--output", help="output .xlsx (default: <input>_cleaned.xlsx)")
    parser.add_argument("-c", "--config", help="TOML config file, see --write-config")
    parser.add_argument("--no-fuzzy", action="store_true", help="skip fuzzy duplicate matching")
    parser.add_argument(
        "--fuzzy-remove",
        action="store_true",
        help="remove fuzzy duplicates at or above the remove threshold instead of only "
        "listing them in Issues; blanks in the kept row are filled from the removed row",
    )
    parser.add_argument(
        "--header-row",
        type=int,
        metavar="N",
        help="row that holds the column names (default: find it)",
    )
    parser.add_argument(
        "--date-order",
        choices=["infer", "flag", "mdy", "dmy"],
        help="how to read dates like 03/04/2024 (default: infer from the column, else flag)",
    )
    parser.add_argument(
        "--report-json",
        metavar="PATH",
        help="also write a JSON summary: changes by rule and column, issues by type, "
        "removed rows by reason, control totals",
    )
    parser.add_argument(
        "--print-config", action="store_true", help="print the default config and exit"
    )
    parser.add_argument(
        "--write-config", metavar="PATH", help="write the default config to PATH and exit"
    )
    parser.add_argument("--version", action="version", version=f"excel-cleaner {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.print_config:
        print(default_config_text(), end="")
        return 0
    if args.write_config:
        if Path(args.write_config).exists():
            print(f"error: {args.write_config} already exists", file=sys.stderr)
            return 2
        Path(args.write_config).write_text(default_config_text(), encoding="utf-8")
        print(f"wrote {args.write_config}")
        return 0
    if not args.input:
        parser.error("an input file is required")

    overrides = option_overrides(
        no_fuzzy=args.no_fuzzy,
        fuzzy_remove=args.fuzzy_remove,
        header_row=args.header_row,
        date_order=args.date_order,
    )
    src = Path(args.input)
    out = Path(args.output) if args.output else src.with_name(f"{src.stem}_cleaned.xlsx")
    try:
        cfg = load_config(args.config, overrides)
        result = clean_file(src, out, cfg)
        if args.report_json:
            write_report_json(result, args.report_json)
    except (ConfigError, ValueError, OSError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    removed = Counter(r.reason for r in result.removed)
    removed_text = ", ".join(f"{n} {reason}" for reason, n in sorted(removed.items())) or "none"
    print(f"{src.name}: {result.rows_in} rows in, {result.rows_out} rows out")
    print(f"  changed cells:   {len(result.changes)}  (sheet Changes)")
    print(f"  removed rows:    {len(result.removed)}  ({removed_text}; sheet Removed)")
    print(f"  flagged to check: {len(result.issues)}  (sheet Issues)")
    failed = sum(1 for s in result.sheets for r in s.validation if r.status != "pass")
    if any(s.validation for s in result.sheets):
        print(f"  validation rules not passed: {failed}  (sheet Validation)")
    print(f"wrote {out}")
    if args.report_json:
        print(f"wrote {args.report_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
