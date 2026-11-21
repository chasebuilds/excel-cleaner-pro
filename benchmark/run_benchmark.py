"""Run the cleaner on generated data, score it against the known truth, write BENCHMARK.md.

Usage: python benchmark/run_benchmark.py [--rows 5000] [--seed 42]
"""

from __future__ import annotations

import argparse
import platform
import statistics
import sys
import tempfile
import time
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from benchmark.generate import COLUMNS, Record, generate, write_csv  # noqa: E402
from excel_cleaner_pro import CleanResult, clean_file, load_config  # noqa: E402

# error kind -> (what the tool should do, the Changes rule or Issues type that does it)
FIX_KINDS = {
    "whitespace": "whitespace",
    "placeholder": "placeholder",
    "name_case": "name_case",
    "email_case": "email_lowercase",
    "phone_format": "phone_format",
    "date_format": "date_iso",
    "date_inferred_order": "date_iso_inferred_order",
    "currency_format": "number_parse",
    "state_name": "state_code",
}
FLAG_KINDS = {
    "email_invalid": "invalid_email",
    "phone_invalid": "invalid_phone",
    "date_ambiguous": "ambiguous_date",
    "date_unreadable": "unparsed_date",
    "number_ambiguous": "ambiguous_number",
    "state_unknown": "unknown_state",
}
LABELS = {
    "whitespace": "Extra whitespace",
    "placeholder": 'Placeholder text ("N/A", "-", "null")',
    "name_case": "Name in ALL CAPS or lower case",
    "email_case": "Email with capitals",
    "phone_format": "Phone in a non-standard format",
    "date_format": "Date in a non-ISO format, one reading",
    "date_inferred_order": "Date readable two ways, column is all month-first",
    "currency_format": 'Currency text ("$1,200.50", "1.200,50")',
    "state_name": "State name or lower-case code",
    "email_invalid": "Invalid email",
    "phone_invalid": "Invalid phone",
    "date_ambiguous": "Date readable two ways, column mixes orders",
    "date_unreadable": 'Not a real date ("02/30/2024", "Q1 2024")',
    "number_ambiguous": 'Number readable two ways ("1,500" / "1.500")',
    "state_unknown": "Misspelled state",
}


@dataclass
class Score:
    label: str
    expected_action: str
    planted: int
    handled: int  # planted errors that ended up right
    actions: int  # times the tool took this action
    correct_actions: int  # of those, how many were right

    @property
    def recall(self) -> float | None:
        return self.handled / self.planted if self.planted else None

    @property
    def precision(self) -> float | None:
        return self.correct_actions / self.actions if self.actions else None


def same(a: Any, b: Any) -> bool:
    if a is None or b is None:
        return a is None and b is None
    if isinstance(a, int | float) and isinstance(b, int | float):
        return abs(a - b) < 0.005
    return a == b


def score(
    records: list[Record], result: CleanResult, removing: bool = False
) -> tuple[list[Score], dict[str, Any]]:
    """removing says whether the run had fuzzy removal turned on (it is off by default)."""
    sheet = result.sheets[0]
    col = {name: i for i, name in enumerate(COLUMNS)}
    by_row = {i + 2: rec for i, rec in enumerate(records)}

    # the cleaned value of every row, including rows that were later removed as duplicates
    final: dict[int, list[Any]] = dict(zip(sheet.source_rows, sheet.rows, strict=True))
    final.update({r.row: r.values for r in sheet.removed if r.reason != "empty row"})

    planted = Counter(kind for rec in records for kind in rec.planted.values())
    scores: list[Score] = []

    for kind, rule in FIX_KINDS.items():
        handled = sum(
            1
            for row, rec in by_row.items()
            for name, k in rec.planted.items()
            if k == kind and row in final and same(final[row][col[name]], rec.expected[col[name]])
        )
        actions = [c for c in sheet.changes if c.rule == rule]
        correct = sum(
            1
            for c in actions
            if same(final[c.row][col[c.column]], by_row[c.row].expected[col[c.column]])
        )
        scores.append(Score(LABELS[kind], "fix", planted[kind], handled, len(actions), correct))

    flagged = {(i.row, i.column): i.issue for i in sheet.issues}
    for kind, issue in FLAG_KINDS.items():
        handled = sum(
            1
            for row, rec in by_row.items()
            for name, k in rec.planted.items()
            if k == kind
            and flagged.get((row, name)) == issue
            and same(final[row][col[name]], rec.expected[col[name]])
        )
        actions = [i for i in sheet.issues if i.issue == issue]
        correct = sum(1 for i in actions if by_row[i.row].planted.get(i.column) == kind)
        scores.append(
            Score(
                LABELS[kind], "flag, leave as typed", planted[kind], handled, len(actions), correct
            )
        )

    row_kinds = Counter(rec.row_kind for rec in records)
    removed = {r.row: r for r in sheet.removed}

    def right_pair(row: int, other: int | None) -> bool:
        return other in by_row and by_row[row].entity == by_row[other].entity

    def removals(reason: str, kind: str) -> Score:
        acts = [r for r in sheet.removed if r.reason == reason]
        good = sum(1 for r in acts if right_pair(r.row, r.duplicate_of_row))
        hit = sum(
            1
            for row, rec in by_row.items()
            if rec.row_kind == kind
            and row in removed
            and right_pair(row, removed[row].duplicate_of_row)
        )
        return Score("", "remove, log in Removed", row_kinds[kind], hit, len(acts), good)

    exact = removals("exact duplicate", "exact_dup")
    exact.label = "Exact duplicate row"
    fuzzy = removals("fuzzy duplicate", "fuzzy_dup")
    fuzzy.label = "Fuzzy duplicate row, removed"

    reviews = [i for i in sheet.issues if i.issue == "possible_duplicate"]
    review_rows = {i.row for i in reviews if right_pair(i.row, i.related_row)}
    fuzzy_rows = [row for row, rec in by_row.items() if rec.row_kind == "fuzzy_dup"]
    caught = sum(
        1
        for row in fuzzy_rows
        if row in review_rows or (row in removed and right_pair(row, removed[row].duplicate_of_row))
    )
    fuzzy_any = Score(
        "Fuzzy duplicate row, removed or listed for review"
        if removing
        else "Fuzzy duplicate row, listed for review",
        "remove or flag" if removing else "flag, keep the row",
        row_kinds["fuzzy_dup"],
        caught,
        fuzzy.actions + len(reviews),
        fuzzy.correct_actions + len(review_rows),
    )
    empties = [r for r in sheet.removed if r.reason == "empty row"]
    empty = Score(
        "Empty row",
        "remove, log in Removed",
        row_kinds["empty_row"],
        sum(1 for r in empties if by_row[r.row].row_kind == "empty_row"),
        len(empties),
        sum(1 for r in empties if by_row[r.row].row_kind == "empty_row"),
    )
    scores += [exact, fuzzy, fuzzy_any, empty] if removing else [exact, fuzzy_any, empty]

    wrong_cells = sum(
        1
        for row, values in final.items()
        for c, value in enumerate(values)
        if not same(value, by_row[row].expected[c])
    )
    wrong_changes = sum(
        1
        for c in sheet.changes
        if c.rule != "header"
        and not same(final[c.row][col[c.column]], by_row[c.row].expected[col[c.column]])
    )
    other_flags = Counter(
        i.issue for i in sheet.issues if i.issue not in {*FLAG_KINDS.values(), "possible_duplicate"}
    )
    false_removed = [
        r for r in sheet.removed
        if r.reason != "empty row" and not right_pair(r.row, r.duplicate_of_row)
    ]  # fmt: skip
    review_only = sum(1 for row in fuzzy_rows if row in review_rows and row not in removed)
    fuzzy_by_edit = Counter()
    for row in fuzzy_rows:
        edit = by_row[row].fuzzy_edit
        if row in removed and right_pair(row, removed[row].duplicate_of_row):
            fuzzy_by_edit[edit, "removed"] += 1
        elif row in review_rows:
            fuzzy_by_edit[edit, "review"] += 1
        else:
            fuzzy_by_edit[edit, "missed"] += 1
    facts = {
        "rows_in": len(records),
        "rows_out": len(sheet.rows),
        "cells": len(records) * len(COLUMNS),
        "planted_cells": sum(planted.values()),
        "wrong_cells_after": wrong_cells,
        "changes": len(sheet.changes),
        "wrong_changes": wrong_changes,
        "issues": len(sheet.issues),
        "other_flags": dict(other_flags),
        "false_removed": len(false_removed),
        "fuzzy_review_only": review_only,
        "fuzzy_by_edit": dict(fuzzy_by_edit),
        "accounted": len(sheet.rows) + len(sheet.removed) == len(records),
        "totals_balance": all(t.balance == 0 for t in sheet.control_totals),
        "survivor_fills": sum(1 for c in sheet.changes if c.rule == "survivor_fill"),
        "by_label": {sc.label: sc for sc in scores},
        "types": {c.name: f"{c.type} ({c.type_source})" for c in sheet.columns},
    }
    return scores, facts


def _pct(value: float | None) -> str:
    return "n/a" if value is None else f"{100 * value:.1f}%"


REMOVE = {"dedupe": {"fuzzy_action": "remove"}}
NAME = LABELS["name_case"]
REVIEW = "Fuzzy duplicate row, listed for review"
REMOVED = "Fuzzy duplicate row, removed"
EITHER = "Fuzzy duplicate row, removed or listed for review"


def time_run(csv_path: Path, out_path: Path, fuzzy: bool, repeats: int) -> float:
    cfg = load_config(overrides={"dedupe": {"fuzzy": fuzzy}})
    times = []
    for _ in range(repeats):
        start = time.perf_counter()
        clean_file(csv_path, out_path, cfg)
        times.append(time.perf_counter() - start)
    return statistics.median(times)


def run_seed(rows: int, seed: int, tmp: Path) -> dict[str, Any]:
    """Score one seed twice: with the defaults, and with fuzzy removal turned on."""
    records = generate(rows, seed)
    csv_path = tmp / f"messy_{seed}.csv"
    write_csv(records, csv_path)
    default = score(records, clean_file(csv_path))
    removing = score(records, clean_file(csv_path, cfg=load_config(overrides=REMOVE)), True)
    return {"seed": seed, "csv": csv_path, "default": default, "removing": removing}


def _table(scores: list[Score]) -> list[str]:
    lines = [
        "| Error type | Expected | Planted | Handled | Recall | Tool actions | Correct "
        "| Precision |",
        "|---|---|---:|---:|---:|---:|---:|---:|",
    ]
    for s in scores:
        lines.append(
            f"| {s.label} | {s.expected_action} | {s.planted} | {s.handled} | {_pct(s.recall)} "
            f"| {s.actions} | {s.correct_actions} | {_pct(s.precision)} |"
        )
    return lines


def render(main: dict[str, Any], args, timing: dict[str, float], others, big) -> str:
    scores, facts = main["default"]
    r_scores, r_facts = main["removing"]
    n = facts["rows_in"]
    edits = r_facts["fuzzy_by_edit"]
    either = r_facts["by_label"][EITHER]
    lines = [
        "# Benchmark",
        "",
        "Written by `benchmark/run_benchmark.py` from a real run. Do not edit by hand.",
        "",
        f"- Data: `benchmark/generate.py --rows {args.rows} --seed {args.seed}` "
        f"({args.rows} clean customer rows, then damaged: {n} rows and "
        f"{facts['cells']} cells in the messy file)",
        "- Config: defaults, no config file, unless a section says otherwise",
        f"- Python {platform.python_version()} on {platform.system()} {platform.machine()}",
        "",
        "## Result by error type, default settings",
        "",
        "Recall: of the errors planted, how many ended up right. "
        "Precision: of the times the tool took that action, how many were right. "
        '"Right" for a fix means the cell equals the truth table. '
        '"Right" for a flag means the cell was left as typed and listed in `Issues`. '
        "By default a fuzzy duplicate is never removed: it is listed in `Issues` with the "
        "row it resembles.",
        "",
        *_table(scores),
        "",
        "## Whole-file checks, default settings",
        "",
        "| Check | Result |",
        "|---|---:|",
        f"| Cells damaged on purpose | {facts['planted_cells']} |",
        f"| Cells logged in `Changes` | {facts['changes']} |",
        f"| Changes that left a cell different from the truth | {facts['wrong_changes']} |",
        f"| Cells still different from the truth after cleaning | {facts['wrong_cells_after']} |",
        "| Rows removed that were not a planted duplicate or empty row "
        f"| {facts['false_removed']} |",
        "| Every input row is in `Cleaned` or `Removed` "
        f"| {'yes' if facts['accounted'] else 'NO'} |",
        "| Control totals balance on every numeric column "
        f"| {'yes' if facts['totals_balance'] else 'NO'} |",
        f"| Flags of a type the generator did not plant | "
        f"{sum(facts['other_flags'].values())} {facts['other_flags'] or ''} |",
        "",
        "## With fuzzy removal turned on (`--fuzzy-remove`)",
        "",
        "Opt-in. A match scoring 95 or more is removed and the kept row takes the removed "
        "row's value wherever its own cell is empty. False removals and missed duplicates "
        "are counted separately.",
        "",
        *_table([s for s in r_scores if s.label in (REMOVED, EITHER)]),
        "",
        "| Check | Result |",
        "|---|---:|",
        f"| Real customer rows removed by mistake (false removals) | {r_facts['false_removed']} |",
        "| Planted fuzzy duplicates neither removed nor listed (missed) "
        f"| {either.planted - either.handled} |",
        f"| Cells filled into a kept row from a removed row | {r_facts['survivor_fills']} |",
        f"| Cells different from the truth after cleaning | {r_facts['wrong_cells_after']} |",
        "| Control totals balance on every numeric column "
        f"| {'yes' if r_facts['totals_balance'] else 'NO'} |",
        "",
        "| Planted slip | Removed | Listed for review | Missed |",
        "|---|---:|---:|---:|",
    ]
    for edit, label in (
        ("name_typo", "One typo in the name"),
        ("email_typo", "One typo in the email"),
    ):
        lines.append(
            f"| {label} | {edits.get((edit, 'removed'), 0)} | {edits.get((edit, 'review'), 0)} "
            f"| {edits.get((edit, 'missed'), 0)} |"
        )
    lines += [
        "",
        "## Speed",
        "",
        "End to end: read the CSV, clean, write the XLSX. "
        f"Median of {args.repeats} runs, one process.",
        "",
        "| Run | Rows | Seconds | Rows per second |",
        "|---|---:|---:|---:|",
        f"| Defaults (fuzzy duplicate matching on) | {n} | {timing['fuzzy']:.2f} "
        f"| {n / timing['fuzzy']:,.0f} |",
        f"| `--no-fuzzy` | {n} | {timing['plain']:.2f} | {n / timing['plain']:,.0f} |",
        f"| Defaults, larger file (1 run) | {big[0]} | {big[1]:.2f} | {big[0] / big[1]:,.0f} |",
        f"| `--no-fuzzy`, larger file (1 run) | {big[0]} | {big[2]:.2f} | {big[0] / big[2]:,.0f} |",
        "",
        "Fuzzy matching compares each row with the rows kept before it, so its cost grows with "
        "the square of the row count. A rows-per-second figure with fuzzy on holds for that "
        "file size only.",
        "",
        "## Other seeds",
        "",
        "The same test on data the rules were not tuned against. Seeds 1 to 5 were looked at "
        "while building the fuzzy rules, so they are not used here. The result columns up to "
        "'Fuzzy listed, precision' are from default settings, the last two from "
        "`--fuzzy-remove`.",
        "",
        "| Seed | Rows | Wrong changes | Cells wrong after | Name case recall "
        "| Fuzzy listed, recall | Fuzzy listed, precision | Removal on: rows wrongly removed "
        "| Removal on: fuzzy removed, recall |",
        "|---:|---:|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for run in others:
        f, rf = run["default"][1], run["removing"][1]
        by, rby = f["by_label"], rf["by_label"]
        lines.append(
            f"| {run['seed']} | {f['rows_in']} | {f['wrong_changes']} | {f['wrong_cells_after']} "
            f"| {_pct(by[NAME].recall)} | {_pct(by[REVIEW].recall)} "
            f"| {_pct(by[REVIEW].precision)} | {rf['false_removed']} "
            f"| {_pct(rby[REMOVED].recall)} |"
        )
    lines += [
        "",
        "## Column types the tool inferred",
        "",
        "| Column | Type (decided by) |",
        "|---|---|",
        *(f"| {name} | {kind} |" for name, kind in facts["types"].items()),
        "",
        "## What this does not show",
        "",
        "The generator plants the formats listed in `benchmark/generate.py`. A real file can "
        "hold formats that are not on that list, and those are not scored here. "
        "`LIMITATIONS.md` lists the known gaps. The surnames that cannot be rebuilt from "
        "ALL CAPS (MacDonald, de la Cruz, LeBlanc) are in the data on purpose, which is why "
        "the name row is below 100%. Every planted fuzzy duplicate has both key columns "
        "filled: rows with a blank key are not fuzzy-matched at all, so that gap is not "
        "measured here. The data is a CSV, so the checks that only apply to Excel files "
        "(merged cells, hidden rows, formulas with no stored value, IDs stored as numbers) are "
        "covered by the test suite, not by this benchmark. These numbers describe this "
        "generated data. On a client file the same figures have to be measured on that file.",
        "",
        "## Reproduce",
        "",
        "```",
        f"python benchmark/run_benchmark.py --rows {args.rows} --seed {args.seed}",
        "```",
        "",
        "Counts repeat exactly for the same seed. Timings vary by machine.",
        "",
    ]
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=5000)
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--repeats", type=int, default=3)
    parser.add_argument("--other-seeds", type=int, nargs="*", default=[101, 102, 103, 104, 105])
    parser.add_argument("--big-rows", type=int, default=20000, help="size of the larger speed run")
    parser.add_argument("--out", default=str(ROOT / "BENCHMARK.md"))
    args = parser.parse_args()

    with tempfile.TemporaryDirectory() as tmp_name:
        tmp = Path(tmp_name)
        out_path = tmp / "cleaned.xlsx"
        main_run = run_seed(args.rows, args.seed, tmp)
        timing = {
            "fuzzy": time_run(main_run["csv"], out_path, True, args.repeats),
            "plain": time_run(main_run["csv"], out_path, False, args.repeats),
        }
        others = [run_seed(args.rows, seed, tmp) for seed in args.other_seeds]
        big_records = generate(args.big_rows, args.seed)
        big_csv = tmp / "big.csv"
        write_csv(big_records, big_csv)
        big = (
            len(big_records),
            time_run(big_csv, out_path, True, 1),
            time_run(big_csv, out_path, False, 1),
        )
    Path(args.out).write_text(render(main_run, args, timing, others, big), encoding="utf-8")
    print(f"wrote {args.out}")
    for label in ("default", "removing"):
        print(label)
        for s in main_run[label][0]:
            print(f"  {s.label:52s} recall {_pct(s.recall):>6s}  precision {_pct(s.precision):>6s}")


if __name__ == "__main__":
    main()
