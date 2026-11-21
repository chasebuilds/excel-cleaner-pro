"""Duplicate handling: exact rows are removed, fuzzy rows go to review unless removal is on."""

from __future__ import annotations

import re
from typing import Any

from rapidfuzz import fuzz, process

from .config import Config
from .model import Change, Issue, Removed, SheetResult

_PUNCT = re.compile(r"[^\w\s@]+")


def remove_exact(sheet: SheetResult) -> None:
    seen: dict[tuple, int] = {}
    keep_rows, keep_src = [], []
    for row, src in zip(sheet.rows, sheet.source_rows, strict=True):
        key = tuple(row)
        if key in seen:
            sheet.removed.append(Removed(sheet.name, src, "exact duplicate", seen[key], 100.0, row))
        else:
            seen[key] = src
            keep_rows.append(row)
            keep_src.append(src)
    sheet.rows, sheet.source_rows = keep_rows, keep_src


def _norm(value: Any) -> str:
    return "" if value is None else " ".join(_PUNCT.sub(" ", str(value).casefold()).split())


def column_score(a: str, b: str, kind: str) -> float:
    score = fuzz.ratio(a, b)
    if kind == "name":
        # token sort catches "Smith, John" vs "John Smith"; plain ratio catches a typo in
        # the first letter, which would otherwise change the sort order
        score = max(score, fuzz.token_sort_ratio(a, b))
    # digits identify things (jsmith12 vs jsmith47, two phone numbers), so a digit
    # mismatch means different records no matter how alike the letters are
    if re.sub(r"\D", "", a) != re.sub(r"\D", "", b):
        score = min(score, 50.0)
    # same mailbox name at another domain is a different address: worth a look, not a removal
    if kind == "email" and a.partition("@")[2] != b.partition("@")[2]:
        score = min(score, 80.0)
    return score


def resolve_keys(sheet: SheetResult, cfg: Config) -> list[int]:
    if cfg.fuzzy_keys:
        missing = [k for k in cfg.fuzzy_keys if k not in sheet.headers]
        if missing:
            sheet.notes.append(f"fuzzy duplicates skipped: key column(s) not found: {missing}")
            return []
        return [sheet.headers.index(k) for k in cfg.fuzzy_keys]
    keys = [i for i, col in enumerate(sheet.columns) if col.type in ("name", "email")]
    if len(keys) < 2:
        sheet.notes.append(
            "fuzzy duplicates skipped: fewer than two name or email columns found, "
            "set dedupe.fuzzy_keys to choose the columns"
        )
        return []
    return keys


def _fill_survivor(sheet: SheetResult, kept: list[Any], kept_src: int, gone: list[Any],
                   gone_src: int) -> None:  # fmt: skip
    """The kept row takes the removed row's value wherever its own cell is empty."""
    for col, value in enumerate(gone):
        if kept[col] is None and value is not None:
            kept[col] = value
            sheet.changes.append(
                Change(
                    sheet.name, kept_src, sheet.headers[col], None, value, "survivor_fill",
                    f"empty in this row, taken from removed duplicate row {gone_src}",
                )
            )  # fmt: skip


def handle_fuzzy(sheet: SheetResult, cfg: Config) -> None:
    keys = resolve_keys(sheet, cfg)
    if not keys:
        return
    removing = cfg.fuzzy_action == "remove"
    kinds = [sheet.columns[i].type for i in keys]
    is_name = [kind == "name" for kind in kinds]
    label = " + ".join(sheet.headers[i] for i in keys)
    action = (
        f"removed at {cfg.remove_threshold:g}, listed in Issues from {cfg.review_threshold:g}"
        if removing
        else f"listed in Issues from {cfg.review_threshold:g}, never removed"
    )
    sheet.notes.append(f"fuzzy duplicates compared on: {label} ({action})")

    kept_rows, kept_src, kept_norm = [], [], []
    # one list per key column, plus a word-sorted copy for name columns
    kept_by_key: list[list[str]] = [[] for _ in keys]
    kept_sorted: list[list[str]] = [[] for _ in keys]
    skipped = 0
    for row, src in zip(sheet.rows, sheet.source_rows, strict=True):
        norm = [_norm(row[i]) for i in keys]
        best_score, best_idx = 0.0, -1
        # A row with a blank key has less evidence than the thresholds assume (two people
        # can share a name), so it is never fuzzy-matched. Exact matching still covers it.
        complete = all(norm)
        skipped += not complete
        if complete and kept_rows:
            # The row score is an average, so it can only reach the review threshold if at
            # least one column does. Searching each column at that cutoff misses no pair.
            candidates: set[int] = set()
            for k, text in enumerate(norm):
                searches = [(text, kept_by_key[k])]
                if is_name[k]:
                    searches.append((" ".join(sorted(text.split())), kept_sorted[k]))
                for query, choices in searches:
                    hits = process.extract(
                        query, choices, scorer=fuzz.ratio, score_cutoff=cfg.review_threshold,
                        limit=None,
                    )  # fmt: skip
                    candidates.update(idx for _, _, idx in hits)
            for idx in sorted(candidates):
                other = kept_norm[idx]
                if not all(other):
                    continue
                score = sum(
                    column_score(a, b, kind) for a, b, kind in zip(norm, other, kinds, strict=True)
                ) / len(keys)
                if score > best_score:
                    best_score, best_idx = score, idx

        if removing and best_score >= cfg.remove_threshold:
            _fill_survivor(sheet, kept_rows[best_idx], kept_src[best_idx], row, src)
            sheet.removed.append(
                Removed(
                    sheet.name,
                    src,
                    "fuzzy duplicate",
                    kept_src[best_idx],
                    round(best_score, 1),
                    row,
                )
            )
            continue
        if best_score >= cfg.review_threshold:
            other_row = kept_rows[best_idx]
            shown, other_shown = (
                " | ".join("" if r[i] is None else str(r[i]) for i in keys)
                for r in (row, other_row)
            )
            sheet.issues.append(
                Issue(
                    sheet.name,
                    src,
                    label,
                    shown,
                    "possible_duplicate",
                    f"similar to row {kept_src[best_idx]} (score {best_score:.1f}): {other_shown}",
                    kept_src[best_idx],
                )
            )
        kept_rows.append(row)
        kept_src.append(src)
        kept_norm.append(norm)
        for k, text in enumerate(norm):
            kept_by_key[k].append(text)
            kept_sorted[k].append(" ".join(sorted(text.split())) if is_name[k] else "")
    if skipped:
        sheet.notes.append(
            f"fuzzy duplicates: {skipped} row(s) with a blank key column were not compared"
        )
    sheet.rows, sheet.source_rows = kept_rows, kept_src
