"""Browser front end for excel-cleaner-pro.

Run with:  streamlit run app/streamlit_app.py

All cleaning is done by the library (clean_file). The helpers below only shape its result for
display, and they import nothing from Streamlit, so tests can call them directly.
"""

from __future__ import annotations

import tempfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from excel_cleaner_pro import CleanResult, SheetResult, clean_file, load_config
from excel_cleaner_pro.config import option_overrides

ACCEPTED_TYPES = ["csv", "xlsx", "xlsm"]
HIGHLIGHT = "background-color: #fde68a; color: #1f2937"


@dataclass
class Preview:
    """Cleaned rows next to the same rows as they were typed, rebuilt from the change log."""

    headers: list[str]
    source_rows: list[int]
    before: list[list[str]]
    after: list[list[str]]
    changed: set[tuple[int, int]] = field(default_factory=set)  # (row index, column index)


def run_clean(filename: str, data: bytes, **options: Any) -> tuple[CleanResult, bytes]:
    """Clean an uploaded file with the library and return the result plus the workbook bytes.

    options are the CLI switches: no_fuzzy, fuzzy_remove, header_row, date_order.
    """
    cfg = load_config(None, option_overrides(**options))
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / Path(filename).name
        src.write_bytes(data)
        out = Path(tmp) / "cleaned.xlsx"
        result = clean_file(src, out, cfg)
        return result, out.read_bytes()


def cleaned_name(filename: str) -> str:
    return f"{Path(filename).stem}_cleaned.xlsx"


def show(value: Any) -> str:
    """One display string per cell, so mixed-type columns render the same everywhere."""
    if value is None:
        return ""
    if isinstance(value, datetime) and value.time() == datetime.min.time():
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def build_preview(sheet: SheetResult, limit: int | None = None) -> Preview:
    """Pair each cleaned row with its original values, using the first logged before-value."""
    col_of = {name: j for j, name in enumerate(sheet.headers)}
    index_of = {src: i for i, src in enumerate(sheet.source_rows)}
    original: dict[tuple[int, int], Any] = {}
    for change in sheet.changes:
        i, j = index_of.get(change.row), col_of.get(change.column)
        if i is None or j is None or change.rule == "header":
            continue
        original.setdefault((i, j), change.before)

    count = len(sheet.rows) if limit is None else min(limit, len(sheet.rows))
    preview = Preview(list(sheet.headers), sheet.source_rows[:count], [], [])
    for i, row in enumerate(sheet.rows[:count]):
        preview.after.append([show(v) for v in row])
        preview.before.append([show(original.get((i, j), v)) for j, v in enumerate(row)])
        preview.changed.update((i, j) for j in range(len(row)) if (i, j) in original)
    return preview


def highlight_mask(preview: Preview) -> list[list[str]]:
    """CSS per cell: highlighted where the cell was changed, empty otherwise."""
    return [
        [HIGHLIGHT if (i, j) in preview.changed else "" for j in range(len(row))]
        for i, row in enumerate(preview.after)
    ]


def changes_rows(result: CleanResult) -> list[dict[str, Any]]:
    return [
        {"sheet": c.sheet, "row": c.row, "column": c.column, "before": show(c.before),
         "after": show(c.after), "rule": c.rule, "note": c.note}
        for c in result.changes
    ]  # fmt: skip


def issues_rows(result: CleanResult) -> list[dict[str, Any]]:
    return [
        {"sheet": i.sheet, "row": i.row, "column": i.column, "value": show(i.value),
         "issue": i.issue, "detail": i.detail, "related_row": i.related_row}
        for i in result.issues
    ]  # fmt: skip


def removed_rows(result: CleanResult) -> list[dict[str, Any]]:
    return [
        {"sheet": r.sheet, "row": r.row, "reason": r.reason,
         "duplicate_of_row": r.duplicate_of_row, "score": r.score,
         "values": " | ".join(show(v) for v in r.values)}
        for r in result.removed
    ]  # fmt: skip


def main() -> None:  # pragma: no cover - UI wiring, exercised by hand in the browser
    import pandas as pd
    import streamlit as st

    st.set_page_config(page_title="Excel Cleaner Pro", layout="wide")
    st.title("Excel Cleaner Pro")
    st.caption("Upload a messy CSV or XLSX file. Every change is shown and logged.")

    with st.sidebar:
        st.header("Options")
        no_fuzzy = st.checkbox("Skip fuzzy duplicate matching")
        fuzzy_remove = st.checkbox("Remove fuzzy duplicates scoring 95 or more")
        date_order = st.selectbox("Dates like 03/04/2024", ["infer", "flag", "mdy", "dmy"])
        limit = st.number_input("Preview rows", min_value=10, max_value=5000, value=200)

    upload = st.file_uploader("File to clean", type=ACCEPTED_TYPES)
    if upload is None:
        st.info("Try it with samples/messy_customers.csv from this repository.")
        return

    try:
        result, workbook = run_clean(
            upload.name, upload.getvalue(),
            no_fuzzy=no_fuzzy, fuzzy_remove=fuzzy_remove, date_order=date_order,
        )  # fmt: skip
    except (ValueError, OSError) as exc:
        st.error(f"Could not clean this file: {exc}")
        return

    cols = st.columns(4)
    cols[0].metric("Rows in", result.rows_in)
    cols[1].metric("Rows out", result.rows_out)
    cols[2].metric("Cells changed", len(result.changes))
    cols[3].metric("Flagged to review", len(result.issues))
    st.download_button(
        "Download cleaned workbook",
        workbook,
        file_name=cleaned_name(upload.name),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    def grid(rows: list[list[str]], preview: Preview) -> Any:
        frame = pd.DataFrame(rows, columns=preview.headers, index=preview.source_rows)
        mask = pd.DataFrame(highlight_mask(preview), columns=frame.columns, index=frame.index)
        return frame.style.apply(lambda _: mask, axis=None)

    for sheet in result.sheets:
        preview = build_preview(sheet, int(limit))
        st.subheader(f"Sheet: {sheet.name}" if len(result.sheets) > 1 else "Before and after")
        st.caption("Changed cells are highlighted. The row label is the row in your file.")
        after_tab, before_tab = st.tabs(["After", "Before"])
        after_tab.dataframe(grid(preview.after, preview))
        before_tab.dataframe(grid(preview.before, preview))

    log, review, gone = st.tabs(
        [f"Changes ({len(result.changes)})", f"Review list ({len(result.issues)})",
         f"Removed rows ({len(result.removed)})"]
    )  # fmt: skip
    log.dataframe(changes_rows(result))
    review.dataframe(issues_rows(result))
    gone.dataframe(removed_rows(result))


if __name__ == "__main__":
    main()
