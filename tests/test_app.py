from io import BytesIO
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app.streamlit_app import (
    HIGHLIGHT,
    build_preview,
    changes_rows,
    cleaned_name,
    highlight_mask,
    issues_rows,
    removed_rows,
    run_clean,
)

SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "messy_customers.csv"


def cleaned_sample(**options):
    return run_clean(SAMPLE.name, SAMPLE.read_bytes(), **options)


def test_run_clean_returns_result_and_a_readable_workbook():
    result, workbook = cleaned_sample()
    book = load_workbook(BytesIO(workbook))
    assert {"Cleaned", "Changes", "Removed", "Issues", "Summary"} <= set(book.sheetnames)
    assert book["Cleaned"].max_row - 1 == result.rows_out
    assert book["Changes"].max_row - 1 == len(result.changes)


def test_options_reach_the_library():
    with_fuzzy, _ = cleaned_sample()
    without, _ = cleaned_sample(no_fuzzy=True)
    assert any(i.issue == "possible_duplicate" for i in with_fuzzy.issues)
    assert not any(i.issue == "possible_duplicate" for i in without.issues)


def test_preview_pairs_logged_before_values_with_cleaned_cells():
    result, _ = cleaned_sample()
    sheet = result.sheets[0]
    preview = build_preview(sheet)
    assert len(preview.after) == len(sheet.rows)
    i = preview.source_rows.index(23)
    j = preview.headers.index(
        next(c.column for c in sheet.changes if c.row == 23 and c.rule == "name_case")
    )
    assert preview.before[i][j] == "SOFIA KING"
    assert preview.after[i][j] == "Sofia King"
    assert (i, j) in preview.changed

    cell_changes = {
        (c.row, c.column)
        for c in sheet.changes
        if c.rule != "header" and c.row in sheet.source_rows
    }
    assert len(preview.changed) == len(cell_changes)
    for r, (b_row, a_row) in enumerate(zip(preview.before, preview.after, strict=True)):
        for c, (b, a) in enumerate(zip(b_row, a_row, strict=True)):
            if (r, c) not in preview.changed:
                assert b == a


def test_preview_limit_and_highlight_mask():
    result, _ = cleaned_sample()
    preview = build_preview(result.sheets[0], limit=5)
    assert len(preview.before) == len(preview.after) == len(preview.source_rows) == 5
    mask = highlight_mask(preview)
    marked = {(i, j) for i, row in enumerate(mask) for j, css in enumerate(row) if css}
    assert marked == preview.changed
    assert all(css in ("", HIGHLIGHT) for row in mask for css in row)


def test_tables_match_result_counts():
    result, _ = cleaned_sample()
    assert len(changes_rows(result)) == len(result.changes)
    assert len(issues_rows(result)) == len(result.issues)
    assert len(removed_rows(result)) == len(result.removed)
    assert {"row", "column", "before", "after", "rule"} <= set(changes_rows(result)[0])


def test_cleaned_name():
    assert cleaned_name("customers.csv") == "customers_cleaned.xlsx"
    assert cleaned_name("C:/data/q3 list.xlsx") == "q3 list_cleaned.xlsx"


def test_page_renders_a_cleaned_upload_without_errors():
    testing = pytest.importorskip("streamlit.testing.v1")

    def page():
        from pathlib import Path

        import streamlit as st

        from app import streamlit_app

        sample = Path(streamlit_app.__file__).resolve().parents[1] / "samples"

        class Upload:
            name = "messy_customers.csv"

            def getvalue(self):
                return (sample / self.name).read_bytes()

        st.file_uploader = lambda *a, **k: Upload()
        streamlit_app.main()

    at = testing.AppTest.from_function(page).run(timeout=60)
    assert not at.exception
    result, _ = cleaned_sample()
    expected = [result.rows_in, result.rows_out, len(result.changes), len(result.issues)]
    assert [m.value for m in at.metric] == [str(n) for n in expected]
    assert len(at.dataframe) == 5  # after, before, changes, review list, removed
