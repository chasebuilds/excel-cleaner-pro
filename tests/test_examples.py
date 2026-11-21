import runpy
import subprocess
import sys
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
EXAMPLE = ROOT / "examples" / "use_as_library.py"


def test_library_example_on_the_sample_file(tmp_path, capsys):
    main = runpy.run_path(str(EXAMPLE))["main"]
    out = tmp_path / "sample_cleaned.xlsx"
    summary = main([str(ROOT / "samples" / "messy_customers.csv"), str(out)])

    assert out.exists()
    assert summary["rows_in"] == 42
    assert summary["rows_out"] + summary["removed"] == summary["rows_in"]
    assert summary["changes_by_rule"]["phone_format"] > 0
    assert summary["issues"] > 0
    cleaned = load_workbook(out)["Cleaned"]
    assert cleaned.max_row - 1 == summary["rows_out"]
    phone_col = [c.value for c in cleaned[1]].index("phone") + 1
    phones = [r[0] for r in cleaned.iter_rows(min_row=2, min_col=phone_col, values_only=True)]
    assert "303-555-0155" in phones  # the dashes setting reached the cleaner

    name, email, phone, state = summary["in_memory_rows"][0]
    assert (name, email, phone, state) == ("Ana Lopez", "ana@example.com", "(502) 555-0199", "KY")
    assert summary["in_memory_rows"][1][0] is None  # N/A placeholder blanked

    printed = capsys.readouterr().out
    assert "messy_customers.csv: 42 rows in" in printed
    assert "cells changed by phone_format" in printed


def test_library_example_runs_as_a_script():
    done = subprocess.run(
        [sys.executable, str(EXAMPLE)], capture_output=True, text=True, timeout=120, cwd=ROOT
    )
    assert done.returncode == 0, done.stderr
    assert "rows in" in done.stdout
