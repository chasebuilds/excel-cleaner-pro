from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

from excel_cleaner_pro import clean_file, load_config
from excel_cleaner_pro.cli import main

CSV = (
    "Name,Email,Phone,State,Joined,Amount\n"
    ' JOHN  SMITH ,John@Example.COM,555.201.0143,kentucky,3/15/2024,"$1,200.50"\n'
    'mary jones,bad@@x,12345,Kentuky,03/04/2024,"1.200,50"\n'
    "John Smith,john@example.com,(555) 201-0143,KY,2024-03-15,1200.50\n"
    "N/A,-,null,,,\n"
)


@pytest.fixture
def csv_path(tmp_path):
    path = tmp_path / "messy.csv"
    path.write_text(CSV, encoding="utf-8")
    return path


def sheet_rows(path, name):
    book = load_workbook(path)
    return [list(r) for r in book[name].iter_rows(values_only=True)]


def test_csv_to_workbook_has_all_five_sheets(csv_path, tmp_path):
    out = tmp_path / "out.xlsx"
    clean_file(csv_path, out)
    assert load_workbook(out).sheetnames == [
        "Read me first", "Cleaned", "Changes", "Removed", "Issues", "Data Dictionary",
        "Summary",
    ]  # fmt: skip

    cleaned = sheet_rows(out, "Cleaned")
    assert cleaned[0] == ["source_row", "Name", "Email", "Phone", "State", "Joined", "Amount"]
    assert cleaned[1] == [
        2,
        "John Smith",
        "john@example.com",
        "(555) 201-0143",
        "KY",
        "2024-03-15",
        1200.5,
    ]
    assert cleaned[2] == [3, "Mary Jones", "bad@@x", "12345", "Kentuky", "03/04/2024", 1200.5]

    changes = sheet_rows(out, "Changes")
    assert changes[0] == ["sheet", "row", "column", "before", "after", "rule", "note"]
    assert ["messy", 2, "Amount", "$1,200.50", 1200.5, "number_parse", None] in changes

    removed = sheet_rows(out, "Removed")
    assert removed[0][:5] == ["sheet", "row", "reason", "duplicate_of_row", "score"]
    assert [r[1:4] for r in removed[1:]] == [[4, "exact duplicate", 2], [5, "empty row", None]]

    issues = {(r[1], r[2], r[4]) for r in sheet_rows(out, "Issues")[1:]}
    assert issues == {
        (3, "Email", "invalid_email"),
        (3, "Phone", "invalid_phone"),
        (3, "State", "unknown_state"),
        (3, "Joined", "ambiguous_date"),
    }

    summary = {r[0]: r[1:] for r in sheet_rows(out, "Summary") if r[0]}
    assert summary["Rows read (header not counted)"][0] == 4
    assert summary["Rows in Cleaned"][0] == 2
    assert summary["Email"][:5] == ["email", "header", 0, 0, 1]


def test_source_row_column_can_be_dropped(csv_path, tmp_path):
    out = tmp_path / "out.xlsx"
    clean_file(csv_path, out, load_config(overrides={"output": {"source_row_column": False}}))
    assert sheet_rows(out, "Cleaned")[0][0] == "Name"


def test_multi_sheet_xlsx(tmp_path):
    book = Workbook()
    ws = book.active
    ws.title = "Customers"
    ws.append(["Name", "Phone", "Joined"])
    ws.append(["ANN LEE", 5552010143, datetime(2024, 3, 5)])
    orders = book.create_sheet("Orders")
    orders.append(["Order ID", "Total"])
    orders.append(["A-1", "$12.50"])
    orders.append(["A-1", "$12.50"])
    book.create_sheet("Empty")
    src = tmp_path / "book.xlsx"
    book.save(src)

    out = tmp_path / "out.xlsx"
    result = clean_file(src, out)
    assert [s.name for s in result.sheets] == ["Customers", "Orders"]
    names = load_workbook(out).sheetnames
    assert names == [
        "Read me first",
        "Cleaned - Customers",
        "Cleaned - Orders",
        "Changes",
        "Removed",
        "Issues",
        "Data Dictionary",
        "Summary",
    ]
    assert sheet_rows(out, "Cleaned - Customers")[1] == [
        2,
        "Ann Lee",
        "(555) 201-0143",
        "2024-03-05",
    ]
    assert sheet_rows(out, "Cleaned - Orders")[1:] == [[2, "A-1", 12.5]]
    assert sheet_rows(out, "Removed")[1][:4] == ["Orders", 3, "exact duplicate", 2]
    assert {r[0] for r in sheet_rows(out, "Changes")[1:]} == {"Customers", "Orders"}


def test_text_starting_with_equals_is_not_written_as_a_formula(tmp_path):
    src = tmp_path / "in.csv"
    src.write_text('Note\n"=HYPERLINK(""http://x"",""click"")"\n', encoding="utf-8")
    out = tmp_path / "out.xlsx"
    clean_file(src, out)
    cell = load_workbook(out)["Cleaned"]["B2"]
    assert cell.data_type == "s"
    assert cell.value.startswith("=HYPERLINK")


def test_semicolon_csv_and_legacy_encoding(tmp_path):
    src = tmp_path / "in.csv"
    src.write_bytes("Name;City\nJos\xe9 Ruiz;M\xe1laga\n".encode("cp1252"))
    result = clean_file(src)
    assert result.sheets[0].rows == [["Jos\xe9 Ruiz", "M\xe1laga"]]


def test_input_is_never_overwritten(tmp_path):
    src = tmp_path / "in.xlsx"
    book = Workbook()
    book.active.append(["Name"])
    book.save(src)
    with pytest.raises(ValueError, match="never overwritten"):
        clean_file(src, src)


def test_cli_writes_default_output_and_reports(csv_path, capsys):
    assert main([str(csv_path)]) == 0
    out = capsys.readouterr().out
    assert "4 rows in, 2 rows out" in out
    assert csv_path.with_name("messy_cleaned.xlsx").exists()


def test_cli_flags_and_config_file(csv_path, tmp_path):
    cfg = tmp_path / "cfg.toml"
    cfg.write_text('[phone]\nformat = "e164"\n', encoding="utf-8")
    out = tmp_path / "o.xlsx"
    assert (
        main([str(csv_path), "-o", str(out), "-c", str(cfg), "--date-order", "mdy", "--no-fuzzy"])
        == 0
    )
    row = sheet_rows(out, "Cleaned")[2]
    assert row[5] == "2024-03-04"
    assert sheet_rows(out, "Cleaned")[1][3] == "+15552010143"


def test_cli_errors_are_clean(tmp_path, capsys):
    assert main([str(tmp_path / "missing.csv")]) == 2
    assert "error:" in capsys.readouterr().err
    bad = tmp_path / "x.pdf"
    bad.write_text("x")
    assert main([str(bad)]) == 2


def test_write_config_and_powershell_redirect_encodings(tmp_path, capsys):
    path = tmp_path / "c.toml"
    assert main(["--write-config", str(path)]) == 0
    assert load_config(path) == load_config()
    assert main(["--write-config", str(path)]) == 2
    text = path.read_text(encoding="utf-8").replace('format = "national"', 'format = "e164"')
    for encoding in ("utf-16", "utf-8-sig"):
        path.write_text(text, encoding=encoding)
        assert load_config(path).phone_format == "e164"
    path.write_text("not [valid toml", encoding="utf-8")
    assert main(["x.csv", "-c", str(path)]) == 2
    assert "not valid TOML" in capsys.readouterr().err


def test_cli_print_config_is_loadable(capsys, tmp_path):
    assert main(["--print-config"]) == 0
    path = tmp_path / "c.toml"
    path.write_text(capsys.readouterr().out, encoding="utf-8")
    assert load_config(path) == load_config()


def test_report_json_counts_match_the_workbook(csv_path, tmp_path):
    import json
    from collections import Counter

    out, report_path = tmp_path / "out.xlsx", tmp_path / "report.json"
    assert main([str(csv_path), "-o", str(out), "--report-json", str(report_path)]) == 0
    report = json.loads(report_path.read_text(encoding="utf-8"))
    result = clean_file(csv_path)

    assert report["report_format"] == 1
    assert report["source"] == "messy.csv"
    assert len(report["input_sha256"]) == 64
    assert report["rows_unaccounted"] == 0
    assert report["rows_in"] == result.rows_in
    assert report["rows_out"] == result.rows_out == len(sheet_rows(out, "Cleaned")) - 1
    assert report["cells_changed"] == len(sheet_rows(out, "Changes")) - 1

    by_rule = Counter(c.rule for c in result.changes)
    assert {r: v["cells"] for r, v in report["changes_by_rule"].items()} == dict(by_rule)
    phone = report["changes_by_rule"]["phone_format"]
    assert phone["columns"] == {"Phone": phone["cells"]}
    counts = [v["cells"] for v in report["changes_by_rule"].values()]
    assert counts == sorted(counts, reverse=True)

    assert report["issues_by_type"] == dict(Counter(i.issue for i in result.issues))
    assert report["removed_by_reason"] == dict(Counter(r.reason for r in result.removed))
    [sheet] = report["sheets"]
    assert sheet["column_types"]["Email"] == "email"
    assert all(t["balance"] == 0 for t in sheet["control_totals"])


def test_report_json_not_written_without_the_flag(csv_path, tmp_path):
    assert main([str(csv_path), "-o", str(tmp_path / "out.xlsx")]) == 0
    assert not list(tmp_path.glob("*.json"))
