"""Nasty files: the ways a spreadsheet misleads a script before any cleaning starts.

Each test builds one damaged file and asserts the tool reports it and does not guess.
"""

import hashlib

import pytest
from openpyxl import Workbook, load_workbook

from excel_cleaner_pro import ConfigError, Table, clean_file, clean_table, load_config
from excel_cleaner_pro import io as writer
from excel_cleaner_pro.cli import main


def run(headers, rows, **overrides):
    return clean_table(Table("t", headers, rows), load_config(overrides=overrides or None))


def issue_kinds(sheet):
    return [i.issue for i in sheet.issues]


def sheet_rows(path, name):
    return [list(r) for r in load_workbook(path)[name].iter_rows(values_only=True)]


def summary(path):
    return {r[0]: r[1:] for r in sheet_rows(path, "Summary") if r[0]}


# ---------------------------------------------------------------- header and layout


def test_title_rows_above_the_header_are_found_and_kept_on_record(tmp_path):
    src = tmp_path / "export.csv"
    src.write_text(
        "Customer export 2024,,\n,,\nName,Email,City\nAnn Lee,ann@example.com,Austin\n",
        encoding="utf-8",
    )
    sheet = clean_file(src).sheets[0]
    assert sheet.header_row == 3
    assert sheet.headers == ["Name", "Email", "City"]
    assert sheet.source_rows == [4]
    assert [(r.row, r.reason) for r in sheet.removed] == [
        (1, "above header row"),
        (2, "above header row"),
    ]
    assert sheet.rows_in == len(sheet.rows) + len(sheet.removed)
    assert any("header row is row 3" in note for note in sheet.notes)


def test_group_labels_above_the_header_are_reported_as_a_possible_multi_row_header(tmp_path):
    src = tmp_path / "grouped.csv"
    src.write_text(
        "Contact,,Billing,\nName,Email,City,State\nAnn Lee,ann@example.com,Austin,TX\n",
        encoding="utf-8",
    )
    # row 1 is half filled with text, so it would be taken as the header: say which row it is
    sheet = clean_file(src, cfg=load_config(overrides={"input": {"header_row": 2}})).sheets[0]
    assert sheet.headers == ["Name", "Email", "City", "State"]
    assert any("possible multi-row header" in note for note in sheet.notes)
    assert any("set in the config" in note for note in sheet.notes)


def test_header_row_past_the_end_is_an_error(tmp_path):
    src = tmp_path / "short.csv"
    src.write_text("Name\nAnn\n", encoding="utf-8")
    with pytest.raises(ValueError, match="header_row"):
        clean_file(src, cfg=load_config(overrides={"input": {"header_row": 9}}))


def test_total_line_is_flagged_and_kept():
    sheet = run(["Item", "Amount"], [["Widget", "10.00"], ["Gadget", "5.50"], ["Total", "15.50"]])
    assert len(sheet.rows) == 3
    assert [(i.row, i.issue) for i in sheet.issues] == [(4, "possible_total_row")]


# ---------------------------------------------------------------- Excel structure


@pytest.fixture
def nasty_xlsx(tmp_path):
    book = Workbook()
    ws = book.active
    ws.title = "Orders"
    ws.append(["Region", "Order ID", "Qty", "Unit", "Total"])
    ws.append(["West", "A-1", 2, 5.0, "=C2*D2"])
    ws.append([None, "A-2", 1, 7.5, 7.5])
    ws.append(["East", "A-3", 4, 1.0, 4.0])
    ws.merge_cells("A2:A3")
    ws.row_dimensions[4].hidden = True
    ws.column_dimensions["D"].hidden = True
    hidden = book.create_sheet("Archive")
    hidden.append(["Code"])
    hidden.append(["X1"])
    hidden.sheet_state = "hidden"
    book.create_sheet("Blank")
    path = tmp_path / "nasty.xlsx"
    book.save(path)
    return path


def test_merged_cells_are_flagged_and_not_filled(nasty_xlsx):
    orders = clean_file(nasty_xlsx).sheets[0]
    merged = [i for i in orders.issues if i.issue == "merged_cells"]
    assert [(i.row, i.column, i.value) for i in merged] == [(2, "Region", "West")]
    assert "A2:A3" in merged[0].detail
    assert orders.rows[1][0] is None


def test_formula_without_a_stored_value_is_flagged(nasty_xlsx):
    orders = clean_file(nasty_xlsx).sheets[0]
    found = [i for i in orders.issues if i.issue == "formula_no_value"]
    assert [(i.row, i.column) for i in found] == [(2, "Total")]
    assert "=C2*D2" in found[0].detail
    assert any("1 formula cell(s)" in note for note in orders.notes)


def test_hidden_rows_columns_and_sheets_are_read_and_reported(nasty_xlsx, tmp_path):
    out = tmp_path / "out.xlsx"
    result = clean_file(nasty_xlsx, out)
    orders, archive = result.sheets
    assert len(orders.rows) == 3
    assert any("1 hidden row(s)" in note for note in orders.notes)
    assert any("hidden column(s)" in note and "'D'" in note for note in orders.notes)
    assert any("sheet is hidden" in note for note in archive.notes)
    record = summary(out)
    assert record["Sheets read"][0] == "Orders, Archive"
    assert record["Sheets skipped (empty)"][0] == "Blank"
    assert record["Date system"][0] == "1900"


# ---------------------------------------------------------------- damaged identifiers


def test_ids_stored_as_numbers_and_short_zips_are_flagged_not_padded(tmp_path):
    book = Workbook()
    ws = book.active
    ws.append(["Name", "ZIP", "Account No"])
    ws.append(["Ann Lee", 2134, 12345678901234560000])
    ws.append(["Bo Ray", 40202, 77])
    src = tmp_path / "ids.xlsx"
    book.save(src)
    sheet = clean_file(src).sheets[0]
    assert sheet.rows[0][1] == 2134
    found = {(i.column, i.issue) for i in sheet.issues}
    assert found == {
        ("ZIP", "id_stored_as_number"),
        ("ZIP", "possible_lost_leading_zero"),
        ("Account No", "id_stored_as_number"),
        ("Account No", "possible_truncated_digits"),
    }
    assert not sheet.changes


def test_scientific_notation_and_date_like_codes_are_flagged():
    rows = [["1.23457E+15", "MARCH1"], ["4411223344556677", "2-Sep"], ["4411223344556688", "DEC1"]]
    sheet = run(["Card ID", "Gene Code"], rows)
    assert sheet.rows[0][0] == "1.23457E+15"
    assert sheet.rows[1][1] == "2-Sep"
    assert {(i.column, i.issue, i.row) for i in sheet.issues} == {
        ("Card ID", "scientific_notation", 2),
        ("Gene Code", "possible_date_conversion", 3),
    }


def test_column_level_issue_names_the_rows():
    sheet = run(["Zip"], [["2134"], ["40202"], ["610"], ["90210"]])
    (issue,) = sheet.issues
    assert issue.issue == "possible_lost_leading_zero"
    assert "2 value(s)" in issue.detail
    assert "rows 2, 4" in issue.detail


# ---------------------------------------------------------------- text encoding


def test_csv_encoding_and_delimiter_are_stated_as_assumptions(tmp_path):
    out = tmp_path / "out.xlsx"
    legacy = tmp_path / "legacy.csv"
    legacy.write_bytes("Name;City\nJos\xe9 Ruiz;M\xe1laga\n".encode("cp1252"))
    clean_file(legacy, out)
    record = summary(out)
    assert "Windows-1252 (guess" in record["Encoding assumed"][0]
    assert record["Delimiter assumed"][0] == "semicolon"

    modern = tmp_path / "modern.csv"
    modern.write_bytes(b"\xef\xbb\xbfName,City\nAnn,Austin\n")
    manifest = clean_file(modern).manifest
    assert manifest["Encoding assumed"] == "UTF-8 with byte order mark"
    assert manifest["Delimiter assumed"] == "comma"


def test_mojibake_is_flagged_and_not_repaired(tmp_path):
    src = tmp_path / "moji.csv"
    # UTF-8 bytes that an earlier tool decoded as Windows-1252 and saved again
    src.write_text("Name,City\nJosÃ© Ruiz,Austin\nAnn Lee,Austin\n", encoding="utf-8")
    sheet = clean_file(src).sheets[0]
    assert sheet.rows[0][0] == "JosÃ© Ruiz"
    assert [(i.row, i.column, i.issue) for i in sheet.issues] == [(2, "Name", "possible_mojibake")]


def test_nfc_joins_letter_and_accent_and_logs_it_but_keeps_other_characters():
    sheet = run(["City", "Unit"], [["Mérida", "m²"]])
    assert sheet.rows == [["Mérida", "m²"]]
    assert [(c.column, c.rule) for c in sheet.changes] == [("City", "unicode_nfc")]
    off = run(["City"], [["Mérida"]], rules={"unicode_nfc": False})
    assert off.rows == [["Mérida"]]


# ---------------------------------------------------------------- duplicates


def test_fuzzy_match_is_only_listed_by_default_even_at_a_very_high_score():
    rows = [
        ["Katherine Sanchez", "katherine.sanchez@example.com"],
        ["Katherine Sanchez", "katherine.sanches@example.com"],
    ]
    sheet = run(["Name", "Email"], rows)
    assert len(sheet.rows) == 2
    assert not sheet.removed
    assert [(i.issue, i.row, i.related_row) for i in sheet.issues] == [("possible_duplicate", 3, 2)]
    assert any("never removed" in note for note in sheet.notes)


def test_removal_fills_blanks_in_the_kept_row_and_logs_where_each_value_came_from():
    rows = [
        ["Katherine Sanchez", "katherine.sanchez@example.com", None, "Austin", "100.00"],
        ["Katherine Sanchez", "katherine.sanches@example.com", "555-201-0143", "Dallas", "40.00"],
    ]
    sheet = run(
        ["Name", "Email", "Phone", "City", "Amount"], rows, dedupe={"fuzzy_action": "remove"}
    )
    assert sheet.rows == [
        ["Katherine Sanchez", "katherine.sanchez@example.com", "(555) 201-0143", "Austin", 100.0]
    ]
    fills = [c for c in sheet.changes if c.rule == "survivor_fill"]
    assert [(c.row, c.column, c.before, c.after) for c in fills] == [
        (2, "Phone", None, "(555) 201-0143")
    ]
    assert "removed duplicate row 3" in fills[0].note
    # the kept row's own city and amount win; the removed row keeps its values in Removed
    assert sheet.removed[0].values[3:] == ["Dallas", 40.0]


# ---------------------------------------------------------------- reconciliation


def test_control_totals_balance_through_removal_and_survivor_fill():
    rows = [
        ["Ann Lee", "ann.lee@example.com", None],
        ["Ann Lee", "ann.lee@example.com", None],
        ["Katherine Sanchez", "katherine.sanchez@example.com", None],
        ["Katherine Sanchez", "katherine.sanches@example.com", "250.00"],
        ["Bo Ray", "bo.ray@example.com", "$1,000.50"],
        ["Cy Fox", "cy.fox@example.com", "ask"],
    ]
    sheet = run(["Name", "Email", "Amount"], rows, dedupe={"fuzzy_action": "remove"})
    (total,) = sheet.control_totals
    assert total.column == "Amount"
    assert total.total_before_removal == 1250.5
    assert total.in_cleaned == 1250.5
    assert total.in_removed == 250.0
    assert total.filled_from_removed == 250.0
    assert total.balance == 0
    assert total.not_a_number == 1


def test_summary_shows_reconciliation_run_record_and_config(tmp_path):
    src = tmp_path / "in.csv"
    src.write_text("Name,Amount\nAnn Lee,10.00\nAnn Lee,10.00\nBo Ray,N/A\nCy Fox,n/a\n", "utf-8")
    out = tmp_path / "out.xlsx"
    result = clean_file(src, out)
    record = summary(out)
    assert record["Rows read (header not counted)"][0] == 4
    assert record["Read - Cleaned - Removed (must be 0)"][0] == 0
    assert record["Input SHA-256"][0] == hashlib.sha256(src.read_bytes()).hexdigest()
    assert record["Input size (bytes)"][0] == src.stat().st_size
    assert record["Tool"][0].startswith("excel-cleaner-pro ")
    assert record["dedupe.fuzzy_action"][0] == "review"
    assert record["dates.order"][0] == "infer"
    rows = sheet_rows(out, "Summary")
    assert ["in", "Amount", 20, 10, 10, 0, 0, 0] in [r[:8] for r in rows]
    assert ["in", "Amount", "n/a", 2] in [r[:4] for r in rows]
    assert result.sheets[0].placeholders == {("Amount", "n/a"): 2}


# ---------------------------------------------------------------- validation rules


def test_declared_rules_report_and_never_edit(tmp_path):
    src = tmp_path / "in.csv"
    src.write_text(
        "ID,Email,State,Qty\nA1,ann@example.com,KY,5\nA1,,OH,900\nB22,bo@example.com,ZZ,x\n",
        encoding="utf-8",
    )
    rules = {
        "unique": ["ID"],
        "required": ["Email", "Missing Column"],
        "allowed": {"State": ["KY", "OH"]},
        "range": {"Qty": [0, 100]},
        "pattern": {"ID": "[A-Z]\\d"},
    }
    cfg = load_config(overrides={"validate": rules, "columns": {"State": "text", "Qty": "number"}})
    out = tmp_path / "out.xlsx"
    result = clean_file(src, out, cfg)
    got = {
        (r.rule.split(":")[0], r.column): (r.status, r.failures)
        for r in result.sheets[0].validation
    }
    assert got["unique", "ID"] == ("fail", [(2, "A1"), (3, "A1")])
    assert got["required", "Email"] == ("fail", [(3, None)])
    assert got["required", "Missing Column"] == ("column not found", [])
    assert got["allowed", "State"] == ("fail", [(4, "ZZ")])
    assert got["range", "Qty"] == ("fail", [(3, 900), (4, "x")])
    assert got["pattern", "ID"] == ("fail", [(4, "B22")])
    assert len(result.sheets[0].rows) == 3

    names = load_workbook(out).sheetnames
    assert names.index("Validation") == names.index("Issues") + 1
    lines = sheet_rows(out, "Validation")
    assert ["in", "unique", "ID", "fail", 2, None, None] in lines
    assert ["in", "unique", "ID", "failing row", None, 3, "A1"] in lines
    assert ["Validation" in str(r[0]) for r in sheet_rows(out, "Read me first")].count(True) == 1


def test_passing_rules_and_cli_report(tmp_path, capsys):
    src = tmp_path / "in.csv"
    src.write_text("ID,Email\nA1,ann@example.com\nA2,\n", encoding="utf-8")
    cfg = tmp_path / "rules.toml"
    cfg.write_text('[validate]\nunique = ["ID"]\nrequired = ["Email"]\n', encoding="utf-8")
    assert main([str(src), "-c", str(cfg)]) == 0
    assert "validation rules not passed: 1" in capsys.readouterr().out


def test_bad_rule_config_is_rejected():
    with pytest.raises(ConfigError, match="range"):
        load_config(overrides={"validate": {"range": {"Qty": [1]}}})
    with pytest.raises(ConfigError, match="pattern"):
        load_config(overrides={"validate": {"pattern": {"ID": "("}}})
    with pytest.raises(ConfigError, match="fuzzy_action"):
        load_config(overrides={"dedupe": {"fuzzy_action": "delete"}})


# ---------------------------------------------------------------- the workbook itself


def test_read_me_first_and_data_dictionary(tmp_path):
    src = tmp_path / "in.csv"
    src.write_text(
        "Name,Joined,Amount\nANN LEE,3/15/2024,$10.50\nBo Ray,2024-01-02,7\nCy Fox,,N/A\n",
        encoding="utf-8",
    )
    out = tmp_path / "out.xlsx"
    clean_file(src, out)
    book = load_workbook(out)
    assert book.sheetnames[0] == "Read me first"
    text = " ".join(str(r[0]) for r in sheet_rows(out, "Read me first") if r[0])
    assert "3 rows read" in text
    assert "Your original file was not modified." in text

    rows = sheet_rows(out, "Data Dictionary")
    assert rows[0][:4] == ["sheet", "column", "type", "type decided by"]
    by_col = {r[1]: r for r in rows[1:]}
    assert by_col["Name"][2:8] == ["name", "header", 3, 0, 0, 3]
    assert by_col["Name"][13] == "name_case"
    assert by_col["Joined"][8:10] == ["2024-01-02", "2024-03-15"]
    assert by_col["Amount"][4:10] == [2, 1, 33.3, 2, 7, 10.5]
    assert by_col["Amount"][13] == "number_parse, placeholder"


def test_text_that_looks_like_a_formula_is_stored_as_text_and_counted(tmp_path):
    src = tmp_path / "in.csv"
    src.write_text(
        'Ref,Note\n1,"=1+1"\n2,+SUM(A1)\n3,-cmd\n4,@import\n5,-12.50\n6,plain\n', encoding="utf-8"
    )
    out = tmp_path / "out.xlsx"
    result = clean_file(src, out)
    ws = load_workbook(out)["Cleaned"]
    for row in range(2, 6):
        assert ws.cell(row, 3).data_type == "s"
    assert any("4 text value(s) start with" in note for note in result.sheets[0].notes)


def test_too_many_rows_for_one_sheet_is_refused(tmp_path, monkeypatch):
    monkeypatch.setattr(writer, "EXCEL_MAX_DATA_ROWS", 2)
    src = tmp_path / "in.csv"
    src.write_text("Name\na\nb\nc\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Split the input file"):
        clean_file(src, tmp_path / "out.xlsx")
    assert main([str(src)]) == 2


def test_cli_fuzzy_remove_and_header_row(tmp_path, capsys):
    src = tmp_path / "in.csv"
    src.write_text(
        "Report\nName,Email\nKatherine Sanchez,katherine.sanchez@example.com\n"
        "Katherine Sanchez,katherine.sanches@example.com\n",
        encoding="utf-8",
    )
    assert main([str(src), "--header-row", "2"]) == 0
    assert "3 rows in, 2 rows out" in capsys.readouterr().out
    assert main([str(src), "--header-row", "2", "--fuzzy-remove"]) == 0
    assert "1 fuzzy duplicate" in capsys.readouterr().out


def test_long_codes_kept_as_text_are_flagged_only_when_every_one_ends_in_zero():
    damaged = run(["Card ID"], [["4411223344556670"], ["4411223344556680"]])
    assert issue_kinds(damaged) == ["possible_truncated_digits"]
    intact = run(["Card ID"], [["4411223344556670"], ["4411223344556688"]])
    assert not intact.issues
