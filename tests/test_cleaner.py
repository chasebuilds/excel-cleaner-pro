from datetime import datetime

import pytest

from excel_cleaner_pro import Table, clean_table, load_config


def run(headers, rows, **overrides):
    return clean_table(Table("t", headers, rows), load_config(overrides=overrides or None))


def rules_for(sheet, column):
    return [c.rule for c in sheet.changes if c.column == column and c.row != 1]


def issues_for(sheet, column):
    return [i.issue for i in sheet.issues if i.column == column]


def test_every_edit_is_logged_with_before_and_after():
    sheet = run(["Name", "Email"], [[" JOHN  SMITH ", "John@Example.COM"]])
    assert sheet.rows == [["John Smith", "john@example.com"]]
    logged = [(c.row, c.column, c.before, c.after, c.rule) for c in sheet.changes]
    assert logged == [
        (2, "Name", " JOHN  SMITH ", "JOHN SMITH", "whitespace"),
        (2, "Name", "JOHN SMITH", "John Smith", "name_case"),
        (2, "Email", "John@Example.COM", "john@example.com", "email_lowercase"),
    ]


def test_no_row_disappears_without_a_record():
    rows = [["Ann Lee", "ann@example.com"], ["", " "], ["Ann Lee", "ann@example.com"], ["N/A", "-"]]
    sheet = run(["Name", "Email"], rows)
    assert sheet.rows_in == 4
    assert len(sheet.rows) + len(sheet.removed) == 4
    assert {(r.row, r.reason, r.duplicate_of_row) for r in sheet.removed} == {
        (3, "empty row", None),
        (4, "exact duplicate", 2),
        (5, "empty row", None),
    }


def test_placeholders_become_empty_and_are_counted_missing():
    sheet = run(["Name", "City"], [["Ann Lee", "N/A"], ["Bo Ray", "null"], ["Cy Fox", "Austin"]])
    assert [r[1] for r in sheet.rows] == [None, None, "Austin"]
    assert rules_for(sheet, "City") == ["placeholder", "placeholder"]
    assert sheet.columns[1].missing == 2


def test_invalid_values_are_flagged_and_left_as_typed():
    sheet = run(
        ["Name", "Email", "Phone", "State"],
        [["Ann Lee", "ann@@example.com", "12345", "Kentuky"], ["Bo Ray", None, None, "OH"]],
    )
    assert sheet.rows[0] == ["Ann Lee", "ann@@example.com", "12345", "Kentuky"]
    assert [i.issue for i in sheet.issues] == ["invalid_email", "invalid_phone", "unknown_state"]
    assert not sheet.changes


def test_ambiguous_date_flagged_when_column_gives_no_proof():
    sheet = run(["Joined"], [["03/04/2024"], ["2024-01-09"], ["05/06/2024"]])
    assert [r[0] for r in sheet.rows] == ["03/04/2024", "2024-01-09", "05/06/2024"]
    assert issues_for(sheet, "Joined") == ["ambiguous_date", "ambiguous_date"]


def test_ambiguous_date_flagged_when_column_mixes_orders():
    rows = [["03/04/2024"], ["3/15/2024"], ["3/16/2024"], ["3/17/2024"], ["25/12/2024"]]
    sheet = run(["Joined"], rows)
    assert sheet.rows[0] == ["03/04/2024"]
    assert issues_for(sheet, "Joined") == ["ambiguous_date"]
    assert [r[0] for r in sheet.rows[1:]] == [
        "2024-03-15",
        "2024-03-16",
        "2024-03-17",
        "2024-12-25",
    ]


def test_ambiguous_date_follows_a_consistent_column_and_says_so():
    rows = [["03/04/2024"], ["3/15/2024"], ["3/16/2024"], ["3/17/2024"]]
    sheet = run(["Joined"], rows)
    assert sheet.rows[0] == ["2024-03-04"]
    assert rules_for(sheet, "Joined")[0] == "date_iso_inferred_order"
    assert any("month-first" in note for note in sheet.notes)
    assert not sheet.issues


def test_date_order_flag_mode_never_infers():
    rows = [["03/04/2024"], ["3/15/2024"], ["3/16/2024"], ["3/17/2024"]]
    sheet = run(["Joined"], rows, dates={"order": "flag"})
    assert sheet.rows[0] == ["03/04/2024"]
    assert issues_for(sheet, "Joined") == ["ambiguous_date"]


def test_date_order_can_be_forced():
    sheet = run(["Joined"], [["03/04/2024"]], dates={"order": "dmy"})
    assert sheet.rows == [["2024-04-03"]]


def test_native_excel_dates_become_iso_without_a_change_entry():
    sheet = run(["Joined"], [[datetime(2024, 3, 5)], [datetime(2024, 3, 6, 14, 30)]])
    assert [r[0] for r in sheet.rows] == ["2024-03-05", "2024-03-06 14:30:00"]
    assert not sheet.changes


def test_currency_parsed_and_plain_numbers_converted_quietly():
    rows = [["a", "$1,200.50"], ["b", "1.200,50"], ["c", "980.00"], ["d", "free"]]
    sheet = run(["Ref", "Amount"], rows)
    assert [r[1] for r in sheet.rows] == [1200.5, 1200.5, 980.0, "free"]
    assert rules_for(sheet, "Amount") == ["number_parse", "number_parse"]
    assert issues_for(sheet, "Amount") == ["unparsed_number"]


def test_ambiguous_number_flagged_only_when_column_mixes_separators():
    us_only = run(["Amount"], [["1,500"], ["$2,300.25"]])
    assert [r[0] for r in us_only.rows] == [1500.0, 2300.25]
    mixed = run(["Amount"], [["1,500"], ["$2,300.25"], ["1.200,50"]])
    assert mixed.rows[0] == ["1,500"]
    assert issues_for(mixed, "Amount") == ["ambiguous_number"]
    eu_only = run(["Amount"], [["1.500"], ["2.300,25"]])
    assert [r[0] for r in eu_only.rows] == [1500.0, 2300.25]


def test_zero_padded_codes_stay_text():
    sheet = run(["Zip", "Ref"], [["02134", "00123"], ["40202", "00456"]])
    assert sheet.rows == [["02134", "00123"], ["40202", "00456"]]
    assert [c.type for c in sheet.columns] == ["id", "id"]


def test_company_names_are_not_title_cased():
    sheet = run(["Company Name", "Contact Name"], [["IBM", "ANN LEE"]])
    assert sheet.rows == [["IBM", "Ann Lee"]]


def test_type_inference_from_values_and_config_override():
    rows = [
        ["ann@example.com", "12", "2024-03-05", "x"],
        ["bo@example.com", "7", "March 6, 2024", "y"],
    ]
    sheet = run(["a", "b", "c", "d"], rows)
    assert [(c.type, c.type_source) for c in sheet.columns] == [
        ("email", "values"), ("number", "values"), ("date", "values"), ("text", "default"),
    ]  # fmt: skip
    forced = run(["a", "Tel"], [["x", "5552010143"]], columns={"Tel": "phone"})
    assert forced.columns[1].type_source == "config"
    assert forced.rows == [["x", "(555) 201-0143"]]


def test_rule_can_be_switched_off():
    sheet = run(["Name"], [["JOHN SMITH"]], rules={"name_case": False})
    assert sheet.rows == [["JOHN SMITH"]]


def test_headers_are_fixed_and_logged():
    sheet = run([" Name ", None, "Name"], [["Ann Lee", "x", "Bo Ray"]])
    assert sheet.headers == ["Name", "column_2", "Name_2"]
    assert [c.rule for c in sheet.changes] == ["header", "header", "header"]


def test_ragged_rows_are_padded():
    sheet = run(["a", "b", "c"], [["x"], ["x", "y", "z"]])
    assert sheet.rows == [["x", None, None], ["x", "y", "z"]]


def test_exact_duplicates_found_after_cleaning():
    sheet = run(
        ["Name", "Email"], [["Ann Lee", "ann@example.com"], [" ANN LEE", "Ann@example.com"]]
    )
    assert len(sheet.rows) == 1
    assert sheet.removed[0].reason == "exact duplicate"


def test_fuzzy_duplicate_removed_above_remove_threshold():
    rows = [
        ["Katherine Sanchez", "katherine.sanchez@example.com", "Austin"],
        ["Katherine Sanchez", "katherine.sanches@example.com", "Dallas"],
    ]
    sheet = run(["Name", "Email", "City"], rows, dedupe={"fuzzy_action": "remove"})
    assert len(sheet.rows) == 1
    removed = sheet.removed[0]
    assert (removed.reason, removed.duplicate_of_row, removed.row) == ("fuzzy duplicate", 2, 3)
    assert removed.score >= 95
    assert removed.values[2] == "Dallas"


def test_fuzzy_duplicate_between_thresholds_goes_to_review():
    rows = [["Ann Lee", "ann.lee@example.com"], ["Anne Leigh", "ann.lee@example.com"]]
    sheet = run(["Name", "Email"], rows)
    assert len(sheet.rows) == 2
    issue = sheet.issues[0]
    assert (issue.issue, issue.row, issue.related_row) == ("possible_duplicate", 3, 2)
    assert not sheet.removed


def test_same_name_with_different_numbered_email_is_not_a_duplicate():
    rows = [
        ["James Smith", "james.smith12@example.com"],
        ["James Smith", "james.smith47@example.com"],
    ]
    sheet = run(["Name", "Email"], rows)
    assert len(sheet.rows) == 2
    assert not sheet.issues
    assert not sheet.removed


def test_row_with_a_blank_key_is_never_fuzzy_matched_and_the_summary_says_so():
    rows = [["Ann Lee", "ann.lee@example.com"], ["Ann Lee", None]]
    sheet = run(["Name", "Email"], rows)
    assert len(sheet.rows) == 2
    assert not sheet.issues
    assert any("1 row(s) with a blank key" in note for note in sheet.notes)


def test_reordered_name_and_first_letter_typo_both_match():
    rows = [
        ["Paul Sanchez", "paul.sanchez80@example.net"],
        ["Sanchez, Paul", "paul.sanchez80@example.net"],
        ["Paul anchez", "paul.sanchez80@example.net"],
    ]
    sheet = run(["Name", "Email"], rows, dedupe={"fuzzy_action": "remove"})
    assert [(r.row, r.duplicate_of_row) for r in sheet.removed] == [(3, 2), (4, 2)]


def test_same_mailbox_at_another_domain_is_reviewed_not_removed():
    rows = [
        ["Barbara Baker", "barbara.baker517@example.com"],
        ["Barbara Baker", "barbara.baker517@example.org"],
    ]
    sheet = run(["Name", "Email"], rows)
    assert len(sheet.rows) == 2
    assert [i.issue for i in sheet.issues] == ["possible_duplicate"]


def test_fuzzy_thresholds_and_keys_come_from_config():
    rows = [["Ann Lee", "Austin"], ["Anne Lee", "Austin"], ["Bo Ray", "Austin"]]
    off = run(["Name", "City"], rows)
    assert len(off.rows) == 3
    assert any("skipped" in note for note in off.notes)
    on = run(
        ["Name", "City"],
        rows,
        dedupe={"fuzzy_keys": ["Name"], "remove_threshold": 90, "fuzzy_action": "remove"},
    )
    assert len(on.rows) == 2
    none = run(["Name", "City"], rows, dedupe={"fuzzy_keys": ["Name"], "fuzzy": False})
    assert len(none.rows) == 3


def test_missing_fuzzy_key_is_reported_not_fatal():
    sheet = run(["Name"], [["Ann Lee"]], dedupe={"fuzzy_keys": ["Nope"]})
    assert any("not found" in note for note in sheet.notes)


def test_bad_config_is_rejected():
    from excel_cleaner_pro import ConfigError

    with pytest.raises(ConfigError):
        load_config(overrides={"dates": {"order": "sideways"}})
    with pytest.raises(ConfigError):
        load_config(overrides={"dedupe": {"review_threshold": 99, "remove_threshold": 90}})
    with pytest.raises(ConfigError):
        load_config(overrides={"nonsense": {"x": 1}})
    with pytest.raises(ConfigError):
        load_config(overrides={"columns": {"Tel": "telephone"}})
