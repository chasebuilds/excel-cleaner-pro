import pytest

from excel_cleaner_pro import rules


@pytest.mark.parametrize(
    ("raw", "want"),
    [
        ("  John   Smith ", "John Smith"),
        ("John\tSmith", "John Smith"),
        ("John Smith ", "John Smith"),
        ("﻿John​", "John"),
        ("   ", ""),
    ],
)
def test_whitespace(raw, want):
    assert rules.clean_whitespace(raw) == want


@pytest.mark.parametrize(
    ("raw", "want"),
    [
        ("JOHN SMITH", "John Smith"),
        ("john smith", "John Smith"),
        ("MARY O'BRIEN", "Mary O'Brien"),
        ("ronald mcdonald", "Ronald McDonald"),
        ("ANNA SMITH-JONES", "Anna Smith-Jones"),
        ("HENRY FORD III", "Henry Ford III"),
        # mixed case is left alone: someone chose it
        ("Maria de la Cruz", "Maria de la Cruz"),
        ("LeBron MacDonald", "LeBron MacDonald"),
    ],
)
def test_name_case(raw, want):
    assert rules.clean_name(raw) == want


@pytest.mark.parametrize(
    "email",
    ["a@example.com", "first.last+tag@sub.example.org", "o'brien@example.co.uk", "A@EXAMPLE.COM"],
)
def test_valid_email(email):
    assert rules.is_valid_email(email)


@pytest.mark.parametrize(
    "email",
    ["a@@example.com", "aexample.com", "a@example", "a b@example.com", "a@example,com",
     ".a@example.com", "a..b@example.com", "a@-example.com", "a@"],
)  # fmt: skip
def test_invalid_email(email):
    assert not rules.is_valid_email(email)


@pytest.mark.parametrize(
    "raw",
    ["555-201-0143", "555.201.0143", "5552010143", "+1 555 201 0143", "1-555-201-0143",
     "(555)201-0143", "(555) 201-0143", "+15552010143"],
)  # fmt: skip
def test_phone_formats(raw):
    assert rules.clean_phone(raw).value == "(555) 201-0143"


def test_phone_output_styles_and_extension():
    assert rules.clean_phone("555-201-0143", "e164").value == "+15552010143"
    assert rules.clean_phone("555-201-0143", "dashes").value == "555-201-0143"
    assert rules.clean_phone("555-201-0143 ext. 22").value == "(555) 201-0143 x22"


@pytest.mark.parametrize(
    ("raw", "issue"),
    [
        ("201-0143", "invalid_phone"),
        ("555-201-01431", "invalid_phone"),
        ("555-CALL-NOW", "invalid_phone"),
        ("155-201-0143", "invalid_phone"),
        ("+44 20 7946 0958", "non_us_phone"),
    ],
)
def test_phone_flagged(raw, issue):
    result = rules.clean_phone(raw)
    assert result.value is None
    assert result.issue == issue


@pytest.mark.parametrize(
    ("raw", "want"),
    [
        ("2024-03-15", "2024-03-15"),
        ("2024/3/5", "2024-03-05"),
        ("March 5, 2024", "2024-03-05"),
        ("Mar 5 2024", "2024-03-05"),
        ("Sept. 5, 2024", "2024-09-05"),
        ("5-Mar-2024", "2024-03-05"),
        ("05 March 2024", "2024-03-05"),
        ("5th March 2024", "2024-03-05"),
        ("Tue, 5 Mar 2024", "2024-03-05"),
        ("3/15/2024", "2024-03-15"),
        ("15/3/2024", "2024-03-15"),
        ("15.03.2024", "2024-03-15"),
        ("05/05/2024", "2024-05-05"),
        ("2024-03-15 00:00:00", "2024-03-15"),
        ("2024-03-15T14:30:00", "2024-03-15 14:30:00"),
        ("3/15/2024 2:30 PM", "2024-03-15 14:30:00"),
    ],
)
def test_date_ok(raw, want):
    result = rules.parse_date(raw)
    assert (result.kind, result.iso) == ("ok", want)


def test_date_evidence():
    assert rules.parse_date("3/15/2024").evidence == "mdy"
    assert rules.parse_date("15/3/2024").evidence == "dmy"
    assert rules.parse_date("05/05/2024").evidence is None


def test_date_ambiguous_is_not_guessed():
    result = rules.parse_date("03/04/2024")
    assert result.kind == "ambiguous"
    assert result.iso is None
    assert (result.mdy, result.dmy) == ("2024-03-04", "2024-04-03")


def test_date_two_digit_year_uses_pivot():
    assert rules.parse_date("3/15/24").iso == "2024-03-15"
    assert rules.parse_date("3/15/87").iso == "1987-03-15"
    assert rules.parse_date("3/15/24").two_digit_year


@pytest.mark.parametrize(
    "raw", ["02/30/2024", "2024-13-01", "Q1 2024", "03/2024", "soon", "13/13/2024"]
)
def test_date_bad(raw):
    assert rules.parse_date(raw).kind == "bad"


@pytest.mark.parametrize(
    ("raw", "want"),
    [
        ("$1,200.50", 1200.5),
        ("1,200.50", 1200.5),
        ("1.200,50", 1200.5),
        ("€1.200,50", 1200.5),
        ("1 200,50", 1200.5),
        ("USD 1200.50", 1200.5),
        ("1200,5", 1200.5),
        ("$45", 45),
        ("(1,200.50)", -1200.5),
        ("-$12.00", -12.0),
        ("1,234,567", 1234567),
        ("1.234.567", 1234567),
        ("0.125", 0.125),
        ("12", 12),
    ],
)
def test_number_ok(raw, want):
    result = rules.parse_number(raw)
    assert result.kind == "ok"
    assert result.value == pytest.approx(want)


def test_number_separator_evidence():
    assert rules.parse_number("1,200.50").evidence == "us"
    assert rules.parse_number("1.200,50").evidence == "eu"
    assert rules.parse_number("$1,200.50").currency


@pytest.mark.parametrize(("raw", "us", "eu"), [("1,200", 1200, 1.2), ("1.200", 1.2, 1200)])
def test_number_ambiguous_keeps_both_readings(raw, us, eu):
    result = rules.parse_number(raw)
    assert result.kind == "ambiguous"
    assert (result.us, result.eu) == (pytest.approx(us), pytest.approx(eu))


@pytest.mark.parametrize("raw", ["twelve", "12%", "1,2,3", "1.2.3", "$", "12-34", "1,20.5"])
def test_number_bad(raw):
    assert rules.parse_number(raw).kind == "bad"


@pytest.mark.parametrize(
    ("raw", "want"),
    [("Kentucky", "KY"), ("kentucky", "KY"), ("NEW YORK", "NY"), ("ky", "KY"), ("N.Y.", "NY"),
     ("District of Columbia", "DC")],
)  # fmt: skip
def test_state(raw, want):
    assert rules.clean_state(raw).value == want


def test_state_typo_is_suggested_not_applied():
    result = rules.clean_state("Kentuky")
    assert result.value is None
    assert "Kentucky (KY)" in result.detail
    assert rules.clean_state("Atlantis").value is None
