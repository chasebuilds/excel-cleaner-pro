"""Value-level cleaning rules. Pure functions: one cell in, one verdict out."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date

from rapidfuzz import fuzz, process

# \s misses zero-width space and the BOM, which both show up in pasted spreadsheet text
_WS_RE = re.compile(r"[\s​﻿]+")


def clean_whitespace(text: str) -> str:
    return _WS_RE.sub(" ", text).strip()


# ---------------------------------------------------------------- names

_ROMAN = {"ii", "iii", "iv", "vi", "vii", "viii"}


def _cap_part(part: str) -> str:
    low = part.lower()
    if low in _ROMAN:
        return low.upper()
    if low.startswith("mc") and len(low) > 3:
        return "Mc" + low[2:].capitalize()
    return low.capitalize()


def _cap_token(token: str) -> str:
    pieces = []
    for hyphen_part in token.split("-"):
        head, sep, tail = hyphen_part.partition("'")
        # O'Brien and D'Angelo get a capital after the apostrophe, a possessive (Smith's) does not
        if sep and len(head) <= 2 and tail:
            pieces.append(_cap_part(head) + "'" + _cap_part(tail))
        elif sep:
            pieces.append(_cap_part(head) + "'" + tail.lower())
        else:
            pieces.append(_cap_part(hyphen_part))
    return "-".join(pieces)


def clean_name(text: str) -> str:
    """Fix case only when the whole value is upper or lower. Mixed case is someone's choice."""
    if not (text.isupper() or text.islower()):
        return text
    return " ".join(_cap_token(tok) for tok in text.split(" "))


# ---------------------------------------------------------------- email

_EMAIL_RE = re.compile(
    r"^[a-z0-9_%+\-']+(?:\.[a-z0-9_%+\-']+)*"
    r"@(?:[a-z0-9](?:[a-z0-9\-]*[a-z0-9])?\.)+[a-z]{2,}$"
)


def is_valid_email(text: str) -> bool:
    return bool(_EMAIL_RE.match(text.lower()))


# ---------------------------------------------------------------- phone

_EXT_RE = re.compile(r"\s*(?:ext\.?|extension|x|#)\s*(\d{1,6})\s*$", re.IGNORECASE)


@dataclass(frozen=True)
class PhoneResult:
    value: str | None
    issue: str | None = None
    detail: str = ""


def clean_phone(text: str, fmt: str = "national") -> PhoneResult:
    body = text
    ext = ""
    if m := _EXT_RE.search(text):
        ext = m.group(1)
        body = text[: m.start()]
    if re.search(r"[A-Za-z]", body):
        return PhoneResult(None, "invalid_phone", "contains letters")
    digits = re.sub(r"\D", "", body)
    if body.strip().startswith("+") and not digits.startswith("1"):
        return PhoneResult(None, "non_us_phone", "country code is not +1, left as typed")
    if len(digits) == 11 and digits[0] == "1":
        digits = digits[1:]
    if len(digits) != 10:
        return PhoneResult(None, "invalid_phone", f"{len(digits)} digits, expected 10")
    # North American numbering plan: area code and exchange never start with 0 or 1
    if digits[0] in "01" or digits[3] in "01":
        return PhoneResult(None, "invalid_phone", "area code or exchange starts with 0 or 1")
    a, b, c = digits[:3], digits[3:6], digits[6:]
    if fmt == "e164":
        out = f"+1{digits}"
    elif fmt == "dashes":
        out = f"{a}-{b}-{c}"
    else:
        out = f"({a}) {b}-{c}"
    return PhoneResult(out + (f" x{ext}" if ext else ""))


# ---------------------------------------------------------------- dates

_MONTHS = {
    name: i
    for i, names in enumerate(
        [
            ("jan", "january"),
            ("feb", "february"),
            ("mar", "march"),
            ("apr", "april"),
            ("may",),
            ("jun", "june"),
            ("jul", "july"),
            ("aug", "august"),
            ("sep", "sept", "september"),
            ("oct", "october"),
            ("nov", "november"),
            ("dec", "december"),
        ],
        start=1,
    )
    for name in names
}

_TIME_RE = re.compile(
    r"[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.\d+)?\s*(am|pm)?\s*(?:z|utc)?$", re.IGNORECASE
)
_WEEKDAY_RE = re.compile(
    r"^(?:mon|tue|tues|wed|thu|thur|thurs|fri|sat|sun)[a-z]*\.?,?\s+", re.IGNORECASE
)
_YMD_RE = re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$")
_NUMERIC_RE = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})$")
_DAY_MONTH_RE = re.compile(
    r"^(\d{1,2})(?:st|nd|rd|th)?[-\s/.]*([A-Za-z]{3,9})\.?,?[-\s/.]*(\d{4}|\d{2})$"
)
_MONTH_DAY_RE = re.compile(r"^([A-Za-z]{3,9})\.?\s*(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$")


@dataclass(frozen=True)
class DateResult:
    """kind: ok (iso set), ambiguous (mdy and dmy set), bad (not a date)."""

    kind: str
    iso: str | None = None
    mdy: str | None = None
    dmy: str | None = None
    evidence: str | None = None  # "mdy" or "dmy" when this value proves the column's order
    two_digit_year: bool = False


def _iso(year: int, month: int, day: int, suffix: str) -> str | None:
    try:
        return date(year, month, day).isoformat() + suffix
    except ValueError:
        return None


def _year(text: str, pivot: int) -> int:
    if len(text) == 4:
        return int(text)
    yy = int(text)
    return 2000 + yy if yy < pivot else 1900 + yy


def parse_date(text: str, pivot: int = 50) -> DateResult:
    s = _WEEKDAY_RE.sub("", text.strip())
    suffix = ""
    if m := _TIME_RE.search(s):
        hour, minute, second = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        meridiem = (m.group(4) or "").lower()
        if meridiem == "pm" and hour < 12:
            hour += 12
        elif meridiem == "am" and hour == 12:
            hour = 0
        if hour > 23 or minute > 59 or second > 59:
            return DateResult("bad")
        # midnight is how a date-only cell exports, so it carries no information
        if (hour, minute, second) != (0, 0, 0):
            suffix = f" {hour:02d}:{minute:02d}:{second:02d}"
        s = s[: m.start()].strip()

    if m := _YMD_RE.match(s):
        iso = _iso(int(m.group(1)), int(m.group(2)), int(m.group(3)), suffix)
        return DateResult("ok", iso) if iso else DateResult("bad")

    if m := _NUMERIC_RE.match(s):
        a, b = int(m.group(1)), int(m.group(2))
        year = _year(m.group(3), pivot)
        short = len(m.group(3)) == 2
        mdy = _iso(year, a, b, suffix)
        dmy = _iso(year, b, a, suffix)
        if mdy and dmy and a != b:
            return DateResult("ambiguous", mdy=mdy, dmy=dmy, two_digit_year=short)
        if mdy and dmy:
            return DateResult("ok", mdy, two_digit_year=short)
        if mdy:
            return DateResult("ok", mdy, evidence="mdy", two_digit_year=short)
        if dmy:
            return DateResult("ok", dmy, evidence="dmy", two_digit_year=short)
        return DateResult("bad")

    if (m := _DAY_MONTH_RE.match(s)) and m.group(2).lower() in _MONTHS:
        year = _year(m.group(3), pivot)
        iso = _iso(year, _MONTHS[m.group(2).lower()], int(m.group(1)), suffix)
        short = len(m.group(3)) == 2
        return DateResult("ok", iso, two_digit_year=short) if iso else DateResult("bad")

    if (m := _MONTH_DAY_RE.match(s)) and m.group(1).lower() in _MONTHS:
        iso = _iso(int(m.group(3)), _MONTHS[m.group(1).lower()], int(m.group(2)), suffix)
        return DateResult("ok", iso) if iso else DateResult("bad")

    return DateResult("bad")


# ---------------------------------------------------------------- numbers

_CURRENCY_RE = re.compile(r"(?:[$€£¥]|\b(?:usd|eur|gbp|cad|aud)\b)", re.IGNORECASE)
_PLAIN_RE = re.compile(r"^-?\d+(?:\.\d+)?$")
_GROUPED_RE = re.compile(r"^\d{1,3}(?:[.,]\d{3})+$")


@dataclass(frozen=True)
class NumberResult:
    """kind: ok (value set), ambiguous (us and eu readings set), bad."""

    kind: str
    value: float | int | None = None
    us: float | int | None = None
    eu: float | int | None = None
    evidence: str | None = None  # "us" or "eu" when the text proves which separator is the decimal
    plain: bool = False  # already a bare number literal, so converting it changes nothing
    currency: bool = False


def _num(int_part: str, frac_part: str = "") -> float | int:
    return float(f"{int_part}.{frac_part}") if frac_part else int(int_part)


def parse_number(text: str) -> NumberResult:
    s = text.strip()
    if _PLAIN_RE.match(s):
        # one dot and exactly three decimals reads as 1.250 or as 1250 depending on the locale
        head, _, tail = s.lstrip("-").partition(".")
        value = float(s) if "." in s else int(s)
        if len(tail) == 3 and 1 <= len(head) <= 3 and not head.startswith("0"):
            eu = int(head + tail) * (-1 if s.startswith("-") else 1)
            return NumberResult("ambiguous", us=value, eu=eu, plain=True)
        return NumberResult("ok", value, evidence="us" if "." in s else None, plain=True)

    currency = bool(_CURRENCY_RE.search(s))
    s = _CURRENCY_RE.sub("", s)
    s = re.sub(r"[\s  ']", "", s)
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative, s = True, s[1:-1]
    if s.startswith("-") or s.endswith("-"):
        negative, s = True, s.strip("-")
    s = s.removeprefix("+")
    if not s or not re.fullmatch(r"[\d.,]+", s) or not s[0].isdigit() or not s[-1].isdigit():
        return NumberResult("bad")

    sign = -1 if negative else 1
    dots, commas = s.count("."), s.count(",")

    def done(value: float | int, evidence: str | None) -> NumberResult:
        return NumberResult("ok", sign * value, evidence=evidence, currency=currency)

    if dots and commas:
        decimal = "." if s.rfind(".") > s.rfind(",") else ","
        if s.count(decimal) != 1:
            return NumberResult("bad")
        int_part, frac_part = s.split(decimal)
        if not _GROUPED_RE.match(int_part):
            return NumberResult("bad")
        return done(
            _num(re.sub(r"[.,]", "", int_part), frac_part), "us" if decimal == "." else "eu"
        )

    sep = "." if dots else ","
    count = dots or commas
    if count == 0:
        return done(int(s), None)
    if count > 1:
        if not _GROUPED_RE.match(s):
            return NumberResult("bad")
        return done(int(s.replace(sep, "")), "us" if sep == "," else "eu")

    head, tail = s.split(sep)
    if len(tail) == 3 and 1 <= len(head) <= 3 and not head.startswith("0"):
        grouped, decimal = int(head + tail), float(f"{head}.{tail}")
        us, eu = (grouped, decimal) if sep == "," else (decimal, grouped)
        return NumberResult("ambiguous", us=sign * us, eu=sign * eu, currency=currency)
    return done(float(f"{head}.{tail}"), "us" if sep == "." else "eu")


# ---------------------------------------------------------------- US states

STATES = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR", "california": "CA",
    "colorado": "CO", "connecticut": "CT", "delaware": "DE", "district of columbia": "DC",
    "florida": "FL", "georgia": "GA", "hawaii": "HI", "idaho": "ID", "illinois": "IL",
    "indiana": "IN", "iowa": "IA", "kansas": "KS", "kentucky": "KY", "louisiana": "LA",
    "maine": "ME", "maryland": "MD", "massachusetts": "MA", "michigan": "MI", "minnesota": "MN",
    "mississippi": "MS", "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK", "oregon": "OR",
    "pennsylvania": "PA", "rhode island": "RI", "south carolina": "SC", "south dakota": "SD",
    "tennessee": "TN", "texas": "TX", "utah": "UT", "vermont": "VT", "virginia": "VA",
    "washington": "WA", "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
    "puerto rico": "PR", "guam": "GU", "american samoa": "AS", "northern mariana islands": "MP",
    "us virgin islands": "VI",
}  # fmt: skip
STATE_CODES = frozenset(STATES.values())


@dataclass(frozen=True)
class StateResult:
    value: str | None
    detail: str = ""


def clean_state(text: str) -> StateResult:
    key = text.replace(".", "").casefold()
    if key.upper() in STATE_CODES:
        return StateResult(key.upper())
    if key in STATES:
        return StateResult(STATES[key])
    # a near miss is offered as a suggestion, never applied
    match = process.extractOne(key, STATES.keys(), scorer=fuzz.ratio, score_cutoff=80)
    if match:
        return StateResult(None, f"closest match: {match[0].title()} ({STATES[match[0]]})")
    return StateResult(None, "not a US state name or code")
