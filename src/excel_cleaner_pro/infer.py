"""Column type inference: config first, then the header, then the values."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

from .config import Config
from .rules import _MONTHS, clean_state, is_valid_email, parse_date, parse_number

_HEADER_HINTS: list[tuple[str, re.Pattern[str]]] = [
    ("email", re.compile(r"\be ?mail\b")),
    ("phone", re.compile(r"\b(phone|mobile|cell|tel|telephone|fax)\b")),
    ("date", re.compile(r"\b(date|dob|birthday|birthdate)\b")),
    ("state", re.compile(r"\b(state|province)\b")),
    ("id", re.compile(r"\b(id|sku|code|zip|zipcode|postal|ssn|ein|acct|no|number|num)\b")),
    ("name", re.compile(r"\b(name|contact|surname)\b")),
    (
        "currency",
        re.compile(
            r"\b(amount|price|total|cost|revenue|balance|salary|fee|paid|payment|value|spend"
            r"|sales|subtotal|tax)\b"
        ),
    ),
]
# title-casing a company, product or login would damage it (IBM, iPhone, jsmith)
_NOT_A_PERSON = re.compile(
    r"\b(company|business|org|organization|product|item|file|user|account|brand|store|city"
    r"|street|country|domain|host|sheet|table|column|field)\b"
)
_LEADING_ZERO = re.compile(r"^0\d+$")

# share of non-empty values that must fit before the values alone decide a type
VALUE_SHARE = 0.8
# share needed when the header already points the same way
HEADER_SHARE = 0.5


def _share(values: list[Any], test) -> float:
    return sum(1 for v in values if test(v)) / len(values) if values else 0.0


# Written like a date even when the date itself is wrong ("4/31/2024", "Aug 2024"). Counts
# toward finding a date column so the bad cells get flagged instead of the column being skipped.
# Needs a year, so "12/5", "2024-001", ZIPs, phones and amounts never qualify.
_DATE_SHAPED = re.compile(
    r"^(\d{1,2}[-/.]\d{1,2}[-/.](\d{4}|\d{2})|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"
    r"|[A-Za-z]{3,9}\.?,?\s+\d{4})$"
)
# with no date-like header, this share of cells must be real, readable dates
DATE_VALID_SHARE = 0.6


def _is_date_text(value: Any) -> bool:
    return isinstance(value, str) and parse_date(value).kind != "bad"


def _is_date_value(value: Any) -> bool:
    return isinstance(value, date | datetime) or _is_date_text(value)


def _is_date_shaped(value: Any) -> bool:
    if _is_date_value(value):
        return True
    if not isinstance(value, str) or not (m := _DATE_SHAPED.match(value.strip())):
        return False
    word = re.match(r"[A-Za-z]+", m.group(0))
    return word is None or word.group(0).lower() in _MONTHS


def _looks_like_date_column(present: list[Any]) -> bool:
    """Most cells read as dates, counting real Excel dates and date-shaped typos."""
    return (
        _share(present, _is_date_value) >= DATE_VALID_SHARE
        and _share(present, _is_date_shaped) >= VALUE_SHARE
    )


def _is_number_text(value: Any) -> bool:
    return isinstance(value, str) and parse_number(value).kind != "bad"


def infer_type(header: str, values: list[Any], cfg: Config) -> tuple[str, str]:
    """Return (type, source) where source says what decided it: config, header, values, default."""
    if header in cfg.columns:
        return cfg.columns[header], "config"

    present = [v for v in values if v is not None]
    texts = [v for v in present if isinstance(v, str)]
    words = re.sub(r"[^a-z0-9]+", " ", re.sub(r"([a-z])([A-Z])", r"\1 \2", header).lower()).strip()

    for col_type, pattern in _HEADER_HINTS:
        if not pattern.search(words):
            continue
        if col_type == "name" and _NOT_A_PERSON.search(words):
            continue
        if col_type == "date" and present and _share(present, _is_date_value) < HEADER_SHARE:
            continue
        if col_type == "state" and _share(texts, lambda v: clean_state(v).value) < HEADER_SHARE:
            continue
        if col_type == "currency" and present:
            fits = _share(present, lambda v: isinstance(v, int | float) or _is_number_text(v))
            if fits < HEADER_SHARE:
                continue
        return col_type, "header"

    if not present:
        return "text", "default"
    # real Excel dates and typed dates are counted together, so a mixed column is still dates
    if _looks_like_date_column(present):
        return "date", "values"
    if all(isinstance(v, int | float) and not isinstance(v, bool) for v in present):
        return "number", "values"
    if not texts:
        return "text", "default"

    if _share(texts, is_valid_email) >= VALUE_SHARE:
        return "email", "values"
    # a zero-padded code (00123) must stay text or the padding is lost
    if any(_LEADING_ZERO.match(v) for v in texts):
        return "id", "values"
    if _share(texts, _is_number_text) >= VALUE_SHARE:
        has_currency = any(parse_number(v).currency for v in texts)
        return ("currency" if has_currency else "number"), "values"
    return "text", "default"
