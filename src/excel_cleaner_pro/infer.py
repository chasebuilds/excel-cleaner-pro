"""Column type inference: config first, then the header, then the values."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

from .config import Config
from .rules import clean_state, is_valid_email, parse_date, parse_number

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


def _is_date_text(value: Any) -> bool:
    return isinstance(value, str) and parse_date(value).kind != "bad"


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
        if col_type == "date" and present:
            fits = _share(present, lambda v: isinstance(v, date | datetime) or _is_date_text(v))
            if fits < HEADER_SHARE:
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
    if all(isinstance(v, date | datetime) for v in present):
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
    if _share(texts, _is_date_text) >= VALUE_SHARE:
        return "date", "values"
    return "text", "default"
