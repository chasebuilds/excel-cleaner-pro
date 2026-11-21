"""Seeded generator: a clean customer table, then a messy copy with a counted set of errors.

The generator does not import the cleaner. Its idea of "correct" is written here, on its
own, so the benchmark cannot pass by agreeing with itself.

Each record keeps three things:
  cells     what goes into the messy file
  expected  what a perfect cleaner should leave in each cell
  planted   {column: error kind} for every cell that was damaged on purpose
"""

from __future__ import annotations

import argparse
import csv
import random
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Any

COLUMNS = [
    "customer_id", "full_name", "email", "phone", "city", "state",
    "signup_date", "last_order_date", "lifetime_value", "orders",
]  # fmt: skip

FIRST = [
    "James", "Mary", "Robert", "Patricia", "John", "Jennifer", "Michael", "Linda", "David",
    "Elizabeth", "William", "Barbara", "Richard", "Susan", "Joseph", "Jessica", "Thomas", "Sarah",
    "Charles", "Karen", "Daniel", "Nancy", "Matthew", "Lisa", "Anthony", "Betty", "Mark",
    "Margaret", "Donald", "Sandra", "Steven", "Ashley", "Paul", "Kimberly", "Andrew", "Emily",
    "Joshua", "Donna", "Kenneth", "Michelle", "Kevin", "Carol", "Brian", "Amanda", "George",
    "Melissa", "Timothy", "Deborah", "Ronald", "Stephanie", "Jason", "Rebecca", "Ryan", "Laura",
    "Jacob", "Sharon", "Gary", "Cynthia", "Nicholas", "Kathleen", "Eric", "Amy", "Jonathan",
    "Angela", "Stephen", "Anna", "Larry", "Ruth", "Justin", "Brenda", "Scott", "Pamela",
    "Brandon", "Nicole", "Benjamin", "Katherine", "Samuel", "Christine", "Gregory", "Rachel",
    "Mei", "Li", "Wei", "Aiko", "Priya", "Arjun", "Omar", "Layla", "Diego", "Sofia",
]  # fmt: skip
LAST = [
    "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez",
    "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor",
    "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez",
    "Clark", "Ramirez", "Lewis", "Robinson", "Walker", "Young", "Allen", "King", "Wright",
    "Scott", "Torres", "Nguyen", "Hill", "Flores", "Green", "Adams", "Nelson", "Baker", "Hall",
    "Rivera", "Campbell", "Mitchell", "Carter", "Roberts", "Gomez", "Phillips", "Evans",
    "Turner", "Diaz", "Parker", "Cruz", "Edwards", "Collins", "Reyes", "Stewart", "Morris",
    "Wu", "Kim", "Patel", "Khan", "O'Brien", "O'Connor", "McCarthy", "McDonald",
    "Smith-Jones", "Garcia-Lopez",
]  # fmt: skip
# Surnames whose capitals cannot be rebuilt from ALL CAPS by any simple rule.
# They are here on purpose, at a realistic rate, so the name-case score is honest.
LAST_HARD = [
    "de la Cruz", "van Dyke", "MacDonald", "DeAngelo", "LeBlanc", "von Braun", "McDonald III",
    "DiMaggio", "St. James", "del Rio",
]  # fmt: skip
HARD_SHARE = 0.04

STATES = {
    "Alabama": "AL", "Arizona": "AZ", "California": "CA", "Colorado": "CO", "Florida": "FL",
    "Georgia": "GA", "Illinois": "IL", "Indiana": "IN", "Kentucky": "KY", "Massachusetts": "MA",
    "Michigan": "MI", "Minnesota": "MN", "Missouri": "MO", "New York": "NY",
    "North Carolina": "NC", "Ohio": "OH", "Oregon": "OR", "Pennsylvania": "PA",
    "Tennessee": "TN", "Texas": "TX", "Virginia": "VA", "Washington": "WA", "Wisconsin": "WI",
}  # fmt: skip
CODE_TO_STATE = {code: name for name, code in STATES.items()}
CITIES = [
    ("Birmingham", "AL"), ("Phoenix", "AZ"), ("San Diego", "CA"), ("Sacramento", "CA"),
    ("Denver", "CO"), ("Tampa", "FL"), ("Orlando", "FL"), ("Atlanta", "GA"), ("Chicago", "IL"),
    ("Indianapolis", "IN"), ("Louisville", "KY"), ("Lexington", "KY"), ("Boston", "MA"),
    ("Detroit", "MI"), ("Minneapolis", "MN"), ("Kansas City", "MO"), ("Buffalo", "NY"),
    ("Charlotte", "NC"), ("Columbus", "OH"), ("Portland", "OR"), ("Pittsburgh", "PA"),
    ("Nashville", "TN"), ("Austin", "TX"), ("Houston", "TX"), ("Richmond", "VA"),
    ("Seattle", "WA"), ("Madison", "WI"),
]  # fmt: skip
AREA_CODES = ["205", "303", "312", "404", "415", "502", "512", "602", "615", "617", "704", "813"]
DOMAINS = ["example.com", "example.org", "example.net"]
MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June", "July", "August", "September",
    "October", "November", "December",
]  # fmt: skip
PLACEHOLDERS = ["N/A", "n/a", "-", "--", "null", "NULL", "none", "NA"]

# share of eligible cells damaged, per error kind
RATES = {
    "whitespace": 0.04,
    "placeholder": 0.35,  # of the cells that are truly empty
    "name_case": 0.06,
    "email_case": 0.05,
    "email_invalid": 0.02,
    "phone_format": 0.12,
    "phone_invalid": 0.02,
    "date": 0.25,
    "date_unreadable": 0.01,
    "currency_format": 0.10,
    "number_ambiguous": 0.10,  # of the whole-dollar amounts of 1,000 or more
    "state_name": 0.08,
    "state_unknown": 0.01,
}
ROW_RATES = {"exact_dup": 0.03, "fuzzy_dup": 0.03, "empty_row": 0.005}


@dataclass
class Record:
    entity: int  # index of the real customer this row describes
    cells: list[str]
    expected: list[Any]
    planted: dict[str, str] = field(default_factory=dict)
    row_kind: str = "base"  # base, exact_dup, fuzzy_dup, empty_row
    fuzzy_edit: str = ""


def build_truth(n: int, rng: random.Random) -> list[list[Any]]:
    rows, emails = [], set()
    for i in range(n):
        first = rng.choice(FIRST)
        last = rng.choice(LAST_HARD) if rng.random() < HARD_SHARE else rng.choice(LAST)
        slug = "".join(ch for ch in f"{first}.{last}".lower() if ch.isalnum() or ch == ".")
        while True:
            email = f"{slug.strip('.')}{rng.randint(1, 999)}@{rng.choice(DOMAINS)}"
            if email not in emails:
                emails.add(email)
                break
        city, state = rng.choice(CITIES)
        signup = date(2019, 1, 1) + timedelta(days=rng.randint(0, 2400))
        last_order = signup + timedelta(days=rng.randint(0, 400))
        value = (
            float(rng.randint(1, 9)) * 500
            if rng.random() < 0.2
            else rng.randint(2000, 900000) / 100
        )
        rows.append(
            [
                f"C{i + 1:05d}",
                f"{first} {last}",
                email if rng.random() > 0.08 else None,
                f"({rng.choice(AREA_CODES)}) 555-01{rng.randint(0, 99):02d}"
                if rng.random() > 0.10
                else None,
                city if rng.random() > 0.05 else None,
                state,
                signup.isoformat(),
                last_order.isoformat() if rng.random() > 0.15 else None,
                value,
                rng.randint(1, 40),
            ]
        )
    return rows


def canonical(value: Any) -> str:
    if value is None:
        return ""
    return f"{value:.2f}" if isinstance(value, float) else str(value)


# ---------------------------------------------------------------- cell damage


def _whitespace(text: str, rng: random.Random) -> str:
    options = [f"  {text}", f"{text}  ", f" {text} ", f"{text}\t", f"{text} "]
    if " " in text:
        options.append(text.replace(" ", "   ", 1))
    return rng.choice(options)


def _email_case(text: str, rng: random.Random) -> str:
    local, _, domain = text.partition("@")
    return rng.choice(
        [text.upper(), f"{local.title()}@{domain.title()}", f"{local}@{domain.upper()}"]
    )


def _email_invalid(text: str, rng: random.Random) -> str:
    local, _, domain = text.partition("@")
    return rng.choice(
        [
            f"{local}{domain}",
            f"{local} at {domain}",
            f"{local}@{domain.rsplit('.', 1)[0]}",
            f"{local}@@{domain}",
            f"{local}@{domain.replace('.', ',')}",
            f"{local}@",
        ]
    )


def _phone_format(text: str, rng: random.Random) -> str:
    d = "".join(ch for ch in text if ch.isdigit())
    a, b, c = d[:3], d[3:6], d[6:]
    return rng.choice(
        [f"{a}-{b}-{c}", f"{a}.{b}.{c}", d, f"+1 {a} {b} {c}", f"1-{a}-{b}-{c}",
         f"({a}){b}-{c}", f"{a} {b} {c}", f"+1{d}"]
    )  # fmt: skip


def _phone_invalid(text: str, rng: random.Random) -> str:
    d = "".join(ch for ch in text if ch.isdigit())
    return rng.choice(
        [f"{d[3:6]}-{d[6:]}", f"({d[:3]}) {d[3:6]}-{d[6:9]}", f"{d[:3]}-{d[3:6]}-{d[6:]}99"]
    )


def _date(iso: str, mixed_order: bool, rng: random.Random) -> tuple[str, str]:
    """Return (text, kind). mixed_order columns hold day-first and month-first dates together."""
    y, m, d = (int(p) for p in iso.split("-"))
    month = MONTH_NAMES[m - 1]
    style = rng.choice(["numeric", "numeric", "numeric", "long", "short", "dmy_text", "slash_ymd"])
    if style == "long":
        return f"{month} {d}, {y}", "date_format"
    if style == "short":
        return f"{month[:3]} {d}, {y}", "date_format"
    if style == "dmy_text":
        return rng.choice([f"{d}-{month[:3]}-{y}", f"{d:02d} {month} {y}"]), "date_format"
    if style == "slash_ymd":
        return f"{y}/{m:02d}/{d:02d}", "date_format"
    day_first = mixed_order and rng.random() < 0.5
    a, b = (d, m) if day_first else (m, d)
    text = f"{a:02d}/{b:02d}/{y}" if rng.random() < 0.6 else f"{a}/{b}/{y}"
    if d <= 12 and d != m:
        # readable both ways: the right move depends on what the rest of the column proves
        return text, ("date_ambiguous" if mixed_order else "date_inferred_order")
    return text, "date_format"


def _date_unreadable(iso: str, rng: random.Random) -> str:
    y, m, _ = iso.split("-")
    return rng.choice([f"{m}/{y}", f"Q{(int(m) + 2) // 3} {y}", f"{y}-{m}-45", f"02/30/{y}",
                       f"early {y}", "see notes"])  # fmt: skip


def _currency(value: float, rng: random.Random) -> str:
    us = f"{value:,.2f}"
    eu = us.replace(",", "X").replace(".", ",").replace("X", ".")
    return rng.choice(
        [f"${us}", f"${us}", us, f"USD {value:.2f}", f"$ {us}", f"${value:.2f}", eu, f"€{eu}",
         f"{value:.2f}".replace(".", ",")]
    )  # fmt: skip


def _number_ambiguous(value: float, rng: random.Random) -> str:
    grouped = f"{int(value):,}"
    return rng.choice([grouped, grouped.replace(",", "."), f"${grouped}"])


def _state_name(code: str, rng: random.Random) -> str:
    name = CODE_TO_STATE[code]
    return rng.choice([name, name, name.lower(), name.upper(), code.lower(), code.title()])


def _state_unknown(code: str, rng: random.Random) -> str:
    name = CODE_TO_STATE[code]
    while True:
        i = rng.randrange(1, len(name))
        typo = name[:i] + name[i + 1 :]
        if typo not in STATES and typo.strip() == typo and "  " not in typo:
            return typo


def _typo(text: str, rng: random.Random) -> str:
    """One letter dropped, doubled or swapped with its neighbour."""
    spots = [i for i in range(1, len(text) - 1) if text[i].isalpha() and text[i + 1].isalpha()]
    i = rng.choice(spots)
    kind = rng.choice(["drop", "double", "swap"])
    if kind == "drop":
        return text[:i] + text[i + 1 :]
    if kind == "double":
        return text[:i] + text[i] + text[i:]
    if text[i] == text[i + 1]:
        return text[:i] + text[i + 1 :]
    return text[:i] + text[i + 1] + text[i] + text[i + 2 :]


# ---------------------------------------------------------------- assembly


def _pick(rng: random.Random, pool: list, rate: float) -> list:
    return rng.sample(pool, round(len(pool) * rate)) if pool else []


def make_messy(truth: list[list[Any]], rng: random.Random) -> list[Record]:
    col = {name: i for i, name in enumerate(COLUMNS)}
    records = [Record(i, [canonical(v) for v in row], list(row)) for i, row in enumerate(truth)]
    free = {(r, c) for r in range(len(records)) for c in range(len(COLUMNS))}

    def plant(kind: str, columns: list[str], damage, eligible=lambda value: value is not None,
              rate_key: str | None = None, fixable: bool = True) -> None:  # fmt: skip
        pool = sorted(
            (r, col[name])
            for name in columns
            for r in range(len(records))
            if (r, col[name]) in free and eligible(truth[r][col[name]])
        )
        for r, c in _pick(rng, pool, RATES[rate_key or kind]):
            out = damage(truth[r][c], rng)
            text, actual_kind = out if isinstance(out, tuple) else (out, kind)
            if text == records[r].cells[c]:
                continue
            free.discard((r, c))
            records[r].cells[c] = text
            records[r].planted[COLUMNS[c]] = actual_kind
            if not fixable or actual_kind == "date_ambiguous":
                # nothing can recover the true value, so the right outcome is: untouched and flagged
                records[r].expected[c] = text

    plant("placeholder", ["email", "phone", "city", "last_order_date"],
          lambda _v, g: g.choice(PLACEHOLDERS), eligible=lambda v: v is None)  # fmt: skip
    plant("email_invalid", ["email"], _email_invalid, fixable=False)
    plant("phone_invalid", ["phone"], _phone_invalid, fixable=False)
    plant("state_unknown", ["state"], _state_unknown, fixable=False)
    plant("number_ambiguous", ["lifetime_value"], _number_ambiguous, fixable=False,
          eligible=lambda v: v >= 1000 and float(v).is_integer())  # fmt: skip
    plant("date_unreadable", ["signup_date", "last_order_date"], _date_unreadable, fixable=False)
    plant("name_case", ["full_name"],
          lambda v, g: v.upper() if g.random() < 0.6 else v.lower())  # fmt: skip
    plant("email_case", ["email"], _email_case)
    plant("phone_format", ["phone"], _phone_format)
    plant("date", ["signup_date"], lambda v, g: _date(v, False, g))
    plant("date", ["last_order_date"], lambda v, g: _date(v, True, g))
    plant("currency_format", ["lifetime_value"], _currency)
    plant("state_name", ["state"], _state_name)
    plant("whitespace", ["full_name", "email", "city", "state"], _whitespace)

    n = len(records)
    with_email = [r for r in records if r.expected[col["email"]] and "email" not in r.planted]
    extras: list[tuple[int, Record]] = []
    for origin in _pick(rng, records, ROW_RATES["exact_dup"]):
        dup = Record(origin.entity, list(origin.cells), list(origin.expected),
                     dict(origin.planted), "exact_dup")  # fmt: skip
        extras.append((origin.entity, dup))
    for k, origin in enumerate(_pick(rng, with_email, ROW_RATES["fuzzy_dup"])):
        # the same customer typed in again: new id, one slip in the name or the email
        cells = [canonical(v) for v in truth[origin.entity]]
        cells[col["customer_id"]] = f"C{n + k + 1:05d}"
        edit = rng.choice(["name_typo", "name_typo", "email_typo"])
        if edit == "name_typo":
            cells[col["full_name"]] = _typo(cells[col["full_name"]], rng)
        else:
            local, _, domain = cells[col["email"]].partition("@")
            cells[col["email"]] = f"{_typo(local, rng)}@{domain}"
        expected = list(truth[origin.entity])
        for name in ("customer_id", "full_name", "email"):
            expected[col[name]] = cells[col[name]]
        extras.append(
            (origin.entity, Record(origin.entity, cells, expected, {}, "fuzzy_dup", edit))
        )
    for _ in range(round(n * ROW_RATES["empty_row"])):
        blank = [rng.choice(["", "", " "]) for _ in COLUMNS]
        extras.append((rng.randrange(n), Record(-1, blank, [None] * len(COLUMNS), {}, "empty_row")))

    # a copy always lands somewhere after the row it copies
    slots: dict[int, list[Record]] = {}
    for origin_index, extra in extras:
        slots.setdefault(rng.randint(origin_index, n - 1), []).append(extra)
    out: list[Record] = []
    for i, record in enumerate(records):
        out.append(record)
        out.extend(slots.get(i, []))
    return out


def generate(rows: int, seed: int) -> list[Record]:
    rng = random.Random(seed)
    return make_messy(build_truth(rows, rng), rng)


def write_csv(records: list[Record], path: str | Path) -> None:
    with open(path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(COLUMNS)
        writer.writerows(r.cells for r in records)


def main() -> None:
    parser = argparse.ArgumentParser(description="Write a seeded messy customer CSV.")
    parser.add_argument("--rows", type=int, default=5000, help="clean rows before damage")
    parser.add_argument("--seed", type=int, default=42)
    parser.add_argument("--out", default="benchmark/messy_customers.csv")
    args = parser.parse_args()
    records = generate(args.rows, args.seed)
    write_csv(records, args.out)
    print(f"wrote {len(records)} rows to {args.out}")


if __name__ == "__main__":
    main()
