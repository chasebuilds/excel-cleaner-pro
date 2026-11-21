"""Config loading. Defaults live in default_config.toml so the documented file is the real one."""

from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, field
from importlib import resources
from pathlib import Path
from typing import Any

COLUMN_TYPES = ("text", "id", "name", "email", "phone", "date", "number", "currency", "state")
DATE_ORDERS = ("infer", "flag", "mdy", "dmy")
PHONE_FORMATS = ("national", "dashes", "e164")


class ConfigError(ValueError):
    pass


@dataclass
class Config:
    rules: dict[str, bool] = field(default_factory=dict)
    placeholder_tokens: frozenset[str] = frozenset()
    columns: dict[str, str] = field(default_factory=dict)
    phone_format: str = "national"
    date_order: str = "infer"
    date_infer_min_evidence: int = 3
    two_digit_year_pivot: int = 50
    dedupe_exact: bool = True
    dedupe_fuzzy: bool = True
    fuzzy_keys: list[str] = field(default_factory=list)
    remove_threshold: float = 95.0
    review_threshold: float = 85.0
    fuzzy_action: str = "review"
    source_row_column: bool = True
    header_row: int = 0
    validate: dict[str, Any] = field(default_factory=dict)
    raw: dict[str, Any] = field(default_factory=dict, compare=False, repr=False)


def default_config_text() -> str:
    return resources.files("excel_cleaner_pro").joinpath("default_config.toml").read_text("utf-8")


def _merge(base: dict[str, Any], override: dict[str, Any], path: str = "") -> dict[str, Any]:
    out = dict(base)
    for key, value in override.items():
        # these tables are keyed by the user's own header names, so any key is allowed
        if key not in base and path not in ("columns", "allowed", "range", "pattern"):
            raise ConfigError(f"unknown config key: {path + '.' if path else ''}{key}")
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            out[key] = _merge(base[key], value, key)
        else:
            out[key] = value
    return out


def load_config(path: str | Path | None = None, overrides: dict[str, Any] | None = None) -> Config:
    raw = tomllib.loads(default_config_text())
    if path is not None:
        data = Path(path).read_bytes()
        # PowerShell's > redirect writes UTF-16, and Notepad may add a UTF-8 byte order mark
        encoding = "utf-16" if data[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig"
        try:
            raw = _merge(raw, tomllib.loads(data.decode(encoding)))
        except (tomllib.TOMLDecodeError, UnicodeDecodeError) as exc:
            raise ConfigError(f"{path}: not valid TOML ({exc})") from exc
    if overrides:
        raw = _merge(raw, overrides)

    cfg = Config(
        rules={k: bool(v) for k, v in raw["rules"].items()},
        placeholder_tokens=frozenset(
            str(t).strip().casefold() for t in raw["placeholders"]["tokens"]
        ),
        columns={str(k): str(v).lower() for k, v in raw["columns"].items()},
        phone_format=raw["phone"]["format"],
        date_order=raw["dates"]["order"],
        date_infer_min_evidence=int(raw["dates"]["infer_min_evidence"]),
        two_digit_year_pivot=int(raw["dates"]["two_digit_year_pivot"]),
        dedupe_exact=bool(raw["dedupe"]["exact"]),
        dedupe_fuzzy=bool(raw["dedupe"]["fuzzy"]),
        fuzzy_keys=[str(k) for k in raw["dedupe"]["fuzzy_keys"]],
        remove_threshold=float(raw["dedupe"]["remove_threshold"]),
        review_threshold=float(raw["dedupe"]["review_threshold"]),
        fuzzy_action=str(raw["dedupe"]["fuzzy_action"]),
        source_row_column=bool(raw["output"]["source_row_column"]),
        header_row=int(raw["input"]["header_row"]),
        validate=raw["validate"],
        raw=raw,
    )
    for name, col_type in cfg.columns.items():
        if col_type not in COLUMN_TYPES:
            raise ConfigError(
                f"columns.{name}: unknown type {col_type!r}, use one of {COLUMN_TYPES}"
            )
    if cfg.date_order not in DATE_ORDERS:
        raise ConfigError(f"dates.order must be one of {DATE_ORDERS}")
    if cfg.fuzzy_action not in ("review", "remove"):
        raise ConfigError('dedupe.fuzzy_action must be "review" or "remove"')
    if cfg.header_row < 0:
        raise ConfigError("input.header_row must be 0 (find it) or a row number")
    for column, bounds in cfg.validate["range"].items():
        if not (isinstance(bounds, list) and len(bounds) == 2):
            raise ConfigError(f"validate.range.{column} must be [min, max]")
    for column, pattern in cfg.validate["pattern"].items():
        try:
            re.compile(pattern)
        except re.error as exc:
            raise ConfigError(f"validate.pattern.{column}: {exc}") from exc
    if cfg.phone_format not in PHONE_FORMATS:
        raise ConfigError(f"phone.format must be one of {PHONE_FORMATS}")
    if not 0 < cfg.review_threshold <= cfg.remove_threshold <= 100:
        raise ConfigError(
            "dedupe thresholds must satisfy 0 < review_threshold <= remove_threshold <= 100"
        )
    return cfg
