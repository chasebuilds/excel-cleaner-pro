"""The two calls most users need: clean_file and clean_tables."""

from __future__ import annotations

import hashlib
import platform
from datetime import UTC, datetime
from pathlib import Path

from .cleaner import clean_table
from .config import Config, load_config
from .io import write_workbook
from .model import CleanResult, Table
from .reader import read_tables
from .version import __version__


def clean_tables(tables: list[Table], cfg: Config | None = None, source: str = "") -> CleanResult:
    cfg = cfg or load_config()
    return CleanResult(source, [clean_table(t, cfg) for t in tables])


def clean_file(
    input_path: str | Path,
    output_path: str | Path | None = None,
    cfg: Config | None = None,
) -> CleanResult:
    """Clean a CSV or XLSX file. Writes the audited workbook when output_path is given."""
    cfg = cfg or load_config()
    input_path = Path(input_path)
    if output_path is not None and Path(output_path).resolve() == input_path.resolve():
        raise ValueError("output path is the input file: the original is never overwritten")
    data = input_path.read_bytes()
    tables = read_tables(input_path, cfg)
    result = clean_tables(tables, cfg, source=input_path.name)
    result.manifest = {
        "Input file": input_path.name,
        "Input SHA-256": hashlib.sha256(data).hexdigest(),
        "Input size (bytes)": len(data),
        **(tables[0].meta if tables else {}),
        "Tool": f"excel-cleaner-pro {__version__}",
        "Python": platform.python_version(),
        "Run at (UTC)": datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S"),
    }
    if output_path is not None:
        write_workbook(result, output_path, cfg)
    return result
