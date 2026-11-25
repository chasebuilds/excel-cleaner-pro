window.DEMO = {
  repo: "excel-cleaner-pro",
  title: "excel-cleaner-pro",
  tagline: "Drop in a messy CSV or Excel file. Get back a clean workbook plus a log of every change, every removed row and every value it would not guess at.",
  pkg: "excel_cleaner_pro",
  packages: ["openpyxl"],
  inputs: [
    { id: "input", label: "Messy CSV or XLSX", accept: ".csv,.xlsx,.xlsm,.txt,.tsv", sample: "samples/messy_customers.csv" },
  ],
  options: [
    { id: "date_order", label: "Dates like 03/04/2024", type: "select", value: "",
      choices: [["", "infer from the column"], ["flag", "flag, do not guess"], ["mdy", "month first"], ["dmy", "day first"]] },
    { id: "no_fuzzy", label: "Skip fuzzy duplicate matching", type: "checkbox", value: false },
    { id: "fuzzy_remove", label: "Remove fuzzy duplicates (default only lists them)", type: "checkbox", value: false },
  ],
  sampleNote: 'The sample is <a href="samples/messy_customers.csv">messy_customers.csv</a>: 42 synthetic customer rows with mixed-case names, odd phone and date formats, euro amounts, placeholders and duplicates.<br>Fuzzy duplicate matching uses rapidfuzz\'s own pure-Python implementation here, because rapidfuzz has no WebAssembly build. Same algorithms and scores, just slower than on a desktop.',
  python: `
import os
os.environ["RAPIDFUZZ_IMPL"] = "python"  # the C extension has no Pyodide build
from pathlib import Path
from excel_cleaner_pro.cli import main as cli_main

ORDER = ["Cleaned", "Changes", "Issues", "Removed"]

def run(paths, opts):
    src = paths["input"][0]
    out = "/work/" + Path(src).stem + "_cleaned.xlsx"
    argv = [src, "-o", out]
    if opts.get("date_order"):
        argv += ["--date-order", opts["date_order"]]
    if opts.get("no_fuzzy"):
        argv.append("--no-fuzzy")
    if opts.get("fuzzy_remove"):
        argv.append("--fuzzy-remove")
    code, log = run_cli(cli_main, argv)
    if code != 0 or not Path(out).exists():
        return {"ok": False, "log": log, "sheets": [], "download": None}
    sheets = sheets_from_xlsx(out)
    rank = lambda s: next((i for i, p in enumerate(ORDER) if s["name"].startswith(p)), len(ORDER))
    sheets.sort(key=rank)
    return {"ok": True, "log": log, "sheets": sheets, "download": out}
`,
};
