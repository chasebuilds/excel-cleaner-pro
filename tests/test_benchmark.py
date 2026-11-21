"""The benchmark is the proof, so its own bookkeeping is tested too."""

from collections import Counter

from benchmark.generate import COLUMNS, canonical, generate, write_csv
from benchmark.run_benchmark import FIX_KINDS, FLAG_KINDS, score

from excel_cleaner_pro import clean_file


def test_generator_repeats_for_a_seed_and_differs_across_seeds():
    a, b, c = generate(300, 7), generate(300, 7), generate(300, 8)
    assert [r.cells for r in a] == [r.cells for r in b]
    assert [r.cells for r in a] != [r.cells for r in c]


def test_ledger_matches_the_file():
    records = generate(400, 3)
    kinds = Counter(k for r in records for k in r.planted.values())
    assert set(kinds) == set(FIX_KINDS) | set(FLAG_KINDS)
    for rec in records:
        if rec.row_kind == "empty_row":
            assert not "".join(rec.cells).strip()
            continue
        for i, name in enumerate(COLUMNS):
            damaged = rec.cells[i] != canonical(rec.expected[i])
            # a flag-only error keeps its damaged text as the expected value
            flag_only = rec.planted.get(name) in FLAG_KINDS
            assert damaged == (name in rec.planted and not flag_only), (name, rec.cells[i])
    rows = Counter(r.row_kind for r in records)
    assert rows["base"] == 400
    assert rows["exact_dup"] == 12
    assert rows["fuzzy_dup"] > 0
    assert rows["empty_row"] == 2


def test_scoring_end_to_end(tmp_path):
    records = generate(600, 11)
    path = tmp_path / "m.csv"
    write_csv(records, path)
    scores, facts = score(records, clean_file(path))
    assert facts["totals_balance"]
    assert not any(s.label == "Fuzzy duplicate row, removed" for s in scores)
    by_label = {s.label: s for s in scores}
    assert facts["accounted"]
    assert facts["false_removed"] == 0
    assert by_label["Extra whitespace"].recall == 1.0
    assert by_label["Exact duplicate row"].recall == 1.0
    assert by_label["Date readable two ways, column mixes orders"].precision == 1.0
    for s in scores:
        assert s.handled <= s.planted
        assert s.correct_actions <= s.actions


def test_scoring_notices_a_cleaner_that_does_nothing(tmp_path):
    from excel_cleaner_pro import load_config

    records = generate(300, 5)
    path = tmp_path / "m.csv"
    write_csv(records, path)
    off = {k: False for k in ("whitespace", "placeholders", "name_case", "email", "phone", "dates",
                              "numbers", "states")}  # fmt: skip
    cfg = load_config(overrides={"rules": off, "dedupe": {"exact": False, "fuzzy": False}})
    scores, facts = score(records, clean_file(path, cfg=cfg))
    fixes = [s for s in scores if s.expected_action == "fix" and s.planted]
    assert fixes
    assert all(s.recall == 0 for s in fixes)
    assert facts["wrong_cells_after"] > 0
