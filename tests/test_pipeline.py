"""End-to-end tests: run the whole cycle in a temporary folder and check each output."""
import sqlite3
from pathlib import Path

import pytest
from openpyxl import load_workbook

from templatekit.pipeline import Pipeline
from templatekit.spec import ROW
from templatekit.workbook import compare_structure, snapshot, year_columns

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="module")
def run(tmp_path_factory):
    out = tmp_path_factory.mktemp("out")
    pl = Pipeline(ROOT / "config.yaml", out)
    summary = pl.run_all()
    return pl, summary, {r.country_code: r for r in pl.results}


def test_rollforward_adds_one_year_and_nothing_else(run):
    pl, _, _ = run
    old, new = snapshot(pl.master(pl.prev)), snapshot(pl.master(pl.sy))
    assert new["years"] == old["years"] + [pl.sy]
    assert all(new["cells"][k] == v for k, v in old["cells"].items())
    ws = load_workbook(pl.master(pl.sy))["Data"]
    col = year_columns(ws)[pl.sy]
    assert ws.cell(ROW["E_PRI_T"], col).value.startswith("=SUM(")


def test_translations_keep_structure(run):
    pl, _, _ = run
    base = snapshot(pl.master(pl.sy))
    for lang in pl.cfg["languages"]:
        other = snapshot(pl.master(pl.sy, lang.upper()))
        assert compare_structure(base, other) == []
        assert other["labels"] != base["labels"]


def test_prefill_writes_history_only(run):
    pl, _, _ = run
    ws = load_workbook(pl.p("templates", "prefilled", f"ARC_{pl.sy}.xlsx"))["Data"]
    cols = year_columns(ws)
    assert ws.cell(ROW["E_G1"], cols[pl.prev]).value is not None
    assert ws.cell(ROW["E_G1"], cols[pl.sy]).value is None


@pytest.mark.parametrize("code,rule,status", [
    ("ARC", None, "Accepted"),
    ("BOR", "SUM01", "Returned to country"),
    ("COR", "FRM01", "Returned to country"),
    ("DAL", "COH01", "Accepted with queries"),
    ("EST", "REQ01", "Returned to country"),
])
def test_validation_catches_each_scenario(run, code, rule, status):
    _, _, res = run
    r = res[code]
    assert r.status == status
    if rule:
        assert rule in {f.rule_id for f in r.flags}
    else:
        assert r.flags == []


def test_revision_detected(run):
    _, _, res = run
    assert "REV01" in {f.rule_id for f in res["DAL"].flags}


def test_release_database(run):
    pl, summary, _ = run
    with sqlite3.connect(summary["db"]) as con:
        countries = {r[0] for r in con.execute("SELECT DISTINCT country_code FROM observations")}
        sources = dict(con.execute("SELECT country_code, MIN(source) FROM observations GROUP BY 1"))
        n_rev = con.execute("SELECT COUNT(*) FROM revisions").fetchone()[0]
    assert countries == {"ARC", "DAL", "FEN"}  # returned submissions are not loaded
    assert sources["FEN"] == "mined"
    assert n_rev == 1
