"""Consolidation edge cases."""
import sqlite3

import pandas as pd

from templatekit.consolidate import consolidate
from templatekit.validate import Flag, Result


def returned(code: str) -> Result:
    res = Result(file=f"{code}_2025.xlsx", sha256="0" * 64, country_code=code,
                 country_name=code, survey_year=2025, version="2025.1", language="EN",
                 source="", values={("E_G1", 2025): 100.0})
    res.flags.append(Flag("STR01", "blocking", "structure changed"))
    return res


def test_release_builds_when_every_template_is_returned(tmp_path):
    previous = pd.DataFrame([{"country_code": "ARC", "country_name": "Arcadia", "year": 2024,
                              "variable": "E_G1", "value": 90.0}])
    summary = consolidate([returned("ARC"), returned("BOR")], [returned("FEN")], previous,
                          tmp_path, "2025")
    assert (summary["rows"], summary["revisions"], summary["loaded"], summary["excluded"]) == \
        (0, 0, 0, 3)
    with sqlite3.connect(summary["db"]) as con:
        assert con.execute("SELECT COUNT(*) FROM observations").fetchone() == (0,)
        assert con.execute("SELECT SUM(loaded) FROM submissions").fetchone() == (0,)
        cols = [r[1] for r in con.execute("PRAGMA table_info(observations)")]
    assert {"country_code", "year", "variable", "value", "source", "release"} <= set(cols)
