"""Consolidation: validated country templates + mined templates -> one release database.

Precedence: country-reported data wins over mined data for the same
(country, year, variable). Submissions returned to the country are logged but
not loaded. A revision report compares the release with the previous one.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

from .validate import Result

PRECEDENCE = {"country": 0, "mined": 1}


def flatten(res: Result, source: str) -> pd.DataFrame:
    """One row per non-blank (country, year, variable) value in a validated template."""
    rows = [{"country_code": res.country_code, "country_name": res.country_name, "year": y,
             "variable": c, "value": v, "source": source, "status": res.status,
             "submission_file": res.file, "sha256": res.sha256}
            for (c, y), v in res.values.items() if v is not None]
    return pd.DataFrame(rows)


def consolidate(country: list[Result], mined: list[Result], previous: pd.DataFrame,
                out_dir: Path, release: str) -> dict:
    """Build the release database and CSV exports; return a summary of what was loaded.

    Returned submissions are excluded, country data takes precedence over mined data,
    and values that differ from `previous` are written to the revisions table.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    loaded = [r for r in country if r.status != "Returned to country"]
    frames = [flatten(r, "country") for r in loaded] + [flatten(r, "mined") for r in mined]
    obs = pd.concat(frames, ignore_index=True)
    obs["_p"] = obs.source.map(PRECEDENCE)
    obs = (obs.sort_values("_p").drop_duplicates(["country_code", "year", "variable"])
              .drop(columns="_p").sort_values(["country_code", "variable", "year"]))
    obs["release"] = release

    subs = pd.DataFrame([{
        "file": r.file, "sha256": r.sha256, "country_code": r.country_code,
        "source": src, "status": r.status, "blocking": r.count("blocking"),
        "to_confirm": r.count("to_confirm"), "info": r.count("info"),
        "loaded": src == "mined" or r.status != "Returned to country",
    } for src, rs in (("country", country), ("mined", mined)) for r in rs])

    merged = obs.merge(previous[["country_code", "year", "variable", "value"]],
                       on=["country_code", "year", "variable"], how="inner",
                       suffixes=("", "_previous"))
    rev = merged[(merged.value - merged.value_previous).abs() > 1e-9].copy()
    rev["change_pct"] = ((rev.value - rev.value_previous) / rev.value_previous * 100).round(2)
    rev = rev[["country_code", "year", "variable", "value_previous", "value", "change_pct",
               "source"]]

    db_path = out_dir / f"release_{release}.sqlite"
    db_path.unlink(missing_ok=True)
    with sqlite3.connect(db_path) as con:
        obs.to_sql("observations", con, index=False)
        subs.to_sql("submissions", con, index=False)
        rev.to_sql("revisions", con, index=False)
        con.execute("CREATE INDEX ix_obs ON observations(country_code, variable, year)")
        pd.DataFrame([{"release": release, "built": datetime.now().isoformat(timespec="seconds"),
                       "rows": len(obs)}]).to_sql("release_info", con, index=False)
    obs.to_csv(out_dir / f"observations_{release}.csv", index=False)
    rev.to_csv(out_dir / f"revisions_{release}.csv", index=False)
    subs.to_csv(out_dir / f"submissions_{release}.csv", index=False)
    return {"db": db_path, "rows": len(obs), "revisions": len(rev),
            "loaded": int(subs.loaded.sum()), "excluded": int((~subs.loaded).sum())}
