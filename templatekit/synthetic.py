"""Synthetic data for the demo: fictional countries, a database extract, a mined
template and simulated country submissions (some with deliberate errors).

Nothing here comes from a real country or organisation.
"""
from __future__ import annotations

import random
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .spec import DATA_SHEET, GRADES, INPUTS, ROW
from .workbook import write_meta, year_columns


def country_series(code: str, years: list[int], seed: int) -> dict[tuple[str, int], int]:
    """Plausible, internally consistent primary-education series for one country.

    Noise is seeded per (country, variable, year), so a value never depends on
    how many years are requested.
    """
    rng = random.Random(f"{seed}-{code}")
    base_g1 = rng.randint(40_000, 600_000)
    g1_growth = rng.uniform(0.01, 0.03)
    survival = rng.uniform(0.93, 0.98)
    male_share = rng.uniform(0.49, 0.53)
    ger0 = rng.uniform(96, 110)
    ptr = rng.uniform(28, 45)
    pop_growth = rng.uniform(0.015, 0.03)

    out: dict[tuple[str, int], int] = {}
    y0 = years[0]
    for y in years:
        t = y - y0
        for g in range(1, 7):
            if t == 0:
                v = base_g1 * survival ** (g - 1)
            elif g == 1:
                v = base_g1 * (1 + g1_growth) ** t
            else:
                v = out[(f"E_G{g-1}", y - 1)] * survival
            noise = random.Random(f"{seed}-{code}-E_G{g}-{y}").uniform(0.995, 1.005)
            out[(f"E_G{g}", y)] = round(v * noise)
    total0 = sum(out[(g, y0)] for g in GRADES)
    pop0 = total0 / (ger0 / 100)
    for y in years:
        total = sum(out[(g, y)] for g in GRADES)
        m = round(total * male_share)
        out[("E_PRI_M", y)], out[("E_PRI_F", y)] = m, total - m
        out[("POP_PRI", y)] = round(pop0 * (1 + pop_growth) ** (y - y0))
        noise = random.Random(f"{seed}-{code}-TCH-{y}").uniform(0.99, 1.01)
        out[("TCH_PRI", y)] = round(total / ptr * noise)
    return out


def database_extract(countries: list[dict], years: list[int], seed: int) -> pd.DataFrame:
    """The 'previously published' database: inputs for all countries and years."""
    rows = []
    for c in countries:
        s = country_series(c["code"], years, seed)
        for (var, y), v in s.items():
            rows.append({"country_code": c["code"], "country_name": c["name"], "year": y,
                         "variable": var, "value": v})
    return pd.DataFrame(rows)


def _fill(ws, values: dict[tuple[str, int], float | None], years: dict[int, int]) -> None:
    for (code, y), v in values.items():
        if y in years and code in ROW:
            ws.cell(ROW[code], years[y], v)


def make_mined_template(master: Path, dst: Path, country: dict, years: list[int],
                        seed: int) -> Path:
    """A template compiled by the statistics office from secondary sources."""
    wb = load_workbook(master)
    ws = wb[DATA_SHEET]
    _fill(ws, country_series(country["code"], years, seed), year_columns(ws))
    write_meta(wb, Country_code=country["code"], Country_name=country["name"],
               Source="mined (secondary sources)")
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)
    return dst


def simulate_submission(prefilled: Path, dst: Path, country: dict, survey_year: int,
                        all_years: list[int], seed: int, scenario: str) -> Path:
    """Fill the survey-year column as a country would, injecting one error scenario."""
    shutil.copyfile(prefilled, dst)
    wb = load_workbook(dst)
    ws = wb[DATA_SHEET]
    cols = year_columns(ws)
    col = cols[survey_year]
    s = country_series(country["code"], all_years, seed)
    vals = {code: s[(code, survey_year)] for code in INPUTS}

    if scenario == "sex_mismatch":
        vals["E_PRI_F"] += 1_500
    elif scenario == "cohort_jump":
        prev_g3 = s[("E_G3", survey_year - 1)]
        old = vals["E_G4"]
        vals["E_G4"] = round(prev_g3 * 1.30)
        delta = vals["E_G4"] - old
        vals["E_PRI_M"] += delta // 2
        vals["E_PRI_F"] += delta - delta // 2
    elif scenario == "missing_and_jump":
        vals["TCH_PRI"] = None
        vals["POP_PRI"] = round(s[("POP_PRI", survey_year - 1)] * 1.35)

    for code, v in vals.items():
        ws.cell(ROW[code], col, v)

    if scenario == "formula_overwritten":
        total = sum(vals[g] for g in GRADES)
        ws.cell(ROW["E_PRI_T"], col, total + 800)  # typed over the SUM formula
    if scenario == "cohort_jump":
        # Also revise one historical value, as countries sometimes do.
        ws.cell(ROW["POP_PRI"], cols[survey_year - 2],
                round(s[("POP_PRI", survey_year - 2)] * 1.02))

    write_meta(wb, Source="country submission")
    wb.save(dst)
    return dst
