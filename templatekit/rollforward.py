"""Roll a template forward by one survey year, prefill it, and translate it."""
from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .spec import DATA_SHEET, HEADER_ROW, INPUTS, LABEL_COL, ROW, VARIABLES
from .workbook import (
    apply_validation, diff_rollforward, read_meta, snapshot, write_meta, write_year_column,
    year_columns,
)


def bump_version(version: str, new_year: int) -> str:
    """Template version for a new cycle: '<year>.1' (the old version is not used yet)."""
    return f"{new_year}.1"


def roll_forward(src: Path, dst: Path, new_year: int | None = None) -> list[dict]:
    """Add the next year column, extend formulas and validation, update metadata.

    Returns the classified diff between src and dst (see diff_rollforward).
    """
    wb = load_workbook(src)
    ws = wb[DATA_SHEET]
    cols = year_columns(ws)
    new_year = new_year or max(cols) + 1
    if new_year in cols:
        raise ValueError(f"{src.name} already contains {new_year}")
    # Previous survey-year input cells lose their highlight.
    for col in cols.values():
        for code in INPUTS:
            ws.cell(ROW[code], col).fill = PatternFill(fill_type=None)
    new_col = max(cols.values()) + 1
    write_year_column(ws, new_year, new_col, survey_year=new_year)
    apply_validation(ws, list(cols.values()) + [new_col])
    meta = read_meta(wb)
    write_meta(wb, Survey_year=new_year,
               Version=bump_version(str(meta.get("Version", "")), new_year))
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)
    return diff_rollforward(snapshot(src), snapshot(dst), new_year)


def prefill(master: Path, dst: Path, db: pd.DataFrame, country: dict, source_name: str) -> int:
    """Copy the master and write historical values for one country from the database.

    Returns the number of cells written. Provenance goes in the Metadata sheet.
    """
    shutil.copyfile(master, dst)
    wb = load_workbook(dst)
    ws = wb[DATA_SHEET]
    cols = year_columns(ws)
    sub = db[(db.country_code == country["code"]) & (db.variable.isin(INPUTS))]
    n = 0
    for r in sub.itertuples(index=False):
        if r.year in cols:
            ws.cell(ROW[r.variable], cols[r.year], r.value)
            n += 1
    write_meta(wb, Country_code=country["code"], Country_name=country["name"],
               Prefilled_from=source_name, Prefill_date=date.today().isoformat(),
               Prefilled_cells=n)
    wb.save(dst)
    return n


def translate(master: Path, dst: Path, labels: pd.DataFrame, lang: str) -> list[str]:
    """Write a language version: labels only, structure untouched.

    Returns codes with no translation (left in English and listed for review).
    """
    wb = load_workbook(master)
    ws = wb[DATA_SHEET]
    lab = labels.set_index("code")[lang].to_dict()
    missing = []
    for var in VARIABLES:
        text = lab.get(var.code)
        if isinstance(text, str) and text.strip():
            ws.cell(ROW[var.code], LABEL_COL, text)
        else:
            missing.append(var.code)
    if isinstance(lab.get("_header_variable"), str):
        ws.cell(HEADER_ROW, LABEL_COL, lab["_header_variable"])
    write_meta(wb, Language=lang.upper())
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)
    return missing
