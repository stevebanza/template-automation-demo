"""Low-level workbook helpers: build, read, style, validate, snapshot and compare."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from .spec import (
    CODE_COL, DATA_SHEET, FIRST_DATA_ROW, FORMULAS, HEADER_ROW, INPUTS, LABEL_COL,
    META_SHEET, ROW, VARIABLES, YEAR_COL_START, formula_for,
)

FORMULA_FILL = PatternFill("solid", fgColor="E7E6E6")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
NO_FILL = PatternFill(fill_type=None)
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF")


# ---------------------------------------------------------------- reading
def year_columns(ws: Worksheet) -> dict[int, int]:
    """Map year -> column index by scanning the header row from column C."""
    out: dict[int, int] = {}
    col = YEAR_COL_START
    while True:
        v = ws.cell(HEADER_ROW, col).value
        if v is None:
            break
        try:
            out[int(v)] = col
        except (TypeError, ValueError):
            break
        col += 1
    return out


def read_meta(wb: Workbook) -> dict[str, object]:
    """Key/value pairs from the Metadata sheet (keys as displayed, e.g. 'Survey year')."""
    ws = wb[META_SHEET]
    return {r[0]: r[1] for r in ws.iter_rows(min_row=1, values_only=True) if r and r[0]}


def write_meta(wb: Workbook, **kv: object) -> None:
    """Set Metadata entries; underscores in keys become spaces, new keys are appended."""
    ws = wb[META_SHEET]
    rows = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1)}
    for key, value in kv.items():
        key = key.replace("_", " ")
        r = rows.get(key) or ws.max_row + 1
        ws.cell(r, 1, key).font = Font(bold=True)
        ws.cell(r, 2, value)
        rows[key] = r


# ---------------------------------------------------------------- writing
def write_year_column(ws: Worksheet, year: int, col: int, survey_year: int) -> None:
    """Write one year column: header, formulas regenerated from the spec, styling."""
    letter = get_column_letter(col)
    h = ws.cell(HEADER_ROW, col, year)
    h.font, h.fill, h.alignment = HEADER_FONT, HEADER_FILL, Alignment(horizontal="center")
    ws.column_dimensions[letter].width = 13
    for var in VARIABLES:
        c = ws.cell(ROW[var.code], col)
        if var.kind == "formula":
            c.value = formula_for(var.code, letter)
            c.fill = FORMULA_FILL
        else:
            c.fill = INPUT_FILL if year == survey_year else NO_FILL
        c.number_format = "0.0" if var.code in ("GER_PRI", "PTR_PRI") else "#,##0"


def apply_validation(ws: Worksheet, cols: list[int]) -> None:
    """(Re)build data validation: input cells must be numbers >= 0."""
    ws.data_validations.dataValidation = []
    dv = DataValidation(
        type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
        showErrorMessage=True, errorTitle="Invalid value",
        error="Enter a number greater than or equal to 0.",
    )
    first, last = get_column_letter(min(cols)), get_column_letter(max(cols))
    for code in INPUTS:
        dv.add(f"{first}{ROW[code]}:{last}{ROW[code]}")
    ws.add_data_validation(dv)


def build_master(path: Path, years: list[int], survey_year: int, version: str) -> Path:
    """Build a blank master template from the spec for the given years."""
    wb = Workbook()
    ws = wb.active
    ws.title = DATA_SHEET
    for col, text in ((CODE_COL, "Code"), (LABEL_COL, "Variable")):
        c = ws.cell(HEADER_ROW, col, text)
        c.font, c.fill = HEADER_FONT, HEADER_FILL
    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 46
    for var in VARIABLES:
        ws.cell(ROW[var.code], CODE_COL, var.code)
        ws.cell(ROW[var.code], LABEL_COL, var.label)
    cols = []
    for i, y in enumerate(years):
        col = YEAR_COL_START + i
        write_year_column(ws, y, col, survey_year)
        cols.append(col)
    apply_validation(ws, cols)
    ws.freeze_panes = "C2"

    meta = wb.create_sheet(META_SHEET)
    meta.column_dimensions["A"].width = 22
    meta.column_dimensions["B"].width = 40
    write_meta(wb, Template="Primary education multi-year template (demo)", Version=version,
               Survey_year=survey_year, Language="EN", Country_code="", Country_name="",
               Source="")
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# ---------------------------------------------------------------- comparing
def snapshot(path: Path) -> dict:
    """Capture what matters structurally: codes, labels, years, cell contents, validation."""
    wb = load_workbook(path)
    ws = wb[DATA_SHEET]
    years = year_columns(ws)
    n = len(VARIABLES)
    codes = [ws.cell(FIRST_DATA_ROW + i, CODE_COL).value for i in range(n)]
    labels = [ws.cell(FIRST_DATA_ROW + i, LABEL_COL).value for i in range(n)]
    cells = {(code, y): ws.cell(FIRST_DATA_ROW + i, col).value
             for i, code in enumerate(codes) for y, col in years.items()}
    dv = sorted(str(r) for d in ws.data_validations.dataValidation for r in d.sqref.ranges)
    return {"codes": codes, "labels": labels, "years": sorted(years), "cols": years,
            "cells": cells, "dv": dv, "meta": read_meta(wb)}


def expected_dv(cols: list[int]) -> list[str]:
    """Data-validation ranges a template with these year columns should have."""
    first, last = get_column_letter(min(cols)), get_column_letter(max(cols))
    return sorted(f"{first}{ROW[c]}:{last}{ROW[c]}" for c in INPUTS)


def compare_structure(a: dict, b: dict, ignore_labels: bool = True) -> list[dict]:
    """Differences between two templates that should be structurally identical."""
    issues = []
    if a["codes"] != b["codes"]:
        issues.append({"what": "variable codes", "a": a["codes"], "b": b["codes"]})
    if not ignore_labels and a["labels"] != b["labels"]:
        issues.append({"what": "labels", "a": a["labels"], "b": b["labels"]})
    if a["years"] != b["years"]:
        issues.append({"what": "year columns", "a": a["years"], "b": b["years"]})
    for key, va in a["cells"].items():
        vb = b["cells"].get(key)
        if va != vb:
            issues.append({"what": f"cell {key[0]} {key[1]}", "a": va, "b": vb})
    if a["dv"] != b["dv"]:
        issues.append({"what": "data validation", "a": a["dv"], "b": b["dv"]})
    return issues


def diff_rollforward(old: dict, new: dict, new_year: int) -> list[dict]:
    """Every difference between the old and rolled-forward template, classed as
    expected (the new year) or unexpected (anything else)."""
    out = []

    def add(status, what, old_v=None, new_v=None):
        out.append({"status": status, "what": what, "old": old_v, "new": new_v})

    if old["codes"] != new["codes"]:
        add("unexpected", "variable codes changed", old["codes"], new["codes"])
    if old["labels"] != new["labels"]:
        add("unexpected", "labels changed", old["labels"], new["labels"])
    added = sorted(set(new["years"]) - set(old["years"]))
    removed = sorted(set(old["years"]) - set(new["years"]))
    if added != [new_year] or removed:
        add("unexpected", "year columns", old["years"], new["years"])
    for key, v in old["cells"].items():
        if new["cells"].get(key) != v:
            add("unexpected", f"existing cell changed: {key[0]} {key[1]}", v, new["cells"].get(key))
    col = new["cols"].get(new_year)
    if col:
        letter = get_column_letter(col)
        for code in FORMULAS:
            exp, got = formula_for(code, letter), new["cells"].get((code, new_year))
            add("expected" if got == exp else "unexpected", f"new formula {code} {new_year}",
                exp, got)
        for code in INPUTS:
            got = new["cells"].get((code, new_year))
            if got is not None:
                add("unexpected", f"new input cell not blank: {code} {new_year}", None, got)
        want = expected_dv(list(new["cols"].values()))
        add("expected" if new["dv"] == want else "unexpected", "data validation extended",
            want, new["dv"])
    return out
