"""Template specification: the single source of truth for the workbook layout.

Every other module (generation, roll-forward, prefill, translation, validation,
consolidation) reads the layout from here, so a change to the template is made
in one place only.
"""
from __future__ import annotations

from dataclasses import dataclass

DATA_SHEET = "Data"
META_SHEET = "Metadata"
HEADER_ROW = 1
FIRST_DATA_ROW = 2
CODE_COL = 1
LABEL_COL = 2
YEAR_COL_START = 3  # column C


@dataclass(frozen=True)
class Variable:
    code: str
    label: str
    kind: str  # "input" or "formula"
    formula: str | None = None  # placeholders: {c} = column letter, {<CODE>} = row number


VARIABLES: list[Variable] = [
    *[Variable(f"E_G{g}", f"Enrolment, primary grade {g}, both sexes", "input")
      for g in range(1, 7)],
    Variable("E_PRI_M", "Enrolment, primary, male", "input"),
    Variable("E_PRI_F", "Enrolment, primary, female", "input"),
    Variable("E_PRI_T", "Enrolment, primary, both sexes", "formula", "=SUM({c}{E_G1}:{c}{E_G6})"),
    Variable("POP_PRI", "Population of official primary school age", "input"),
    Variable("TCH_PRI", "Teachers, primary", "input"),
    Variable(
        "GER_PRI",
        "Gross enrolment ratio, primary (%)",
        "formula",
        '=IF({c}{POP_PRI}>0,ROUND({c}{E_PRI_T}/{c}{POP_PRI}*100,1),"")',
    ),
    Variable(
        "PTR_PRI",
        "Pupil-teacher ratio, primary",
        "formula",
        '=IF({c}{TCH_PRI}>0,ROUND({c}{E_PRI_T}/{c}{TCH_PRI},1),"")',
    ),
]

BY_CODE: dict[str, Variable] = {v.code: v for v in VARIABLES}
ROW: dict[str, int] = {v.code: FIRST_DATA_ROW + i for i, v in enumerate(VARIABLES)}
INPUTS: list[str] = [v.code for v in VARIABLES if v.kind == "input"]
FORMULAS: list[str] = [v.code for v in VARIABLES if v.kind == "formula"]
GRADES: list[str] = [f"E_G{g}" for g in range(1, 7)]


def formula_for(code: str, col_letter: str) -> str:
    """Return the expected Excel formula for a formula variable in a given column."""
    var = BY_CODE[code]
    assert var.formula is not None, f"{code} is not a formula variable"
    return var.formula.format(c=col_letter, **{k: str(r) for k, r in ROW.items()})


def compute(code: str, v: dict[str, float | None]) -> float | None:
    """Recompute a formula variable in Python from input values for one year.

    Mirrors the Excel formulas so indicators can be verified without Excel.
    """
    if code == "E_PRI_T":
        parts = [v.get(g) for g in GRADES]
        return None if any(p is None for p in parts) else float(sum(parts))
    total = v.get("E_PRI_T")
    if total is None:
        total = compute("E_PRI_T", v)
    if code == "GER_PRI":
        pop = v.get("POP_PRI")
        return None if total is None or not pop else round(total / pop * 100, 1)
    if code == "PTR_PRI":
        tch = v.get("TCH_PRI")
        return None if total is None or not tch else round(total / tch, 1)
    raise KeyError(code)
