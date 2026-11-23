"""Rule-based validation of a submitted template.

Rules are declared in rules/validation_rules.yaml (id, type, severity, parameters,
plain-language message). The engine below implements each rule *type*; adding or
tuning a rule of an existing type needs no code change.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from .spec import (
    BY_CODE, CODE_COL, DATA_SHEET, FIRST_DATA_ROW, FORMULAS, INPUTS, META_SHEET, ROW, VARIABLES,
    compute, formula_for,
)
from .workbook import read_meta, year_columns

SEVERITIES = ("blocking", "to_confirm", "info")
# Encrypted (password-protected) .xlsx and legacy .xls files are OLE compound files.
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
# Used when the loaded rule set has no structure rule (mirrors STR01 in the YAML).
DEFAULT_STRUCTURE_RULE = {
    "id": "STR01", "type": "structure", "severity": "blocking",
    "message": "The template structure was changed ({detail}). "
               "Please enter data in the original template.",
}


@dataclass
class Flag:
    rule_id: str
    severity: str
    message: str
    variable: str = ""
    year: int | None = None
    cell: str = ""


@dataclass
class Result:
    file: str
    sha256: str
    country_code: str
    country_name: str
    survey_year: int | None
    version: str
    language: str
    source: str
    flags: list[Flag] = field(default_factory=list)
    values: dict[tuple[str, int], float | None] = field(default_factory=dict)
    error: str = ""  # technical reason the file could not be read (for the QC log)

    def count(self, severity: str) -> int:
        """Number of flags with the given severity."""
        return sum(f.severity == severity for f in self.flags)

    @property
    def title(self) -> str:
        """'Arcadia (ARC), survey year 2025', or the file name if the metadata is unreadable."""
        if not self.country_code:
            return self.file
        who = self.country_code
        if self.country_name:
            who = f"{self.country_name} ({self.country_code})"
        return f"{who}, survey year {self.survey_year}" if self.survey_year else who

    @property
    def status(self) -> str:
        """Decision driven by the most severe flag."""
        if self.count("blocking"):
            return "Returned to country"
        if self.count("to_confirm"):
            return "Accepted with queries"
        return "Accepted"


def load_rules(path: Path, only: list[str] | None = None) -> list[dict]:
    """Read rules from YAML, optionally keeping only the given rule ids."""
    rules = yaml.safe_load(Path(path).read_text(encoding="utf-8"))["rules"]
    return [r for r in rules if only is None or r["id"] in only]


def sha256(path: Path) -> str:
    """Fingerprint of a file, logged so every check traces back to exact bytes."""
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def open_template(path: Path, data: bytes | None = None) -> tuple[Workbook | None, str, str]:
    """Open a template without raising: (workbook, problem for the country, technical error).

    A file that cannot be opened (corrupt, not .xlsx, password-protected) or that lacks the
    Data or Metadata sheet gives (None, problem, error) so it can be returned with STR01.
    """
    try:
        wb = load_workbook(path)
    except Exception as exc:  # any unreadable file is a structural problem, not a crash
        data = Path(path).read_bytes() if data is None else data
        if data[:8] == OLE_MAGIC:
            problem = "the file is password-protected or in the old .xls format; save it as .xlsx"
        else:
            problem = "the file could not be opened as an Excel workbook (damaged or not .xlsx)"
        return None, problem, f"{type(exc).__name__}: {exc}"
    missing = [s for s in (DATA_SHEET, META_SHEET) if s not in wb.sheetnames]
    if missing:
        names = " and ".join(f"'{s}'" for s in missing)
        verb = "sheets are" if len(missing) > 1 else "sheet is"
        return None, f"the {names} {verb} missing", f"missing sheet(s): {', '.join(missing)}"
    return wb, "", ""


def _int_or_none(x: object) -> int | None:
    try:
        return int(x) if x not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _fmt(x) -> str:
    if x is None:
        return "blank"
    if isinstance(x, float) and not x.is_integer():
        return f"{x:,.1f}"
    if isinstance(x, (int, float)):
        return f"{x:,.0f}"
    return str(x)


def _num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def validate_file(path: Path, rules: list[dict], previous: pd.DataFrame | None = None) -> Result:
    """Apply the rules to one template and return its flags, status and values.

    Never raises for a bad submission: a file that cannot be opened, a missing sheet or an
    altered layout becomes a blocking STR01 flag and stops further checks. `previous`
    (the last release) enables the revision checks.
    """
    path = Path(path)
    data = path.read_bytes()  # fingerprint the exact bytes before trying to open them
    wb, problem, error = open_template(path, data)
    meta = read_meta(wb) if wb else {}
    cols = year_columns(wb[DATA_SHEET]) if wb else {}
    res = Result(
        file=path.name, sha256=hashlib.sha256(data).hexdigest(),
        country_code=str(meta.get("Country code") or ""),
        country_name=str(meta.get("Country name") or ""),
        survey_year=_int_or_none(meta.get("Survey year")),
        version=str(meta.get("Version") or ""), language=str(meta.get("Language") or ""),
        source=str(meta.get("Source") or ""), error=error,
    )
    by_type: dict[str, list[dict]] = {}
    for r in rules:
        by_type.setdefault(r["type"], []).append(r)
    # An unreadable or damaged template must never pass, even with a reduced rule set.
    structure_rules = by_type.get("structure") or [DEFAULT_STRUCTURE_RULE]

    def flag(rule, var="", year=None, **ctx):
        col = cols.get(year) if year else None
        cell = f"{get_column_letter(col)}{ROW[var]}" if col and var in ROW else ""
        label = BY_CODE[var].label if var in BY_CODE else var
        msg = rule["message"].format(label=label, year=year, cell=cell, **{
            k: (_fmt(v) if k in ("observed", "expected") else v) for k, v in ctx.items()})
        res.flags.append(Flag(rule["id"], rule["severity"], msg, var, year, cell))

    # ---- file could not be opened, or a sheet is missing: nothing else can be checked
    if wb is None:
        for rule in structure_rules:
            flag(rule, detail=problem)
        return res

    # ---- structure (stops processing if the template was altered)
    ws = wb[DATA_SHEET]
    codes = [ws.cell(FIRST_DATA_ROW + i, CODE_COL).value for i in range(len(VARIABLES))]
    problems = []
    if codes != [v.code for v in VARIABLES]:
        problems.append("variable rows were added, removed or reordered")
    ys = sorted(cols)
    if not ys or ys != list(range(ys[0], ys[-1] + 1)):
        problems.append("year columns are missing or not consecutive")
    if res.survey_year not in cols:
        problems.append("the survey year column is missing")
    if not res.country_code:
        problems.append("the country code is missing in the Metadata sheet")
    for rule in structure_rules:
        for p in problems:
            flag(rule, detail=p)
    if problems:
        return res

    sy = res.survey_year
    raw = {(c, y): ws.cell(ROW[c], col).value for c in ROW for y, col in cols.items()}

    # ---- input values (non-numeric text treated as missing, and flagged)
    vals: dict[tuple[str, int], float | None] = {}
    for c in INPUTS:
        for y in cols:
            v = raw[(c, y)]
            if v is None or _num(v):
                vals[(c, y)] = v
            else:
                vals[(c, y)] = None
                for rule in by_type.get("required", []):
                    flag(rule, c, y, observed=f"text '{v}'")

    # ---- formulas: recompute in Python; detect overwritten or edited formulas
    typed: dict[tuple[str, int], float] = {}
    for y, col in cols.items():
        yv = {c: vals[(c, y)] for c in INPUTS}
        for c in FORMULAS:
            vals[(c, y)] = compute(c, yv)
            if c == "E_PRI_T":
                yv["E_PRI_T"] = vals[(c, y)]
        letter = get_column_letter(col)
        for c in FORMULAS:
            got, exp = raw[(c, y)], formula_for(c, letter)
            if got != exp:
                for rule in by_type.get("formula_integrity", []):
                    flag(rule, c, y, expected=exp, observed=got)
                if _num(got):
                    typed[(c, y)] = float(got)

    # ---- indicator recalculation vs. typed values
    for rule in by_type.get("indicator_recalc", []):
        for (c, y), got in typed.items():
            exp = vals[(c, y)]
            if exp is not None and abs(got - exp) > abs(exp) * rule["tolerance_pct"] / 100:
                flag(rule, c, y, observed=got, expected=exp)

    def effective(c, y):
        return typed.get((c, y), vals.get((c, y)))

    for rule in by_type.get("required", []):
        for c in INPUTS:
            if raw[(c, sy)] is None:
                flag(rule, c, sy, observed=None)

    for rule in by_type.get("non_negative", []):
        for c in INPUTS:
            for y in cols:
                v = vals[(c, y)]
                if v is not None and v < 0:
                    flag(rule, c, y, observed=v)

    for rule in by_type.get("sum_equals", []):
        target = effective(rule["target"], sy)
        parts = [vals[(p, sy)] for p in rule["parts"]]
        if target is not None and None not in parts:
            s = sum(parts)
            if abs(s - target) > rule.get("tolerance", 0):
                flag(rule, rule["target"], sy, observed=s, expected=target)

    for rule in by_type.get("range", []):
        v = vals.get((rule["variable"], sy))
        if v is not None and not (rule["min"] <= v <= rule["max"]):
            flag(rule, rule["variable"], sy, observed=v, min=rule["min"], max=rule["max"])

    for rule in by_type.get("yoy_change", []):
        for c in rule["variables"]:
            a, b = vals.get((c, sy - 1)), effective(c, sy)
            if a and b is not None:
                pct = (b - a) / a * 100
                if abs(pct) > rule["max_pct"]:
                    flag(rule, c, sy, pct=f"{pct:+.1f}", prev_year=sy - 1, max_pct=rule["max_pct"])

    for rule in by_type.get("cohort", []):
        g = rule["grades"]
        for i in range(len(g) - 1):
            a, b = vals.get((g[i], sy - 1)), vals.get((g[i + 1], sy))
            if a and b is not None and b / a > rule["max_ratio"]:
                flag(rule, g[i + 1], sy, observed=b, expected=a, ratio=f"{b / a:.2f}",
                     grade_from=i + 1, grade_to=i + 2, prev_year=sy - 1,
                     max_pct=round((rule["max_ratio"] - 1) * 100))

    if previous is not None:
        prev = previous[previous.country_code == res.country_code]
        prev = {(r.variable, r.year): r.value for r in prev.itertuples(index=False)}
        for rule in by_type.get("revision", []):
            for (c, y), old in prev.items():
                new = vals.get((c, y))
                if y != sy and new is not None and abs(new - old) > 1e-9:
                    flag(rule, c, y, observed=new, expected=old)

    res.values = vals
    return res
