"""Damaged submissions are returned to the country with STR01, never crash the run."""
import csv
import sqlite3
from pathlib import Path

import pytest
from openpyxl import load_workbook

from templatekit.pipeline import Pipeline
from templatekit.report import submission_report
from templatekit.validate import load_rules, validate_file
from templatekit.workbook import build_master, write_meta

ROOT = Path(__file__).resolve().parents[1]
RULES = load_rules(ROOT / "rules" / "validation_rules.yaml")


@pytest.fixture
def template(tmp_path) -> Path:
    """A valid (blank) country template to damage in different ways."""
    path = build_master(tmp_path / "ARC_2025_submission.xlsx", [2024, 2025], 2025, "2025.1")
    wb = load_workbook(path)
    write_meta(wb, Country_code="ARC", Country_name="Arcadia")
    wb.save(path)
    return path


def without_sheet(path: Path, sheet: str) -> Path:
    wb = load_workbook(path)
    del wb[sheet]
    wb.save(path)
    return path


def assert_returned_with_str01(path: Path, expect_in_message: str, expect_in_error: str) -> None:
    res = validate_file(path, RULES)
    assert res.status == "Returned to country"
    assert [f.rule_id for f in res.flags] == ["STR01"]
    assert res.flags[0].severity == "blocking"
    assert expect_in_message in res.flags[0].message
    assert expect_in_error in res.error
    assert len(res.sha256) == 64
    html = submission_report(res, path.with_suffix(".html")).read_text(encoding="utf-8")
    assert "Returned to country" in html and "STR01" in html


def test_corrupt_file(template):
    template.write_bytes(template.read_bytes()[:2000])  # truncated download
    assert_returned_with_str01(template, "could not be opened", "BadZipFile")


def test_non_excel_file_renamed_xlsx(tmp_path):
    path = tmp_path / "BOR_2025_submission.xlsx"
    path.write_text("country,year,value\nBOR,2025,1\n", encoding="utf-8")
    assert_returned_with_str01(path, "could not be opened", "BadZipFile")


def test_password_protected_file(tmp_path):
    # Encrypted .xlsx files are OLE compound files: same signature as below.
    path = tmp_path / "COR_2025_submission.xlsx"
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(4096))
    assert_returned_with_str01(path, "password-protected", "BadZipFile")


def test_missing_data_sheet(template):
    without_sheet(template, "Data")
    assert_returned_with_str01(template, "'Data' sheet is missing", "missing sheet(s): Data")


def test_missing_metadata_sheet(template):
    without_sheet(template, "Metadata")
    assert_returned_with_str01(template, "'Metadata' sheet is missing",
                               "missing sheet(s): Metadata")


def test_returned_even_without_str01_in_rule_set(tmp_path):
    path = tmp_path / "x.xlsx"
    path.write_bytes(b"not a workbook")
    res = validate_file(path, [r for r in RULES if r["id"] != "STR01"])
    assert res.status == "Returned to country"
    assert [f.rule_id for f in res.flags] == ["STR01"]


def test_batch_continues_and_logs_errors(tmp_path):
    pl = Pipeline(ROOT / "config.yaml", tmp_path)
    for step in (pl.generate, pl.rollforward, pl.translate, pl.prefill, pl.simulate):
        step()
    subs = pl.p("submissions", "x").parent
    (subs / "ZZA_2025_submission.xlsx").write_bytes(b"not a workbook")
    good = next(subs.glob("ARC_*.xlsx")).read_bytes()
    (subs / "ZZB_2025_submission.xlsx").write_bytes(good)
    without_sheet(subs / "ZZB_2025_submission.xlsx", "Data")

    results = {r.file: r for r in pl.validate()}
    assert len(results) == 7
    assert results["ARC_2025_submission.xlsx"].status == "Accepted"
    assert results["ZZA_2025_submission.xlsx"].status == "Returned to country"
    assert results["ZZB_2025_submission.xlsx"].status == "Returned to country"
    for name in ("ZZA", "ZZB"):
        assert pl.p("reports", f"{name}_2025_submission_review.html").exists()

    with pl.p("qc_log.csv").open(encoding="utf-8") as f:
        notes = {r["file"]: r["note"] for r in csv.DictReader(f) if r["step"] == "validate"}
    assert "BadZipFile" in notes["ZZA_2025_submission.xlsx"]
    assert "missing sheet(s): Data" in notes["ZZB_2025_submission.xlsx"]

    summary = pl.consolidate()  # the release still builds, without the damaged files
    assert summary["excluded"] == 5
    index = pl.p("reports", "index.html").read_text(encoding="utf-8")
    assert ">ZZA_2025_submission.xlsx</a>" in index  # no metadata: linked by file name


def test_returned_mined_template_is_not_loaded(tmp_path):
    pl = Pipeline(ROOT / "config.yaml", tmp_path)
    for step in (pl.generate, pl.rollforward, pl.translate, pl.prefill, pl.simulate):
        step()
    mined = pl.p("mined", f"FEN_{pl.sy}_mined.xlsx")
    wb = load_workbook(mined)
    wb["Data"].delete_rows(3)  # a variable row removed: fails STR01, metadata intact
    wb.save(mined)

    summary = pl.consolidate()
    assert summary["loaded"] == 2  # ARC and DAL only
    assert summary["excluded"] == 4  # BOR, COR, EST and the mined FEN template
    with sqlite3.connect(summary["db"]) as con:
        countries = {r[0] for r in con.execute("SELECT DISTINCT country_code FROM observations")}
        fen = con.execute("SELECT status, loaded, blocking FROM submissions "
                          "WHERE country_code = 'FEN'").fetchall()
    assert countries == {"ARC", "DAL"}
    assert fen == [("Returned to country", 0, 1)]
