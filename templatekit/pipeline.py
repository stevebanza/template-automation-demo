"""Pipeline steps. Each step reads config + earlier outputs and writes to the output folder.

    generate     -> previous-cycle master template + database extract   (demo setup)
    rollforward  -> new-cycle master template + diff check
    translate    -> FR/ES versions + structural check
    prefill      -> one prefilled template per country
    simulate     -> simulated country submissions + a mined template    (demo setup)
    validate     -> review report per submission + QC log
    consolidate  -> release database (SQLite + CSV) + revisions report
"""
from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

from . import synthetic
from .consolidate import consolidate
from .report import index_report, submission_report
from .rollforward import prefill, roll_forward, translate
from .validate import Result, load_rules, validate_file
from .workbook import build_master, compare_structure, snapshot


class Pipeline:
    """The survey cycle as a sequence of steps sharing one config and output folder."""

    def __init__(self, config_path: str | Path = "config.yaml", out: str | Path = "output"):
        self.root = Path(config_path).resolve().parent
        self.cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))
        self.out = Path(out)
        c = self.cfg
        self.prev, self.sy = c["previous_survey_year"], c["survey_year"]
        self.prev_years = list(range(c["first_year"], self.prev + 1))
        self.all_years = list(range(c["first_year"], self.sy + 1))

    # -------------------------------------------------------------- paths
    def p(self, *parts: str) -> Path:
        """Path under the output folder; parent directories are created."""
        path = self.out.joinpath(*parts)
        path.parent.mkdir(parents=True, exist_ok=True)
        return path

    def master(self, year: int, lang: str = "EN") -> Path:
        """Path of the master template for a survey year and language."""
        return self.p("templates", f"master_{year}_{lang}.xlsx")

    @property
    def db_path(self) -> Path:
        """Path of the previous release's database extract."""
        return self.p("db", f"database_extract_{self.prev}.csv")

    def log(self, step: str, res: Result | None = None, note: str = "") -> None:
        """Append a line to qc_log.csv (file fingerprint and counts if `res` given)."""
        path = self.p("qc_log.csv")
        new = not path.exists()
        with path.open("a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new:
                w.writerow(["timestamp", "step", "file", "sha256", "country", "survey_year",
                            "version", "status", "blocking", "to_confirm", "info", "note"])
            row = [datetime.now().isoformat(timespec="seconds"), step]
            if res:
                row += [res.file, res.sha256, res.country_code, res.survey_year, res.version,
                        res.status, res.count("blocking"), res.count("to_confirm"),
                        res.count("info")]
            else:
                row += [""] * 9
            w.writerow(row + [note])

    # -------------------------------------------------------------- steps
    def generate(self) -> None:
        """Demo setup: previous-cycle master template and database extract."""
        build_master(self.master(self.prev), self.prev_years, self.prev, f"{self.prev}.1")
        everyone = self.cfg["countries"] + self.cfg["mined_countries"]
        synthetic.database_extract(everyone, self.prev_years, self.cfg["seed"]).to_csv(
            self.db_path, index=False)
        self.log("generate", note=f"master {self.prev} and database extract written")

    def rollforward(self) -> list[dict]:
        """Roll the master forward one year; stop if the diff has unexpected changes."""
        diff = roll_forward(self.master(self.prev), self.master(self.sy), self.sy)
        pd.DataFrame(diff).to_csv(self.p("checks", "rollforward_diff.csv"), index=False)
        bad = [d for d in diff if d["status"] == "unexpected"]
        self.log("rollforward",
                 note=f"{len(diff) - len(bad)} expected, {len(bad)} unexpected changes")
        if bad:
            raise RuntimeError(f"Roll-forward produced unexpected changes: {bad[:3]}")
        return diff

    def translate(self) -> list[dict]:
        """Write each language version; stop if one differs structurally."""
        labels = pd.read_csv(self.root / self.cfg["labels"])
        base = snapshot(self.master(self.sy))
        rows = []
        for lang in self.cfg["languages"]:
            dst = self.master(self.sy, lang.upper())
            missing = translate(self.master(self.sy), dst, labels, lang)
            issues = compare_structure(base, snapshot(dst))
            rows.append({"language": lang, "untranslated": ";".join(missing),
                         "structural_differences": len(issues)})
            if issues:
                raise RuntimeError(f"{dst.name} differs structurally from the master: {issues[:3]}")
        pd.DataFrame(rows).to_csv(self.p("checks", "translation_check.csv"), index=False)
        self.log("translate", note=f"{len(rows)} language versions, structure identical")
        return rows

    def prefill(self) -> dict[str, int]:
        """One prefilled template per country; returns cells written per country."""
        db = pd.read_csv(self.db_path)
        out = {}
        for c in self.cfg["countries"]:
            dst = self.p("templates", "prefilled", f"{c['code']}_{self.sy}.xlsx")
            out[c["code"]] = prefill(self.master(self.sy), dst, db, c, self.db_path.name)
        self.log("prefill", note=f"{len(out)} country templates prefilled")
        return out

    def simulate(self) -> None:
        """Demo setup: simulated country submissions and mined templates."""
        for c in self.cfg["countries"]:
            synthetic.simulate_submission(
                self.p("templates", "prefilled", f"{c['code']}_{self.sy}.xlsx"),
                self.p("submissions", f"{c['code']}_{self.sy}_submission.xlsx"),
                c, self.sy, self.all_years, self.cfg["seed"], c["scenario"])
        for c in self.cfg["mined_countries"]:
            synthetic.make_mined_template(self.master(self.sy),
                                          self.p("mined", f"{c['code']}_{self.sy}_mined.xlsx"),
                                          c, self.all_years, self.cfg["seed"])
        self.log("simulate", note="synthetic submissions and mined templates written")

    def validate(self) -> list[Result]:
        """Validate every submission, write its review report and log it."""
        rules = load_rules(self.root / self.cfg["rules"])
        prev = pd.read_csv(self.db_path)
        results = []
        for f in sorted(self.p("submissions", "x").parent.glob("*.xlsx")):
            res = validate_file(f, rules, prev)
            submission_report(res, self.p("reports", f"{f.stem}_review.html"))
            self.log("validate", res)
            results.append(res)
        self.results = results
        return results

    def consolidate(self) -> dict:
        """Validate mined templates, write the status index and build the release."""
        results = getattr(self, "results", None) or self.validate()
        rules = load_rules(self.root / self.cfg["rules"], only=self.cfg["mined_rules"])
        mined = []
        for f in sorted(self.p("mined", "x").parent.glob("*.xlsx")):
            res = validate_file(f, rules)
            self.log("validate-mined", res)
            mined.append(res)
        index_report([(r, f"{Path(r.file).stem}_review.html") for r in results] +
                     [(r, "") for r in mined],
                     self.p("reports", "index.html"), f"Survey {self.sy}: submission status")
        summary = consolidate(results, mined, pd.read_csv(self.db_path),
                              self.p("release", "x").parent, str(self.sy))
        self.log("consolidate", note=f"{summary['rows']} rows, {summary['revisions']} revisions, "
                                     f"{summary['excluded']} submissions excluded")
        return summary

    def run_all(self) -> dict:
        """Run every step in order; returns the consolidation summary."""
        self.generate()
        self.rollforward()
        self.translate()
        self.prefill()
        self.simulate()
        self.validate()
        return self.consolidate()
