# Code walkthrough (10 minutes)

A guided tour of the repository for a screen-share: what to open, in what order, what to say,
and the questions it tends to raise. All data is synthetic.

Before starting: run `python run_pipeline.py all` so `output/` is fresh, and keep
`output/reports/index.html` open in a browser tab.

## Order to show files in

| Min | Open | Point to make |
| --- | --- | --- |
| 0–1 | `README.md`, then `output/reports/index.html` | The outcome first: five submissions, three returned, one accepted with queries, one mined template. |
| 1–2 | `config.yaml` | A new cycle is a change of years here. Countries, languages and rule files are config, not code. |
| 2–3.5 | `templatekit/spec.py` | The single source of truth: `VARIABLES` (rows, input vs formula), `ROW`, `formula_for()` (Excel formula for any column) and `compute()` (the same formula in Python). |
| 3.5–5 | `templatekit/workbook.py`: `write_year_column()`, `diff_rollforward()` | Roll-forward regenerates formulas from the spec, then a cell-by-cell diff classifies every change as expected or unexpected. Show `output/checks/rollforward_diff.csv`. |
| 5–7 | `rules/validation_rules.yaml`, then `templatekit/validate.py`: `validate_file()` | Rules are data (id, type, severity, message). The engine implements rule *types*. Structure is checked first and stops the run if broken. |
| 7–8 | A review report, e.g. `output/reports/COR_2025_submission_review.html` | What the country sees: the cell, the rule, a plain-language fix. COR typed a number over the `SUM` formula (FRM01 + IND01). |
| 8–9 | `templatekit/consolidate.py` | Returned files are excluded, country data beats mined data, revisions to past years are reported. One SQLite file plus CSV exports. |
| 9–10 | `tests/`, `.github/workflows/ci.yml` | One end-to-end test per scenario, plus damaged-file tests (corrupt, not Excel, password-protected, missing sheets); CI runs ruff, the tests and the full pipeline on every push. |

## What each module does

- **`spec.py`**: the template layout and the Python mirror of each Excel formula.
- **`workbook.py`**: build a master, write a year column, data validation, Metadata sheet,
  `snapshot()` of a workbook and the two diffs (roll-forward, structural).
- **`rollforward.py`**: `roll_forward()` (add a year), `prefill()` (history from the database
  extract, provenance in Metadata), `translate()` (labels only, from `labels/labels.csv`).
- **`validate.py`**: reads a submission (SHA-256 fingerprint, metadata, values), applies the
  YAML rules, returns a `Result` whose `status` follows the most severe flag.
- **`report.py`**: one HTML review report per submission plus the status index.
- **`consolidate.py`**: release database, precedence rules and revisions table.
- **`pipeline.py`**: wires the steps together and writes `qc_log.csv` (one line per step/file).
- **`synthetic.py`**: fictional countries and the deliberate errors (demo only).

## Three design decisions worth discussing

1. **One specification drives everything.** Generation, roll-forward, translation, validation
   and consolidation all read `spec.py`. Changing the template means changing one list, and the
   validator automatically knows the new expected formulas. The alternative (each step reading
   the previous year's workbook as its reference) lets errors carry forward silently.
2. **Regenerate, then prove.** The roll-forward doesn't copy last year's column. It writes the
   new column from the spec, then diffs the whole workbook against the previous version and
   stops on anything other than the new year. Translations get the same structural diff, so a
   language version can differ only in labels.
3. **Rules as data, severity as decision.** Statisticians tune thresholds and messages in YAML.
   Severity maps directly to the outcome (returned / accepted with queries / logged), which keeps
   the decision explainable to a country and the same in every review.

## Likely questions, short answers

**Why regenerate formulas instead of copying them?**
Copying propagates whatever is in last year's file, including a formula someone fixed by hand
or broke. Regenerating from a pattern gives a known-correct formula for any column, and the
same function (`formula_for`) is what the validator compares against (FRM01). So "expected" is
defined once, and the diff proves nothing else moved.

**How is the rule engine extended?**
A new rule of an existing type (another range, another year-on-year check) is a YAML block,
with no code change. A new rule *type* needs a short block in `validate_file()` that loops over
`by_type["<new type>"]` and calls `flag()`, plus a test. With many more types I would move each
type into its own function registered in a dict (`type -> check`) so the engine stays flat.

**How would it scale to 200 countries and 4 languages?**
- Each file is processed independently, so validation parallelises directly
  (`ProcessPoolExecutor`, or one job per country in a scheduler). 200 small workbooks is minutes,
  not hours.
- Languages are a column in `labels.csv`. openpyxl writes formulas with English function names,
  which is how Excel stores them internally, so one formula pattern serves every language.
- At that scale I would replace the CSV `qc_log` with a database table (safe under parallel
  writes), use openpyxl's `read_only` mode for validation, and give the reports and index a
  filter by region and status.
- The template itself grows by editing `VARIABLES`; the checks follow automatically.

**How would the SQLite step move to SQL Server or PostgreSQL?**
`consolidate.py` already writes through pandas `to_sql`. Swap the `sqlite3` connection for a
SQLAlchemy engine (`mssql+pyodbc://…` or `postgresql+psycopg://…`). For production I would also:
create the tables up front with primary keys on (country, year, variable, release); load into a
staging table and `MERGE` / `INSERT … ON CONFLICT` in one transaction instead of deleting the
file; keep each release as rows rather than separate files, so revisions become a SQL query
between two releases.

**Why recompute indicators in Python?**
openpyxl reads formulas, not results, and files saved by scripts carry no cached values. Python
recomputation (`compute()`) checks indicators without opening Excel and catches typed-over
formulas (IND01). Results were checked against a LibreOffice recalculation.

**What happens with a damaged file?**
Its bytes are fingerprinted (SHA-256) before it is opened. If it cannot be opened (corrupt,
not really .xlsx, password-protected) or has lost its `Data` or `Metadata` sheet,
`open_template()` returns the reason instead of raising. The same applies to rows moved or
year columns missing. Each becomes a blocking STR01 flag with a plain-language explanation,
the file gets a "Returned to country" review report, the technical error goes into
`qc_log.csv`, and the run continues for the other countries. The Streamlit checker shows the
same flag. `tests/test_damaged_files.py` covers each case.

**What would you improve next?**
A rule registry (above), version numbering that increments within a cycle (`bump_version`
currently always resets to `<year>.1`), a QC log in the database, and an R reader for the
release so analysts working in R use the same tables.
