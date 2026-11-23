"""HTML review reports: one per submission (for the country focal point) and an index."""
from __future__ import annotations

import html
from datetime import datetime
from pathlib import Path

from .validate import SEVERITIES as SEVERITY_ORDER
from .validate import Result

CSS = """
body{font-family:Segoe UI,Roboto,Arial,sans-serif;margin:32px auto;max-width:1000px;color:#1b1b1b;
  padding:0 16px}
h1{font-size:22px;margin-bottom:4px}.sub{color:#555;margin-top:0}
table{border-collapse:collapse;width:100%;margin:16px 0;font-size:14px}
th,td{border-bottom:1px solid #ddd;padding:8px;text-align:left;vertical-align:top}
th{background:#1f4e79;color:#fff;font-weight:600}
.pill{display:inline-block;white-space:nowrap;padding:2px 10px;border-radius:12px;font-size:12px;
  font-weight:600}
.blocking{background:#fde2e1;color:#9b1c1c}.to_confirm{background:#fff4d6;color:#7a5200}
.info{background:#e3effa;color:#1f4e79}.ok{background:#e3f4e8;color:#1e6b34}
.cards{display:flex;gap:12px;margin:16px 0}
.card{flex:1;border:1px solid #ddd;border-radius:8px;padding:12px}
.card b{font-size:24px;display:block}
"""

STATUS_CLASS = {"Returned to country": "blocking", "Accepted with queries": "to_confirm",
                "Accepted": "ok"}
SEV_LABEL = {"blocking": "Must fix", "to_confirm": "Please confirm", "info": "For information"}


def _page(title: str, body: str) -> str:
    return (f"<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'>"
            f"<title>{html.escape(title)}</title><style>{CSS}</style></head><body>{body}"
            f"<p class='sub'>Generated {datetime.now():%Y-%m-%d %H:%M} · synthetic demo data</p>"
            f"</body></html>")


def submission_report(res: Result, path: Path) -> Path:
    """Write the HTML review report sent back to the country focal point."""
    e = html.escape
    st = res.status
    rows = "".join(
        f"<tr><td><span class='pill {f.severity}'>{SEV_LABEL[f.severity]}</span></td>"
        f"<td>{e(f.rule_id)}</td><td>{e(f.cell)}</td><td>{e(str(f.year or ''))}</td>"
        f"<td>{e(f.message)}</td></tr>"
        for f in sorted(res.flags, key=lambda f: SEVERITY_ORDER.index(f.severity))
    ) or "<tr><td colspan='5'>No issues found.</td></tr>"
    body = f"""
<h1>Data review: {e(res.country_name)} ({e(res.country_code)}), survey year {res.survey_year}</h1>
<p class='sub'>File {e(res.file)} · template version {e(res.version)} · language {e(res.language)}
· SHA-256 {res.sha256[:12]}…</p>
<p>Status: <span class='pill {STATUS_CLASS[st]}'>{st}</span></p>
<div class='cards'>
 <div class='card'><b>{res.count('blocking')}</b>must be fixed before the data can be used</div>
 <div class='card'><b>{res.count('to_confirm')}</b>need confirmation or an explanatory note</div>
 <div class='card'><b>{res.count('info')}</b>for information (e.g. revisions to past years)</div>
</div>
<table><tr><th>Level</th><th>Rule</th><th>Cell</th><th>Year</th><th>What to check</th></tr>
{rows}</table>
<p>Please correct the cells listed as <i>Must fix</i> in the same template and send it back.
For <i>Please confirm</i> items, either correct the value or reply with a short explanation.</p>
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_page(f"Review {res.country_code} {res.survey_year}", body), encoding="utf-8")
    return path


def index_report(results: list[tuple[Result, str]], path: Path, title: str) -> Path:
    """Write the status table for all templates; a link of '' means no review report."""
    e = html.escape

    def name(r, link):
        return f"<a href='{e(link)}'>{e(r.country_name)}</a>" if link else e(r.country_name)

    rows = "".join(
        f"<tr><td>{name(r, link)}</td><td>{e(r.country_code)}</td>"
        f"<td>{e(r.source)}</td>"
        f"<td><span class='pill {STATUS_CLASS[r.status]}'>{r.status}</span></td>"
        f"<td>{r.count('blocking')}</td><td>{r.count('to_confirm')}</td>"
        f"<td>{r.count('info')}</td></tr>"
        for r, link in results)
    body = (f"<h1>{e(title)}</h1><table><tr><th>Country</th><th>Code</th><th>Source</th>"
            f"<th>Status</th><th>Must fix</th><th>Confirm</th><th>Info</th></tr>{rows}</table>")
    path.write_text(_page(title, body), encoding="utf-8")
    return path
