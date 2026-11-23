"""Streamlit front end: check one template file without writing any code.

    streamlit run app.py
"""
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from templatekit.report import SEV_LABEL, submission_report
from templatekit.validate import load_rules, validate_file

st.set_page_config(page_title="Template checker", layout="wide")
st.title("Template checker")
st.caption("Upload a completed template to run the same validation rules as the pipeline.")

rules = load_rules(Path(__file__).parent / "rules" / "validation_rules.yaml")
prev_file = st.sidebar.file_uploader("Previous release extract (CSV, optional)", type="csv")
up = st.file_uploader("Template (.xlsx)", type="xlsx")

if up:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / up.name
        path.write_bytes(up.getvalue())
        prev = pd.read_csv(prev_file) if prev_file else None
        res = validate_file(path, rules, prev)
        report = submission_report(res, Path(tmp) / f"{path.stem}_review.html").read_bytes()

    st.subheader(res.title)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Status", res.status)
    c2.metric("Must fix", res.count("blocking"))
    c3.metric("Please confirm", res.count("to_confirm"))
    c4.metric("For information", res.count("info"))
    if res.flags:
        st.dataframe(pd.DataFrame([{
            "Level": SEV_LABEL[f.severity], "Rule": f.rule_id, "Cell": f.cell,
            "Year": f.year, "What to check": f.message} for f in res.flags]),
            hide_index=True)
    else:
        st.success("No issues found.")
    st.download_button("Download review report (HTML)", report,
                       file_name=f"{path.stem}_review.html", mime="text/html")
