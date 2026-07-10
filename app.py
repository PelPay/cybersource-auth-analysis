"""
CyberSource Merchant Authorization-Response Analysis — Streamlit UI.

Upload a raw CyberSource Transaction Detail Report (CSV); the app produces the
Merchant Response Tables + Raw Data Summary workbook and previews the results.

    python -m streamlit run app.py
"""
import os
import sys
import tempfile
from io import BytesIO

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analysis_core import load_rows, analyze, build_workbook

st.set_page_config(page_title="CyberSource Auth Analysis", layout="wide")
st.title("CyberSource Merchant Authorization-Response Analysis")
st.caption("Upload a raw CyberSource Transaction Detail Report (CSV). "
           "Groups by merchant reference, extracts the authorization result by "
           "ics_auth position, and builds the per-merchant response tables.")

uploaded = st.file_uploader("Transaction Detail Report (.csv)", type=["csv"],
                            label_visibility="collapsed")

if not uploaded:
    with st.expander("How it works"):
        st.markdown(
            "1. **Upload** the raw CyberSource Transaction Detail Report CSV "
            "(metadata rows and the header row are detected automatically).\n"
            "2. Rows are grouped by **`merchant_ref_number`** (one customer transaction each).\n"
            "3. Only transactions where **`ics_auth`** was attempted are kept; the "
            "authorization **code / flag / description** are read from the same position "
            "`ics_auth` occupies in `ics_applications`.\n"
            "4. Results are grouped by **`merchant_id`** into response tables, plus a "
            "**Raw Data Summary** that reconciles auth attempts back to the source.\n\n"
            "Download the formatted `.xlsx` when it's ready."
        )
    st.stop()

# ---- Analyse ----
try:
    idx, rows = load_rows(uploaded.getvalue())
    a = analyze(idx, rows)
except Exception as e:
    st.error(f"Could not process this file: {e}")
    st.stop()

# ---- Reconciliation banner ----
sum_auth = sum(len(d["auth_ref"]) for d in a["summ"].values())
c1, c2, c3, c4 = st.columns(4)
c1.metric("Raw rows", f"{a['n_rows']:,}")
c2.metric("Unique transactions", f"{a['n_unique_ref']:,}")
c3.metric("Authorization attempts", f"{a['n_auth']:,}")
c4.metric("Merchants", len(a["by_merch"]))

if a["align_fail"]:
    st.warning(f"{len(a['align_fail'])} row(s) could not align ics_rmsg to ics_rcode "
               "(a response description may contain an unrecognised comma). "
               "These are listed below — the rest of the analysis is unaffected.")
    st.dataframe(a["align_fail"], use_container_width=True)
else:
    st.success(f"All rows aligned cleanly. auth_attempt_transactions reconciles: "
               f"{sum_auth} = {a['n_auth']} intermediate records.")

# ---- Raw Data Summary preview ----
st.subheader("Raw Data Summary")
from analysis_core import _summary_rows
cols, body, total = _summary_rows(a, top_n=5)
def fmt(row):
    d = dict(zip(cols, row))
    d["auth_attempt_transaction_percentage"] = f"{d['auth_attempt_transaction_percentage']:.2%}"
    d["non_auth_transaction_percentage"] = f"{d['non_auth_transaction_percentage']:.2%}"
    return d
st.dataframe([fmt(r) for r in body] + [fmt(total)], use_container_width=True)
st.caption("First five merchants by raw volume, plus a TOTAL computed from all merchants.")

# ---- Per-merchant response table preview ----
st.subheader("Merchant Response Tables")
mid = st.selectbox("Merchant", sorted(a["by_merch"].keys()))
total_m = a["merch_total"][mid]
st.markdown(f"**Total authorization attempts:** {total_m}")
items = sorted(a["by_merch"][mid].items(), key=lambda kv: (-kv[1], kv[0][0], kv[0][1]))
table = [{
    "response_code": code if code else "(none)",
    "response_description": desc if desc else "(no authorization response returned)",
    "number": n,
    "percentage_of_total": f"{(n / total_m if total_m else 0):.2%}",
} for (code, desc), n in items]
st.dataframe(table, use_container_width=True)

# ---- Build + download ----
tmp = tempfile.mkdtemp()
out_path = os.path.join(tmp, "CyberSource_Merchant_Response_Analysis.xlsx")
build_workbook(a, out_path, top_n=5)
with open(out_path, "rb") as fh:
    st.download_button(
        "Download analysis workbook (.xlsx)",
        data=fh.read(),
        file_name="CyberSource_Merchant_Response_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
