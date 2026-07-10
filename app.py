"""
CyberSource Merchant Authorization-Response Analysis — Streamlit UI (batch).

Upload one OR many raw CyberSource Transaction Detail Reports (CSV). Each file is
analysed independently (no cross-file mixing), producing its own workbook. Download
them all at once as a ZIP, or preview/download any single file's result.

    python -m streamlit run app.py
"""
import os
import sys
import zipfile
from io import BytesIO

import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analysis_core import load_rows, analyze, build_workbook, _summary_rows

st.set_page_config(page_title="CyberSource Auth Analysis", layout="wide")
st.title("CyberSource Merchant Authorization-Response Analysis")
st.caption("Upload one or many Transaction Detail Report CSVs. Each file is analysed "
           "on its own — group by merchant reference, extract the authorization result "
           "by ics_auth position — and gets its own workbook.")

files = st.file_uploader("Transaction Detail Report(s) — .csv or .zip", type=["csv", "zip"],
                         accept_multiple_files=True, label_visibility="collapsed")

if not files:
    with st.expander("How it works"):
        st.markdown(
            "1. **Upload** one or more raw CyberSource Transaction Detail Report CSVs — "
            "or drop in **`.zip`** archives of CSVs (or a mix). Every CSV inside is unpacked.\n"
            "2. Each CSV is processed **independently** — rows grouped by "
            "`merchant_ref_number`, only `ics_auth` transactions kept, the authorization "
            "code/flag/description read from the `ics_auth` position.\n"
            "3. You get a **batch summary** with a reconciliation check per file, a "
            "**Download all (ZIP)** button, and a per-file preview + individual download."
        )
    st.stop()


def expand_inputs(uploads):
    """Flatten uploads into (display_name, csv_bytes, error) — unpacking any .zip."""
    items = []
    for f in uploads:
        if f.name.lower().endswith(".zip"):
            try:
                zf = zipfile.ZipFile(BytesIO(f.getvalue()))
            except Exception as e:
                items.append((f.name, None, f"not a valid zip: {e}"))
                continue
            csv_members = [i for i in zf.infolist()
                           if not i.is_dir()
                           and i.filename.lower().endswith(".csv")
                           and not i.filename.startswith("__MACOSX/")
                           and not os.path.basename(i.filename).startswith("._")]
            if not csv_members:
                items.append((f.name, None, "zip contains no .csv files"))
            for info in csv_members:
                items.append((f"{f.name} → {os.path.basename(info.filename)}",
                              zf.read(info), None))
        else:
            items.append((f.name, f.getvalue(), None))
    return items


inputs = expand_inputs(files)

# ---- Analyse every CSV independently ----
results = []
for name, data, err in inputs:
    if err:
        results.append({"name": name, "a": None, "xlsx": None, "error": err})
        continue
    try:
        idx, rows = load_rows(data)
        a = analyze(idx, rows)
        bio = BytesIO()
        build_workbook(a, bio, top_n=5)
        results.append({"name": name, "a": a, "xlsx": bio.getvalue(), "error": None})
    except Exception as e:
        results.append({"name": name, "a": None, "xlsx": None, "error": str(e)})

def out_base(display):
    """Output basename for a workbook, stripping any 'zip → member.csv' decoration."""
    part = display.split(" → ")[-1]
    return os.path.splitext(os.path.basename(part))[0]


# ---- Batch summary ----
st.subheader(f"Batch summary — {len(results)} CSV file(s)")
summary_tbl = []
for r in results:
    if r["error"]:
        summary_tbl.append({"file": r["name"], "status": f"ERROR: {r['error']}",
                            "rows": "", "unique_txns": "", "auth_attempts": "",
                            "merchants": "", "alignment_failures": ""})
    else:
        a = r["a"]
        sum_auth = sum(len(d["auth_ref"]) for d in a["summ"].values())
        reconciles = (sum_auth == a["n_auth"]) and (len(a["align_fail"]) == 0)
        summary_tbl.append({
            "file": r["name"],
            "status": "OK" if reconciles else "review",
            "rows": a["n_rows"],
            "unique_txns": a["n_unique_ref"],
            "auth_attempts": a["n_auth"],
            "merchants": len(a["by_merch"]),
            "alignment_failures": len(a["align_fail"]),
        })
st.dataframe(summary_tbl, use_container_width=True)

ok = [r for r in results if r["xlsx"]]
errs = [r for r in results if r["error"]]
if errs:
    st.warning(f"{len(errs)} file(s) could not be processed — see the ERROR rows above. "
               "The rest are ready below.")

# ---- Download all as ZIP ----
if ok:
    zbuf = BytesIO()
    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
        used = {}
        for r in ok:
            base = out_base(r["name"])
            name = f"{base}_analysis.xlsx"
            # avoid collisions if two uploads share a base name
            n = used.get(name, 0)
            if n:
                name = f"{base}_analysis_{n+1}.xlsx"
            used[f"{base}_analysis.xlsx"] = n + 1
            z.writestr(name, r["xlsx"])
    st.download_button(
        f"Download all {len(ok)} workbook(s) as ZIP",
        data=zbuf.getvalue(),
        file_name="cybersource_analyses.zip",
        mime="application/zip",
        type="primary",
    )

# ---- Per-file preview + individual download ----
if ok:
    st.subheader("Per-file result")
    pick = st.selectbox("File", [r["name"] for r in ok])
    r = next(x for x in ok if x["name"] == pick)
    a = r["a"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Raw rows", f"{a['n_rows']:,}")
    c2.metric("Unique transactions", f"{a['n_unique_ref']:,}")
    c3.metric("Authorization attempts", f"{a['n_auth']:,}")
    c4.metric("Merchants", len(a["by_merch"]))

    if a["align_fail"]:
        st.warning(f"{len(a['align_fail'])} row(s) could not align ics_rmsg to ics_rcode.")
        st.dataframe(a["align_fail"], use_container_width=True)
    else:
        st.success("All rows aligned cleanly; auth attempts reconcile.")

    st.markdown("**Raw Data Summary**")
    cols, body, total = _summary_rows(a, top_n=5)

    def fmt(row):
        d = dict(zip(cols, row))
        d["auth_attempt_transaction_percentage"] = f"{d['auth_attempt_transaction_percentage']:.2%}"
        d["non_auth_transaction_percentage"] = f"{d['non_auth_transaction_percentage']:.2%}"
        return d

    st.dataframe([fmt(x) for x in body] + [fmt(total)], use_container_width=True)

    st.markdown("**Merchant Response Table**")
    mid = st.selectbox("Merchant", sorted(a["by_merch"].keys()))
    total_m = a["merch_total"][mid]
    st.markdown(f"Total authorization attempts: **{total_m}**")
    items = sorted(a["by_merch"][mid].items(), key=lambda kv: (-kv[1], kv[0][0], kv[0][1]))
    table = [{
        "response_code": code if code else "(none)",
        "response_description": desc if desc else "(no authorization response returned)",
        "number": n,
        "percentage_of_total": f"{(n / total_m if total_m else 0):.2%}",
    } for (code, desc), n in items]
    st.dataframe(table, use_container_width=True)

    st.download_button(
        "Download this workbook (.xlsx)",
        data=r["xlsx"],
        file_name=f"{out_base(r['name'])}_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
