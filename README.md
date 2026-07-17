# CyberSource Merchant Authorization-Response Analysis

Turns a raw **CyberSource Transaction Detail Report** (CSV) into a merchant-level
authorization-response analysis workbook.

## Files

| File | Purpose |
|------|---------|
| `analysis_core.py` | Core engine — importable, no I/O globals: `load_rows`, `analyze`, `build_workbook`, `run` |
| `app.py` | Streamlit UI — **batch upload** one or many reports as **`.csv` / `.xlsx`**, or **`.zip`** archives of them; each analysed independently; preview + download individual `.xlsx` or all as a ZIP |
| `requirements.txt` | `streamlit`, `openpyxl` |

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What it produces

An `.xlsx` with two sheets:

- **Merchant Response Tables** — one table per `merchant_id`:
  `response_code | response_description | number | percentage_of_total`,
  sorted by count ↓ then code ↑ then description ↑, percentages against that
  merchant's own total authorization attempts.
- **Raw Data Summary** — first five merchants (by raw volume) plus a `TOTAL`
  row computed from **all** merchants, with per-merchant counts and auth/non-auth
  percentages.

## How it works

1. Rows are grouped by **`merchant_ref_number`** (one customer transaction each) —
   never by `request_id`.
2. Only transactions where **`ics_auth`** appears in `ics_applications` are kept.
3. The authorization **code / flag / description** are read from the *same
   position* `ics_auth` occupies in the comma-separated `ics_applications` list,
   using `ics_rcode` / `ics_rflag` / `ics_rmsg`.
4. `ics_rmsg` values that contain internal commas (e.g. *"Lost card, pick up
   (fraud account)"*) would otherwise be torn across positions. These are
   **auto-healed**: since `ics_rcode` / `ics_rflag` never contain internal commas,
   they give a reliable "which positions have a message" pattern, and `ics_rmsg`
   is rebuilt to match it — so new comma-bearing messages align automatically with
   no hardcoded list to maintain. Anything genuinely ambiguous is flagged, never
   guessed.

## Reconciliation

`auth_attempt_transactions` in the Raw Data Summary equals each merchant's
"Total authorization attempts" in the Merchant Response Tables, and the `TOTAL`
row equals the size of the intermediate authorization-attempt dataset.

## API

```python
from analysis_core import run
summary = run("TransactionDetailReport.csv", "analysis.xlsx")
# summary: {output_path, rows, unique_transactions, auth_attempts, merchants, alignment_failures}
```
