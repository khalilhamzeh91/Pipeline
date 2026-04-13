# Sales Pipeline Dashboard — Project README

## Overview

A Streamlit-based sales analytics dashboard deployed on Streamlit Cloud, with companion standalone Excel report generators. The project handles three data domains: **Pipeline** (active opportunities), **Awarded Deals** (won contracts), and **AM Pipeline** (Account Manager / Capability Sales view).

**Live app:** Streamlit Cloud — repo `khalilhamzeh91/Pipeline`, branch `main`, entry point `pipeline_dashboard.py`

---

## Repository Files

| File | Purpose |
|------|---------|
| `pipeline_dashboard.py` | Main Streamlit app — all tabs, filters, charts, and Excel export functions |
| `generate_pipeline_report.py` | Standalone script — generates Pipeline Excel report from local file |
| `generate_awarded_report.py` | Standalone script — generates Awarded Deals Excel report from local file |
| `generate_am_pipeline_report.py` | Standalone script — generates AM Pipeline Excel report from local file |

---

## Input Excel Files

| Upload / File | Data | Key Columns |
|---|---|---|
| **Pipeline Excel** | Active opportunities | Stage, Account Manager, Sector, BU, DU, Total Gross, Total Net, Winning Probability, Forecasted, Strategic Opportunity, Closure Due Quarter, Est. Close Date, Gross (breakdown), Net (breakdown), Overdue |
| **Awarded Deals 2026** | Won contracts 2026 | Opportunity Name, Account Manager, Stage, BU, DU, Total Gross, Total Net, Project Value, Award Quarter, Contracted, Type (New/Renew), Gross (breakdown), Net (breakdown) |
| **Awarded Deals 2025** | Won contracts 2025 | Same structure as 2026 — loaded together, differentiated by "Year" column |
| **AM Pipeline** | Capability Sales view | Same as Pipeline but with **Capability Sales** column (multi-AM per deal, newline-separated) and monthly columns (January–December) |
| **Book3** | Resource forecast | BU, Project Type, Project Name, monthly columns, Grand Total |
| **Charter of Accounts** | DU → BU mapping | `C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx` |

---

## Dashboard Tabs

### 🔵 Pipeline Tab
Upload: **Pipeline Excel (weekly report)**

Charts & tables: Stage funnel, Sector bar, AM performance, Quarterly plan, Win probability pie, Strategic vs Regular, Source of Opportunity, DU Breakdown table, Forecast by DU, Overdue deals table, Full pipeline table.

**Excel export:** 10 sheets
1. **Summary** — KPI metrics (SUMIF/COUNTIF formulas referencing Full Pipeline), Stage table (SUMIF formulas)
2. **DU Breakdown** — Two-pass layout with formula-based BU subtotals and Grand Total
3. **Forecast per DU** — Summary + Detail sections, BU subtotals as SUM formulas
4. **Sector & AM** — Sector table + AM table, both with TOTAL rows
5. **Quarterly & Probability** — Quarter table + Probability table, both with TOTAL rows
6. **Forecast** — Forecasted deals, TOTAL row
7. **Overdue Deals** — Overdue deals, TOTAL row
8. **Full Pipeline** — All opportunities, TOTAL row (fp_last_row used by Summary formulas)
9. **Pipeline Breakdown** — Styled, one row per DU per opportunity; colored column groups (deal/DU/finance/other); SUM formulas for Total Gross/Net per deal
10. **Book3 Mapping** — Only appears if Book3 file uploaded

---

### 🟢 Awarded Deals Tab
Upload: **Awarded Deals 2026** + optionally **Awarded Deals 2025**

**Excel export:** 6 sheets
1. **Summary** — KPI metrics, Year table (SUM formula TOTAL), Stage breakdown
2. **DU Breakdown** — Two-pass layout, formula-based BU subtotals & Grand Total, 6 columns (Gross, Net, Net 2025, Net 2026)
3. **Account Manager** — AM totals with TOTAL row (SUM formulas)
4. **Award Quarter** — Quarterly breakdown + New vs Renew breakdown, both with TOTAL rows
5. **All Awarded Deals** — Full data table
6. **Awarded Breakdown** — Styled, one row per DU per opportunity, SUM formulas for Total Gross/Net

---

### 🟠 AM Pipeline Tab
Upload: **AM Pipeline (Capability Sales)** — expects `data (2) (1).xlsx` format

Charts: Net by Account Manager bar, Monthly pipeline bar. Table: Full AM Pipeline.

**Excel export:** 8 sheets
1. **Summary** — KPI metrics (cross-sheet formulas), Stage table
2. **By Account Manager** — AM totals with TOTAL row + deal detail by AM with subtotals
3. **Monthly Pipeline** — One row per deal with Jan–Dec columns, TOTAL row, Monthly Summary section
4. **DU Breakdown** — Two-pass layout, formula-based subtotals
5. **Quarterly Plan** — Quarter table with TOTAL row
6. **AM Breakdown** — Styled, one row per AM per opportunity, colored column groups
7. **Full Pipeline** — All opportunities with TOTAL row
8. **Pipeline Breakdown** — Styled, one row per DU per opportunity (same format as main Pipeline Breakdown)

---

### 🔗 Book3 Mapping Tab
Upload: **Book3** + optionally Pipeline and/or Awarded

Fuzzy-matches Book3 project names against Pipeline and Awarded deal names. Color-coded: green (≥0.70), yellow (0.55–0.69), red (<0.55).

---

## Key Technical Patterns

### Formula-Based Excel Output
All totals/subtotals use Excel SUM/SUMIF/COUNTIF formulas (not hardcoded Python values):
- Summary KPIs: `=SUMIF('Full Pipeline'!L3:L{n},"Yes",'Full Pipeline'!I3:I{n})`
- Stage table: `=COUNTIF('Full Pipeline'!D3:D{n},D4)`
- DU subtotals: `=C5+C7+C10` (explicit cell references to DU rows)
- TOTAL rows: `=SUM(C3:C{last_row})`

### Two-Pass DU Breakdown
BU subtotal rows can't be written before knowing child DU row positions. Solution:
1. **First pass:** Walk the data, record xlsxwriter row number for every BU/DU/opp. Store `(row_number, deal_row)` tuples.
2. **Second pass:** Write with formula-based subtotals using the pre-recorded positions.

### AM Name Normalization
In `_clean_am_list()` and applied at load time in `load_am_pipeline()`:
- Any name containing "khalil" → `"Khalil Hamzeh"`
- Any name containing "yazan" → `"Yazan Al Razem"`
- Deduplication via `dict.fromkeys()` preserving insertion order

### Column Name Normalization (AM Pipeline)
The AM pipeline file uses `"Gross (Breakdown)"` (capital B). At load time, `load_am_pipeline()` renames to `"Gross (breakdown)"` (lowercase) so `_expand_deals()` works correctly.

### _expand_deals()
Expands a DataFrame with multi-value DU/Gross/Net cells (newline-separated) into one row per DU. Adds `BU_exp`, `DU_exp`, `Gross_exp`, `Net_exp`, `_is_first`, `_du_count`, `_deal_idx` columns. Used by Pipeline Breakdown and Awarded Breakdown styled sheets.

### Caching
All load/export functions decorated with `@st.cache_data`. Cache key is the file bytes, so re-uploading the same file returns cached result. New Streamlit deployment clears cache.

---

## Local Standalone Scripts

### generate_pipeline_report.py
```
INPUT_FILE = r"C:\Users\khali\Downloads\<pipeline_file>.xlsx"
COA_FILE   = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
python generate_pipeline_report.py
```

### generate_awarded_report.py
```
INPUT_FILE_26 = r"C:\Users\khali\Downloads\<awarded_2026>.xlsx"
INPUT_FILE_25 = r"C:\Users\khali\Downloads\<awarded_2025>.xlsx"
COA_FILE      = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
python generate_awarded_report.py
```

### generate_am_pipeline_report.py
```
INPUT_FILE = r"C:\Users\khali\Downloads\data (2) (1).xlsx"
COA_FILE   = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
python generate_am_pipeline_report.py
```

---

## Common Issues & Fixes

| Issue | Cause | Fix Applied |
|---|---|---|
| `NameError: tot_lbl` | Format defined inside DU section, used in Summary | Moved to global FORMATS block |
| `f-string syntax error` nested generator | Python f-strings can't contain generator expressions with quotes | Pre-compute join string: `_joined = "+".join(...); "=" + _joined` |
| `KeyError: 'Project value...'` in awarded | 2025 file has different column names than 2026 | `load_awarded()` uses fuzzy column matching |
| `IndexError` in AM pipeline export | Lambda using `opp_rows.index()` on list of row numbers | Store `(row_num, deal_row)` tuples in first pass; unpack in second pass |
| `IndexError` in Pipeline Breakdown | Lambda apply used wrong indexing for Gross column | Normalize column name at load time; removed lambda entirely |
| AM names not normalized | `_clean_am_list` only called during explosion, not at raw data level | Apply normalization in `load_am_pipeline()` to raw DataFrame |

---

## Deployment

```bash
# Push to trigger Streamlit Cloud redeploy
git add -A
git commit -m "message"
git push
```

Streamlit Cloud watches `main` branch of `khalilhamzeh91/Pipeline`. Redeploy is automatic on push. In-memory cache is cleared on each new deployment.
