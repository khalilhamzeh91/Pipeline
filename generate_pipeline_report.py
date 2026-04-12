"""
Weekly Pipeline Report Generator
Produces a formatted Excel workbook with multiple analysis sheets.
Usage: python generate_pipeline_report.py
Output: C:/Users/khali/Downloads/Pipeline_Report_YYYY-MM-DD.xlsx
"""

import pandas as pd
import re
import warnings
from datetime import date
warnings.filterwarnings("ignore")

INPUT_FILE   = r"C:\Users\khali\Downloads\data (2).xlsx"
COA_FILE     = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
BOOK3_FILE   = r"C:\Users\khali\Downloads\Book3.xlsx"
AWARDED_FILE = r"C:\Users\khali\Downloads\data (1).xlsx"   # 2026 awarded
TODAY        = date.today()
OUT_FILE     = rf"C:\Users\khali\Downloads\Pipeline_Report_{TODAY}_v2.xlsx"

# ── LOAD & CLEAN ─────────────────────────────────────────────────────────────
df = pd.read_excel(INPUT_FILE)
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Stage"])
df["Total Gross"]    = pd.to_numeric(df["Total Gross"], errors="coerce").fillna(0)
df["Total Net"]      = pd.to_numeric(df["Total Net"],   errors="coerce").fillna(0)
df["Est. Close Date"]= pd.to_datetime(df["Est. Close Date"], errors="coerce")
df["Overdue"]        = (df["Est. Close Date"] < pd.Timestamp(TODAY))

STAGE_SHORT = {
    "Stage 1: Assessment & Qualification":                    "S1 - Assessment",
    "Stage 2: Discovery & Scoping":                           "S2 - Discovery",
    "Stage 3.1: RFP & BID Qualification":                     "S3.1 - RFP",
    "Stage 3.2: Solution Development & Proposal Submission":  "S3.2 - Solution Dev",
    "Stage 4: Technical Evaluation By Customer":              "S4 - Tech Eval",
    "Stage 5: Resolution/Financial Negotiation":              "S5 - Negotiation",
}
df["Stage_Short"] = df["Stage"].map(STAGE_SHORT).fillna(df["Stage"])

# ── CHARTER OF ACCOUNTS — DU → BU MAPPING ───────────────────────────────────
coa = pd.read_excel(COA_FILE)
coa.columns = coa.columns.str.strip()
coa["_code"] = coa["DU"].str.extract(r"(\d{6})")
DU_TO_BU = coa.dropna(subset=["_code"]).set_index("_code")["BU"].to_dict()

def du_to_bu(du_str):
    m = re.match(r"(\d{6})", str(du_str).strip())
    if m:
        return DU_TO_BU.get(m.group(1), "Unknown")
    return "Unknown"

# ── EXPANSION HELPERS ─────────────────────────────────────────────────────────
def _split_field(value):
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return []
    return [x.strip() for x in str(value).replace(", \n", "\n").replace(",\n", "\n").split("\n") if x.strip()]

def _parse_num(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None

def _expand_deals(df_in, opp_col=None):  # noqa: ARG001
    rows = []
    for deal_idx, row in df_in.iterrows():
        du_parts = _split_field(row.get("DU", ""))
        gr_parts = _split_field(row.get("Gross (breakdown)", ""))
        nt_parts = _split_field(row.get("Net (breakdown)", ""))
        if not du_parts:
            du_parts = [str(row.get("DU", ""))]
        n = len(du_parts)
        def av(parts, n):
            if len(parts) == n:
                return parts
            if len(parts) <= 1:
                return (parts or [None]) + [None] * (n - 1)
            return [parts[0]] + [None] * (n - 1)
        gr = av(gr_parts, n)
        nt = av(nt_parts, n)
        for i, du in enumerate(du_parts):
            nr = {c: row.get(c) for c in df_in.columns}
            nr["BU_exp"]    = du_to_bu(du)
            nr["DU_exp"]    = du
            nr["Gross_exp"] = _parse_num(gr[i])
            nr["Net_exp"]   = _parse_num(nt[i])
            nr["_is_first"] = (i == 0)
            nr["_du_count"] = n
            nr["_deal_idx"] = deal_idx
            rows.append(nr)
    return pd.DataFrame(rows)

# ── DU EXPLOSION ─────────────────────────────────────────────────────────────
du_rows = []
for _, row in df.iterrows():
    dus   = str(row["DU"]).split("\n")   if pd.notna(row["DU"])               else ["Unknown"]
    gross = str(row["Gross (breakdown)"]).replace(",","").split("\n") if pd.notna(row["Gross (breakdown)"]) else ["0"]
    net   = str(row["Net (breakdown)"]).replace(",","").split("\n")   if pd.notna(row["Net (breakdown)"])   else ["0"]
    n = max(len(dus), len(gross), len(net))
    for i in range(n):
        du = dus[i].strip()   if i < len(dus)   else dus[-1].strip()
        g  = gross[i].strip() if i < len(gross) else "0"
        nt = net[i].strip()   if i < len(net)   else "0"
        try: g_val = float(g)
        except: g_val = 0.0
        try: n_val = float(nt)
        except: n_val = 0.0
        du_rows.append({
            "BU":                  du_to_bu(du),
            "DU":                  du,
            "Gross":               g_val,
            "Net":                 n_val,
            "Forecasted":          str(row.get("Forecasted","")).strip(),
            "Account Manager":     row.get("Account Manager",""),
            "Stage":               row.get("Stage",""),
            "Sector":              row.get("Sector",""),
            "Closure Due Quarter": row.get("Closure Due Quarter",""),
            "Account Name":        row.get("Account Name",""),
            "Lead/Opp Name":       row.get("Lead/Opp Name",""),
            "Winning Probability": row.get("Winning Probability",""),
            "Est. Close Date":     row.get("Est. Close Date", pd.NaT),
        })
du_exp = pd.DataFrame(du_rows)

# ── LOAD AWARDED DEALS ────────────────────────────────────────────────────────
import os
aw_df = pd.DataFrame()
if os.path.exists(AWARDED_FILE):
    aw_df = pd.read_excel(AWARDED_FILE, sheet_name="Export")
    aw_df.columns = aw_df.columns.str.strip()
    aw_df["Total Gross"] = pd.to_numeric(aw_df["Total Gross"], errors="coerce").fillna(0)
    aw_df["Total Net"]   = pd.to_numeric(aw_df["Total Net"],   errors="coerce").fillna(0)

# ── LOAD & PARSE BOOK3 (Resource/Revenue Forecast) ────────────────────────────
MONTHS = ["Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
SKIP_KEYWORDS = ["Total","Grand Total","Existing Renewal Total",
                 "Opportunity (ORF) Total","Opportunity Pipeline Total"]

def load_book3(path):
    raw = pd.read_excel(path, header=None)
    # Row 1 = header, data starts row 2
    cols = ["_","BU","Project Type","Project Name"] + MONTHS + ["Grand Total"]
    raw.columns = cols[:len(raw.columns)]
    raw = raw.iloc[2:].reset_index(drop=True)   # skip blank row 0 and header row 1

    current_bu   = None
    current_type = None
    rows = []
    for _, r in raw.iterrows():
        bu   = str(r["BU"]).strip()   if pd.notna(r["BU"])           else ""
        ptype= str(r["Project Type"]).strip() if pd.notna(r["Project Type"]) else ""
        name = str(r["Project Name"]).strip() if pd.notna(r["Project Name"]) else ""

        # Update running BU / Project Type (merged cells come as NaN after first)
        if bu and bu not in ["nan"] and not any(k in bu for k in SKIP_KEYWORDS):
            current_bu = bu
        if ptype and ptype not in ["nan"] and not any(k in ptype for k in SKIP_KEYWORDS):
            current_type = ptype

        # Skip subtotal / blank rows
        if not name or name == "nan" or any(k in name for k in SKIP_KEYWORDS):
            continue
        if not name or any(k in (ptype or "") for k in SKIP_KEYWORDS):
            continue

        row_data = {"BU": current_bu, "Project Type": current_type, "Project Name": name}
        for m in MONTHS:
            val = r.get(m, None)
            row_data[m] = float(val) if pd.notna(val) else 0.0
        row_data["Grand Total"] = float(r.get("Grand Total", 0)) if pd.notna(r.get("Grand Total")) else 0.0
        rows.append(row_data)

    return pd.DataFrame(rows)

book3 = load_book3(BOOK3_FILE)

# ── MATCH BOOK3 ↔ PIPELINE & AWARDED ─────────────────────────────────────────
from difflib import SequenceMatcher

def _clean(s):
    return re.sub(r"[^a-z0-9]", " ", str(s).lower()).split()

def best_match(name, candidates, threshold=0.55):
    """Return (best_candidate, score) or (None, 0) if below threshold."""
    tokens = set(_clean(name))
    best, best_score = None, 0.0
    for cand in candidates:
        cand_tokens = set(_clean(cand))
        # Jaccard token overlap
        overlap = len(tokens & cand_tokens) / max(len(tokens | cand_tokens), 1)
        # Sequence ratio
        seq = SequenceMatcher(None, name.lower(), cand.lower()).ratio()
        score = max(overlap, seq)
        if score > best_score:
            best, best_score = cand, score
    return (best, best_score) if best_score >= threshold else (None, 0.0)

pipe_names  = df["Lead/Opp Name"].dropna().tolist()
award_names = aw_df["Opportunity Name"].dropna().tolist() if not aw_df.empty else []
all_opp_names = pipe_names + award_names

mapping_rows = []
for _, b3 in book3.iterrows():
    pipe_match,  pipe_score  = best_match(b3["Project Name"], pipe_names)
    award_match, award_score = best_match(b3["Project Name"], award_names)

    # Find pipeline row details
    pipe_row  = df[df["Lead/Opp Name"] == pipe_match].iloc[0]  if pipe_match  else None
    award_row = aw_df[aw_df["Opportunity Name"] == award_match].iloc[0] if award_match else None

    row = {
        "Book3 BU":           b3["BU"],
        "Book3 Project Type": b3["Project Type"],
        "Book3 Project Name": b3["Project Name"],
        "Book3 Grand Total":  b3["Grand Total"],
    }
    for m in MONTHS:
        row[f"Book3 {m}"] = b3[m]

    # Pipeline match
    row["Pipeline Match"]       = pipe_match or ""
    row["Pipeline Score"]       = round(pipe_score, 2)
    row["Pipeline Gross (QAR)"] = float(pipe_row["Total Gross"]) if pipe_row is not None else 0.0
    row["Pipeline Net (QAR)"]   = float(pipe_row["Total Net"])   if pipe_row is not None else 0.0
    row["Pipeline Stage"]       = str(pipe_row["Stage"])         if pipe_row is not None else ""
    row["Pipeline AM"]          = str(pipe_row["Account Manager"]) if pipe_row is not None else ""

    # Awarded match
    row["Awarded Match"]        = award_match or ""
    row["Awarded Score"]        = round(award_score, 2)
    row["Awarded Gross (QAR)"]  = float(award_row["Total Gross"]) if award_row is not None else 0.0
    row["Awarded Net (QAR)"]    = float(award_row["Total Net"])   if award_row is not None else 0.0
    row["Awarded Stage"]        = str(award_row["Stage"])         if award_row is not None else ""
    row["Awarded AM"]           = str(award_row["Account Manager"]) if award_row is not None else ""

    mapping_rows.append(row)

mapping_df = pd.DataFrame(mapping_rows)

# ── BUILD SUMMARY TABLES ──────────────────────────────────────────────────────

# 1. KPIs
total_gross      = df["Total Gross"].sum()
total_net        = df["Total Net"].sum()
forecasted_net   = df[df["Forecasted"]=="Yes"]["Total Net"].sum()
forecasted_gross = df[df["Forecasted"]=="Yes"]["Total Gross"].sum()
strategic_count  = len(df[df["Strategic Opportunity"]=="Yes"])
overdue_count    = int(df["Overdue"].sum())

kpi_df = pd.DataFrame([
    {"Metric": "Total Opportunities",          "Value": len(df)},
    {"Metric": "Total Gross Pipeline (QAR)",   "Value": total_gross},
    {"Metric": "Total Net Pipeline (QAR)",     "Value": total_net},
    {"Metric": "Forecasted Gross (QAR)",       "Value": forecasted_gross},
    {"Metric": "Forecasted Net (QAR)",         "Value": forecasted_net},
    {"Metric": "Strategic Opportunities",      "Value": strategic_count},
    {"Metric": "Overdue Deals",                "Value": overdue_count},
])

# 2. By Stage
stage_df = (
    df.groupby("Stage_Short")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index()
    .rename(columns={"Stage_Short":"Stage"})
)
stage_order = list(STAGE_SHORT.values())
stage_df["_ord"] = stage_df["Stage"].map({s:i for i,s in enumerate(stage_order)})
stage_df = stage_df.sort_values("_ord").drop(columns="_ord")

# 3. By Sector
sector_df = (
    df.groupby("Sector")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index()
    .sort_values("Net", ascending=False)
)

# 4. By Account Manager
am_df = (
    df.groupby("Account Manager")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index()
    .sort_values("Net", ascending=False)
)
# Add forecasted per AM
fore_am = df[df["Forecasted"]=="Yes"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
am_df = am_df.merge(fore_am, on="Account Manager", how="left").fillna({"Forecasted Net":0})

# 5. By Quarter
q_df = (
    df.groupby("Closure Due Quarter")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index()
    .sort_values("Closure Due Quarter")
)
fore_q = df[df["Forecasted"]=="Yes"].groupby("Closure Due Quarter")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
q_df = q_df.merge(fore_q, on="Closure Due Quarter", how="left").fillna({"Forecasted Net":0})

# 6. By Winning Probability
prob_df = (
    df.groupby("Winning Probability")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index()
    .sort_values("Net", ascending=False)
)

# 7. DU Breakdown (with BU mapping)
du_totals = (
    du_exp.groupby(["BU","DU"])[["Gross","Net"]]
    .sum().reset_index().sort_values(["BU","Net"], ascending=[True,False])
)
fore_du = du_exp[du_exp["Forecasted"]=="Yes"].groupby("DU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"})
du_totals = du_totals.merge(fore_du, on="DU", how="left").fillna({"Forecasted Net":0})

# 7b. Forecast per DU (detailed — forecasted deals exploded by DU)
fore_du_detail = du_exp[du_exp["Forecasted"]=="Yes"].copy()
fore_du_detail = fore_du_detail.sort_values(["BU","DU","Net"], ascending=[True,True,False])

# 7c. Forecast per DU summary (BU > DU > Quarter)
fore_du_summary = (
    fore_du_detail.groupby(["BU","DU","Closure Due Quarter"])
    .agg(Count=("Lead/Opp Name","count"), Gross=("Gross","sum"), Net=("Net","sum"))
    .reset_index()
    .sort_values(["BU","DU","Closure Due Quarter"])
)

# 8. Forecasted Deals
fore_df = df[df["Forecasted"]=="Yes"][
    ["Account Name","Lead/Opp Name","Stage_Short","Account Manager","Sector",
     "Total Gross","Total Net","Winning Probability","Closure Due Quarter","Est. Close Date"]
].sort_values("Total Net", ascending=False).rename(columns={"Stage_Short":"Stage"})

# 9. Overdue Deals
overdue_df = df[df["Overdue"]][
    ["Account Name","Lead/Opp Name","Stage_Short","Account Manager",
     "Total Net","Est. Close Date","Winning Probability","Closure Due Quarter"]
].sort_values("Est. Close Date").rename(columns={"Stage_Short":"Stage"})

# 10. Full Pipeline
full_df = df[[
    "SNo.","Account Name","Lead/Opp Name","Stage_Short","Account Manager","Sector",
    "BU","DU","Total Gross","Total Net","Winning Probability","Forecasted",
    "Strategic Opportunity","Closure Due Quarter","Est. Close Date","Source of Opportunity","Overdue"
]].sort_values("Total Net", ascending=False).rename(columns={"Stage_Short":"Stage"})

# ── PRE-COMPUTE ROW COUNTS FOR FORMULAS ─────────────────────────────────────
# Full Pipeline: header at Excel row 2, data rows 3..fp_last_row
fp_last_row = 2 + len(full_df)

# ── WRITE EXCEL ───────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as writer:
    wb = writer.book

    # ── FORMATS ──────────────────────────────────────────────────────────────
    fmt_title   = wb.add_format({"bold":True, "font_size":14, "font_color":"#FFFFFF",
                                  "bg_color":"#1a3a6b", "align":"center", "valign":"vcenter"})
    fmt_header  = wb.add_format({"bold":True, "font_color":"#FFFFFF", "bg_color":"#1a3a6b",
                                  "border":1, "align":"center", "valign":"vcenter", "text_wrap":True})
    fmt_kpi_lbl = wb.add_format({"bold":True, "bg_color":"#D9E1F2", "border":1, "align":"left"})
    fmt_kpi_val = wb.add_format({"bold":True, "bg_color":"#EBF0FB", "border":1,
                                  "num_format":"#,##0", "align":"right"})
    fmt_kpi_pct = wb.add_format({"bold":True, "bg_color":"#EBF0FB", "border":1,
                                  "num_format":"0.0%", "align":"right"})
    fmt_num     = wb.add_format({"num_format":"#,##0",   "border":1, "align":"right"})
    fmt_pct     = wb.add_format({"num_format":"0.0",     "border":1, "align":"right"})
    fmt_text    = wb.add_format({"border":1, "align":"left"})
    fmt_date    = wb.add_format({"num_format":"dd-mmm-yyyy", "border":1, "align":"center"})
    fmt_alt     = wb.add_format({"bg_color":"#F2F5FB", "border":1, "align":"left"})
    fmt_alt_num = wb.add_format({"bg_color":"#F2F5FB", "num_format":"#,##0", "border":1, "align":"right"})
    fmt_alt_pct = wb.add_format({"bg_color":"#F2F5FB", "num_format":"0.0", "border":1, "align":"right"})
    fmt_red     = wb.add_format({"bg_color":"#FFE0E0", "border":1, "align":"left"})
    fmt_red_num = wb.add_format({"bg_color":"#FFE0E0", "num_format":"#,##0", "border":1, "align":"right"})
    fmt_red_dt  = wb.add_format({"bg_color":"#FFE0E0", "num_format":"dd-mmm-yyyy", "border":1, "align":"center"})
    fmt_grn     = wb.add_format({"bg_color":"#E2EFDA", "border":1, "align":"left"})
    fmt_grn_num = wb.add_format({"bg_color":"#E2EFDA", "num_format":"#,##0", "border":1, "align":"right"})
    fmt_grn_dt  = wb.add_format({"bg_color":"#E2EFDA", "num_format":"dd-mmm-yyyy", "border":1, "align":"center"})
    tot_fmt = wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF",
                              "num_format":"#,##0","border":1,"align":"right"})
    tot_lbl = wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF",
                              "border":1,"align":"left"})

    def write_header_row(ws, row, cols, widths=None):
        for c, col in enumerate(cols):
            ws.write(row, c, col, fmt_header)
        if widths:
            for c, w in enumerate(widths):
                ws.set_column(c, c, w)

    def write_data_rows(ws, start_row, data, col_types):
        """col_types: list of 'text'|'num'|'pct'|'date' matching data columns"""
        for r, row_data in enumerate(data):
            alt = (r % 2 == 1)
            for c, (val, typ) in enumerate(zip(row_data, col_types)):
                if typ == "num":
                    fmt = fmt_alt_num if alt else fmt_num
                    ws.write_number(start_row+r, c, val if pd.notna(val) else 0, fmt)
                elif typ == "pct":
                    fmt = fmt_alt_pct if alt else fmt_pct
                    ws.write_number(start_row+r, c, val if pd.notna(val) else 0, fmt)
                elif typ == "date":
                    if pd.notna(val):
                        ws.write_datetime(start_row+r, c, val.to_pydatetime(), fmt_date)
                    else:
                        ws.write_blank(start_row+r, c, None, fmt_date)
                else:
                    fmt = fmt_alt if alt else fmt_text
                    ws.write(start_row+r, c, str(val) if pd.notna(val) else "", fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 1 — SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    ws = writer.sheets.get("Summary") or wb.add_worksheet("Summary")
    ws.set_zoom(90)
    ws.set_tab_color("#1a3a6b")

    # Title
    ws.merge_range("A1:C1", f"Weekly Pipeline Summary — {TODAY.strftime('%d %B %Y')}", fmt_title)
    ws.set_row(0, 28)

    # KPIs
    ws.write(2, 0, "Metric", fmt_header)
    ws.write(2, 1, "Value", fmt_header)
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 22)

    # KPI formulas reference the Full Pipeline sheet (data rows 3..fp_last_row)
    _fp = "'Full Pipeline'"
    _n  = fp_last_row
    kpi_formulas = [
        ("Total Opportunities",        f"=COUNTA({_fp}!C3:C{_n})",                                        fmt_kpi_val),
        ("Total Gross Pipeline (QAR)", f"=SUM({_fp}!I3:I{_n})",                                           fmt_kpi_val),
        ("Total Net Pipeline (QAR)",   f"=SUM({_fp}!J3:J{_n})",                                           fmt_kpi_val),
        ("Forecasted Gross (QAR)",     f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!I3:I{_n})',                   fmt_kpi_val),
        ("Forecasted Net (QAR)",       f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!J3:J{_n})',                   fmt_kpi_val),
        ("Strategic Opportunities",    f'=COUNTIF({_fp}!M3:M{_n},"Yes")',                                 fmt_kpi_val),
        ("Overdue Deals",              f'=COUNTIF({_fp}!Q3:Q{_n},"YES")',                                 fmt_kpi_val),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_formulas):
        ws.write(3+i, 0, label, fmt_kpi_lbl)
        ws.write_formula(3+i, 1, formula, fmt)

    # Stage table (offset right)
    ws.merge_range("D1:H1", "Pipeline by Stage", fmt_title)
    stage_cols = ["Stage","Count","Gross (QAR)","Net (QAR)"]
    write_header_row(ws, 2, stage_cols)
    ws.set_column(3, 3, 22)
    ws.set_column(4, 4, 8)
    ws.set_column(5, 5, 18)
    ws.set_column(6, 6, 18)
    for r, row in stage_df.reset_index(drop=True).iterrows():
        alt  = (r % 2 == 1)
        xl1  = 4 + r   # 1-based row of this stage cell (0-based = 3+r → 1-based = 4+r)
        ws.write(3+r, 3, row["Stage"], fmt_alt if alt else fmt_text)
        ws.write_formula(3+r, 4, f'=COUNTIF({_fp}!D3:D{_n},D{xl1})',            fmt_alt_num if alt else fmt_num)
        ws.write_formula(3+r, 5, f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!I3:I{_n})', fmt_alt_num if alt else fmt_num)
        ws.write_formula(3+r, 6, f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!J3:J{_n})', fmt_alt_num if alt else fmt_num)
    # Stage TOTAL row
    _sg_first = 4; _sg_last = 3 + len(stage_df)
    ws.write(3+len(stage_df), 3, "TOTAL", tot_lbl)
    ws.write_formula(3+len(stage_df), 4, f"=SUM(E{_sg_first}:E{_sg_last})", tot_fmt)
    ws.write_formula(3+len(stage_df), 5, f"=SUM(F{_sg_first}:F{_sg_last})", tot_fmt)
    ws.write_formula(3+len(stage_df), 6, f"=SUM(G{_sg_first}:G{_sg_last})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 — DU BREAKDOWN (BU mapped)
    # ════════════════════════════════════════════════════════════════════════
    du_ws = wb.add_worksheet("DU Breakdown")
    du_ws.set_zoom(90)
    du_ws.set_tab_color("#FF8C00")
    du_ws.merge_range("A1:G1", "Gross & Net Breakdown by BU / Delivery Unit", fmt_title)
    du_ws.set_row(0, 28)

    fmt_bu_hdr = wb.add_format({"bold":True,"bg_color":"#2E5FA3","font_color":"#FFFFFF",
                                 "border":1,"align":"left","font_size":11})
    fmt_bu_num = wb.add_format({"bold":True,"bg_color":"#D9E1F2","num_format":"#,##0","border":1,"align":"right"})
    fmt_bu_pct = wb.add_format({"bold":True,"bg_color":"#D9E1F2","num_format":"0.0","border":1,"align":"right"})
    fmt_bu_lbl = wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"})

    du_cols   = ["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    du_widths = [42, 52, 20, 20, 22]
    write_header_row(du_ws, 1, du_cols, du_widths)

    fmt_opp     = wb.add_format({"italic":True,"bg_color":"#FAFAFA","border":1,"align":"left","indent":2})
    fmt_opp_num = wb.add_format({"italic":True,"bg_color":"#FAFAFA","num_format":"#,##0","border":1,"align":"right"})

    # ── Pre-compute layout for formula-based subtotals ──────────────────────
    # Each BU has: 1 header row, then for each DU: 1 DU row + N opp rows
    # We pre-compute which 0-based row each element lands on so we can write
    # SUM formulas: opp rows hold raw values; DU row sums its opp rows;
    # BU row sums its DU rows; Grand Total sums all BU rows.
    _layout = []   # [(bu_name, bu_r0, [(du_name, du_r0, [opp_r0, ...])])]
    _pos = 0
    for bu_name, bu_grp in du_totals.groupby("BU"):
        bu_r0 = 2 + _pos; _pos += 1
        du_list = []
        for _, drow in bu_grp.iterrows():
            du_r0 = 2 + _pos; _pos += 1
            du_deals = du_exp[du_exp["DU"] == drow["DU"]].copy()
            opp_rows = []
            for _, deal in du_deals.iterrows():
                opp_rows.append(2 + _pos); _pos += 1
            du_list.append((drow, du_r0, opp_rows))
        _layout.append((bu_name, bu_r0, du_list))
    _grand_r0 = 2 + _pos   # Grand Total row (0-based)

    # ── Write all rows using formula references ───────────────────────────────
    all_du_rows = []   # collect DU row positions (1-based) for Grand Total
    for bu_name, bu_r0, du_list in _layout:
        # DU rows for this BU (used in BU SUM formula)
        _du_row_nums = [dr0+1 for (_, dr0, _) in du_list]  # 1-based
        # BU header row — SUM formula over its DU rows (pre-join to avoid f-string nesting)
        _bu_C = "+".join("C" + str(r) for r in _du_row_nums)
        _bu_D = "+".join("D" + str(r) for r in _du_row_nums)
        _bu_E = "+".join("E" + str(r) for r in _du_row_nums)
        du_ws.write(bu_r0, 0, bu_name, fmt_bu_lbl)
        du_ws.write(bu_r0, 1, "", fmt_bu_lbl)
        du_ws.write_formula(bu_r0, 2, "=" + _bu_C, fmt_bu_num)
        du_ws.write_formula(bu_r0, 3, "=" + _bu_D, fmt_bu_num)
        du_ws.write_formula(bu_r0, 4, "=" + _bu_E, fmt_bu_num)
        all_du_rows.extend(_du_row_nums)

        for drow, du_r0, opp_rows in du_list:
            alt = (du_r0 % 2 == 1)
            # DU row — SUM formula over its opp rows
            if opp_rows:
                _opp_sum = lambda col: f"=SUM({col}{min(opp_rows)+1}:{col}{max(opp_rows)+1})"
            else:
                _opp_sum = lambda col: f"=0"
            du_ws.write(du_r0, 0, "", fmt_alt if alt else fmt_text)
            du_ws.write(du_r0, 1, drow["DU"], fmt_alt if alt else fmt_text)
            du_ws.write_formula(du_r0, 2, _opp_sum("C"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 3, _opp_sum("D"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 4, _opp_sum("E"), fmt_alt_num if alt else fmt_num)

            du_deals = du_exp[du_exp["DU"] == drow["DU"]].copy()
            for opp_r0, (_, deal) in zip(opp_rows, du_deals.iterrows()):
                opp_label = f"  ↳  {deal['Lead/Opp Name']}"
                fore_net = deal["Net"] if str(deal.get("Forecasted","")).strip() == "Yes" else 0
                du_ws.write(opp_r0, 0, "", fmt_opp)
                du_ws.write(opp_r0, 1, opp_label, fmt_opp)
                du_ws.write_number(opp_r0, 2, deal["Gross"], fmt_opp_num)
                du_ws.write_number(opp_r0, 3, deal["Net"],   fmt_opp_num)
                du_ws.write_number(opp_r0, 4, fore_net,      fmt_opp_num)

    # Grand Total — SUM of all BU rows (which themselves SUM their DU/opp rows)
    _bu_row_nums = [br0+1 for (_, br0, _) in _layout]  # 1-based (renamed to avoid shadowing bu_r0)
    _gt_C = "+".join("C" + str(r) for r in _bu_row_nums)
    _gt_D = "+".join("D" + str(r) for r in _bu_row_nums)
    _gt_E = "+".join("E" + str(r) for r in _bu_row_nums)
    t = _grand_r0
    du_ws.write(t, 0, "GRAND TOTAL", tot_lbl)
    du_ws.write(t, 1, "", tot_lbl)
    du_ws.write_formula(t, 2, "=" + _gt_C, tot_fmt)
    du_ws.write_formula(t, 3, "=" + _gt_D, tot_fmt)
    du_ws.write_formula(t, 4, "=" + _gt_E, tot_fmt)
    r_out = _pos   # keep r_out in sync for downstream references

    # DU × Stage detail below
    du_ws.merge_range(t+2, 0, t+2, 5, "BU / DU × Stage Detail", fmt_title)
    du_stage = (
        du_exp.groupby(["BU","DU","Stage"])[["Gross","Net"]]
        .sum().reset_index().sort_values(["BU","DU","Net"], ascending=[True,True,False])
    )
    fore_du_stage = (
        du_exp[du_exp["Forecasted"]=="Yes"]
        .groupby(["DU","Stage"])["Net"].sum().reset_index()
        .rename(columns={"Net":"Forecasted Net"})
    )
    du_stage = du_stage.merge(fore_du_stage, on=["DU","Stage"], how="left").fillna({"Forecasted Net":0})
    ds_cols = ["BU","Delivery Unit","Stage","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    write_header_row(du_ws, t+3, ds_cols)
    _ds_data_start = t + 5   # 1-based first data row of DU×Stage detail
    for r, row in du_stage.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        du_ws.write(t+4+r, 0, row["BU"],   fmt_alt if alt else fmt_text)
        du_ws.write(t+4+r, 1, row["DU"],   fmt_alt if alt else fmt_text)
        du_ws.write(t+4+r, 2, row["Stage"],fmt_alt if alt else fmt_text)
        du_ws.write_number(t+4+r, 3, row["Gross"], fmt_alt_num if alt else fmt_num)
        du_ws.write_number(t+4+r, 4, row["Net"],   fmt_alt_num if alt else fmt_num)
        du_ws.write_number(t+4+r, 5, row["Forecasted Net"], fmt_alt_num if alt else fmt_num)
    # DU×Stage TOTAL row
    _ds_data_end = t + 4 + len(du_stage)   # 1-based last data row
    _ds_tot_r = t + 4 + len(du_stage)      # 0-based total row
    du_ws.write(_ds_tot_r, 0, "TOTAL", tot_lbl)
    du_ws.write(_ds_tot_r, 1, "", tot_lbl)
    du_ws.write(_ds_tot_r, 2, "", tot_lbl)
    du_ws.write_formula(_ds_tot_r, 3, f"=SUM(D{_ds_data_start}:D{_ds_data_end})", tot_fmt)
    du_ws.write_formula(_ds_tot_r, 4, f"=SUM(E{_ds_data_start}:E{_ds_data_end})", tot_fmt)
    du_ws.write_formula(_ds_tot_r, 5, f"=SUM(F{_ds_data_start}:F{_ds_data_end})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — FORECAST PER DU
    # ════════════════════════════════════════════════════════════════════════
    fd_ws = wb.add_worksheet("Forecast per DU")
    fd_ws.set_zoom(90)
    fd_ws.set_tab_color("#228B22")
    fd_ws.merge_range("A1:J1", f"Forecasted Pipeline by BU / Delivery Unit — {TODAY.strftime('%d %B %Y')}", fmt_title)
    fd_ws.set_row(0, 28)

    # ── Part 1: Summary (BU > DU > Quarter) ─────────────────────────────────
    fd_ws.merge_range("A2:J2", "Summary: Forecasted Net by BU / DU / Quarter", fmt_bu_hdr)
    fd_ws.set_row(2, 22)

    sum_cols   = ["BU","Delivery Unit","Quarter","Count","Gross (QAR)","Net (QAR)"]
    sum_widths = [42, 38, 10, 8, 20, 20]
    write_header_row(fd_ws, 3, sum_cols, sum_widths)

    # Pre-compute layout: BU header row + detail rows per BU
    _fds_layout = []   # [(bu_name, bu_xl_r, [detail_xl_r, ...])]   (0-based)
    _fds_pos = 4
    for bu_name, bu_sub in fore_du_summary.groupby("BU", sort=False):
        bu_xl_r = _fds_pos; _fds_pos += 1
        detail_rows = []
        for _ in bu_sub.itertuples():
            detail_rows.append(_fds_pos); _fds_pos += 1
        _fds_layout.append((bu_name, bu_xl_r, detail_rows))
    _fds_grand_r = _fds_pos

    for bu_name, bu_xl_r, detail_rows in _fds_layout:
        bu_sub = fore_du_summary[fore_du_summary["BU"] == bu_name]
        _fds_d = "+".join("D" + str(r+1) for r in detail_rows)
        _fds_e = "+".join("E" + str(r+1) for r in detail_rows)
        _fds_f = "+".join("F" + str(r+1) for r in detail_rows)
        fd_ws.write(bu_xl_r, 0, bu_name, fmt_bu_lbl)
        fd_ws.write(bu_xl_r, 1, f"Total: {len(bu_sub)} rows", fmt_bu_lbl)
        fd_ws.write(bu_xl_r, 2, "", fmt_bu_lbl)
        fd_ws.write_formula(bu_xl_r, 3, "=" + _fds_d, fmt_bu_num)
        fd_ws.write_formula(bu_xl_r, 4, "=" + _fds_e, fmt_bu_num)
        fd_ws.write_formula(bu_xl_r, 5, "=" + _fds_f, fmt_bu_num)
        for r_pos, (_, row) in zip(detail_rows, bu_sub.iterrows()):
            alt = (r_pos % 2 == 1)
            fd_ws.write(r_pos, 0, "", fmt_alt if alt else fmt_text)
            fd_ws.write(r_pos, 1, row["DU"],                 fmt_alt if alt else fmt_text)
            fd_ws.write(r_pos, 2, row["Closure Due Quarter"], fmt_alt if alt else fmt_text)
            fd_ws.write_number(r_pos, 3, row["Count"], fmt_alt_num if alt else fmt_num)
            fd_ws.write_number(r_pos, 4, row["Gross"], fmt_alt_num if alt else fmt_num)
            fd_ws.write_number(r_pos, 5, row["Net"],   fmt_alt_num if alt else fmt_num)

    # Grand Total summary — SUM of all BU header rows
    _bu_hdr_rows = [bu_xl_r for (_, bu_xl_r, _) in _fds_layout]
    _gt_D = "+".join("D" + str(r+1) for r in _bu_hdr_rows)
    _gt_E = "+".join("E" + str(r+1) for r in _bu_hdr_rows)
    _gt_F = "+".join("F" + str(r+1) for r in _bu_hdr_rows)
    ts = _fds_grand_r
    fd_ws.write(ts, 0, "GRAND TOTAL", tot_lbl)
    fd_ws.write(ts, 1, "", tot_lbl)
    fd_ws.write(ts, 2, "", tot_lbl)
    fd_ws.write_formula(ts, 3, "=" + _gt_D, tot_fmt)
    fd_ws.write_formula(ts, 4, "=" + _gt_E, tot_fmt)
    fd_ws.write_formula(ts, 5, "=" + _gt_F, tot_fmt)

    # ── Part 2: Deal-level detail ─────────────────────────────────────────
    det_start = ts + 2
    fd_ws.merge_range(det_start, 0, det_start, 9, "Detail: Forecasted Deals by BU / DU", fmt_bu_hdr)
    fd_ws.set_row(det_start, 22)

    det_cols   = ["BU","Delivery Unit","Account Name","Opportunity","Stage",
                  "Account Manager","Quarter","Gross (QAR)","Net (QAR)","Win Prob"]
    det_widths = [42, 38, 28, 36, 22, 24, 10, 18, 18, 12]
    write_header_row(fd_ws, det_start+1, det_cols, det_widths)
    fd_ws.set_column(6, 6, 10)
    fd_ws.set_column(7, 7, 18)
    fd_ws.set_column(8, 8, 18)
    fd_ws.set_column(9, 9, 12)

    # Pre-compute layout for detail section
    _fdd_layout = []   # [(bu_name, bu_xl_r, [detail_xl_r, ...])]
    _fdd_pos = det_start + 2
    for bu_name, bu_grp in fore_du_detail.groupby("BU", sort=False):
        bu_xl_r = _fdd_pos; _fdd_pos += 1
        detail_rows = []
        for _ in bu_grp.itertuples():
            detail_rows.append(_fdd_pos); _fdd_pos += 1
        _fdd_layout.append((bu_name, bu_xl_r, bu_grp, detail_rows))
    _fdd_grand_r = _fdd_pos

    for bu_name, bu_xl_r, bu_grp, detail_rows in _fdd_layout:
        _fdd_H = "+".join("H" + str(r+1) for r in detail_rows)
        _fdd_I = "+".join("I" + str(r+1) for r in detail_rows)
        fd_ws.write(bu_xl_r, 0, bu_name, fmt_bu_lbl)
        for cc in range(1, 10):
            fd_ws.write(bu_xl_r, cc, "", fmt_bu_lbl)
        fd_ws.write_formula(bu_xl_r, 7, "=" + _fdd_H, fmt_bu_num)
        fd_ws.write_formula(bu_xl_r, 8, "=" + _fdd_I, fmt_bu_num)
        for r_pos, (_, row) in zip(detail_rows, bu_grp.iterrows()):
            alt = (r_pos % 2 == 1)
            fmtx = fmt_grn if not alt else fmt_alt
            fmtn = fmt_grn_num if not alt else fmt_alt_num
            fd_ws.write(r_pos, 0, "", fmtx)
            fd_ws.write(r_pos, 1, str(row["DU"])                  if pd.notna(row["DU"])                  else "", fmtx)
            fd_ws.write(r_pos, 2, str(row["Account Name"])        if pd.notna(row["Account Name"])        else "", fmtx)
            fd_ws.write(r_pos, 3, str(row["Lead/Opp Name"])       if pd.notna(row["Lead/Opp Name"])       else "", fmtx)
            fd_ws.write(r_pos, 4, str(row["Stage"])               if pd.notna(row["Stage"])               else "", fmtx)
            fd_ws.write(r_pos, 5, str(row["Account Manager"])     if pd.notna(row["Account Manager"])     else "", fmtx)
            fd_ws.write(r_pos, 6, str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "", fmtx)
            fd_ws.write_number(r_pos, 7, row["Gross"], fmtn)
            fd_ws.write_number(r_pos, 8, row["Net"],   fmtn)
            fd_ws.write(r_pos, 9, str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "", fmtx)

    # Grand Total detail — SUM of all BU header rows
    _bu_hdr_H = "+".join("H" + str(bxr+1) for (_, bxr, _, _) in _fdd_layout)
    _bu_hdr_I = "+".join("I" + str(bxr+1) for (_, bxr, _, _) in _fdd_layout)
    td = _fdd_grand_r
    fd_ws.write(td, 0, "GRAND TOTAL", tot_lbl)
    for cc in range(1, 7):
        fd_ws.write(td, cc, "", tot_lbl)
    fd_ws.write_formula(td, 7, "=" + _bu_hdr_H, tot_fmt)
    fd_ws.write_formula(td, 8, "=" + _bu_hdr_I, tot_fmt)
    fd_ws.write(td, 9, "", tot_lbl)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — SECTOR & ACCOUNT MANAGER
    # ════════════════════════════════════════════════════════════════════════
    sa_ws = wb.add_worksheet("Sector & AM")
    sa_ws.set_zoom(90)
    sa_ws.set_tab_color("#228B22")
    sa_ws.merge_range("A1:E1", "Pipeline by Sector", fmt_title)
    sa_ws.set_row(0, 28)

    sec_cols = ["Sector","Count","Gross (QAR)","Net (QAR)"]
    write_header_row(sa_ws, 1, sec_cols, [24,8,20,20])
    for r, row in sector_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        sa_ws.write(2+r, 0, row["Sector"], fmt_alt if alt else fmt_text)
        sa_ws.write_number(2+r, 1, row["Count"],  fmt_alt_num if alt else fmt_num)
        sa_ws.write_number(2+r, 2, row["Gross"],  fmt_alt_num if alt else fmt_num)
        sa_ws.write_number(2+r, 3, row["Net"],    fmt_alt_num if alt else fmt_num)
    # Sector TOTAL row
    _sec_r1 = 3; _sec_rN = 2 + len(sector_df)
    sa_ws.write(2+len(sector_df), 0, "TOTAL", tot_lbl)
    sa_ws.write_formula(2+len(sector_df), 1, f"=SUM(B{_sec_r1}:B{_sec_rN})", tot_fmt)
    sa_ws.write_formula(2+len(sector_df), 2, f"=SUM(C{_sec_r1}:C{_sec_rN})", tot_fmt)
    sa_ws.write_formula(2+len(sector_df), 3, f"=SUM(D{_sec_r1}:D{_sec_rN})", tot_fmt)

    off = 2 + len(sector_df) + 1 + 2   # +1 for TOTAL row
    sa_ws.merge_range(off, 0, off, 4, "Pipeline by Account Manager", fmt_title)
    am_cols = ["Account Manager","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    write_header_row(sa_ws, off+1, am_cols, [28,8,20,20,22])
    sa_ws.set_column(4, 4, 22)
    for r, row in am_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        sa_ws.write(off+2+r, 0, row["Account Manager"], fmt_alt if alt else fmt_text)
        sa_ws.write_number(off+2+r, 1, row["Count"],         fmt_alt_num if alt else fmt_num)
        sa_ws.write_number(off+2+r, 2, row["Gross"],         fmt_alt_num if alt else fmt_num)
        sa_ws.write_number(off+2+r, 3, row["Net"],           fmt_alt_num if alt else fmt_num)
        sa_ws.write_number(off+2+r, 4, row["Forecasted Net"],fmt_alt_num if alt else fmt_num)
    # AM TOTAL row
    _am_r1 = off + 3; _am_rN = off + 2 + len(am_df)
    sa_ws.write(off+2+len(am_df), 0, "TOTAL", tot_lbl)
    sa_ws.write_formula(off+2+len(am_df), 1, f"=SUM(B{_am_r1}:B{_am_rN})", tot_fmt)
    sa_ws.write_formula(off+2+len(am_df), 2, f"=SUM(C{_am_r1}:C{_am_rN})", tot_fmt)
    sa_ws.write_formula(off+2+len(am_df), 3, f"=SUM(D{_am_r1}:D{_am_rN})", tot_fmt)
    sa_ws.write_formula(off+2+len(am_df), 4, f"=SUM(E{_am_r1}:E{_am_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 — QUARTERLY & PROBABILITY
    # ════════════════════════════════════════════════════════════════════════
    qp_ws = wb.add_worksheet("Quarterly & Probability")
    qp_ws.set_zoom(90)
    qp_ws.set_tab_color("#DAA520")
    qp_ws.merge_range("A1:E1", "Quarterly Close Plan", fmt_title)
    qp_ws.set_row(0, 28)

    q_cols = ["Quarter","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    write_header_row(qp_ws, 1, q_cols, [12,8,20,20,22])
    for r, row in q_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        qp_ws.write(2+r, 0, row["Closure Due Quarter"], fmt_alt if alt else fmt_text)
        qp_ws.write_number(2+r, 1, row["Count"],         fmt_alt_num if alt else fmt_num)
        qp_ws.write_number(2+r, 2, row["Gross"],         fmt_alt_num if alt else fmt_num)
        qp_ws.write_number(2+r, 3, row["Net"],           fmt_alt_num if alt else fmt_num)
        qp_ws.write_number(2+r, 4, row["Forecasted Net"],fmt_alt_num if alt else fmt_num)
    # Quarter TOTAL row
    _q_r1 = 3; _q_rN = 2 + len(q_df)
    qp_ws.write(2+len(q_df), 0, "TOTAL", tot_lbl)
    qp_ws.write_formula(2+len(q_df), 1, f"=SUM(B{_q_r1}:B{_q_rN})", tot_fmt)
    qp_ws.write_formula(2+len(q_df), 2, f"=SUM(C{_q_r1}:C{_q_rN})", tot_fmt)
    qp_ws.write_formula(2+len(q_df), 3, f"=SUM(D{_q_r1}:D{_q_rN})", tot_fmt)
    qp_ws.write_formula(2+len(q_df), 4, f"=SUM(E{_q_r1}:E{_q_rN})", tot_fmt)

    off2 = 2 + len(q_df) + 1 + 2   # +1 for TOTAL row
    qp_ws.merge_range(off2, 0, off2, 3, "Pipeline by Winning Probability", fmt_title)
    pb_cols = ["Winning Probability","Count","Gross (QAR)","Net (QAR)"]
    write_header_row(qp_ws, off2+1, pb_cols, [22,8,20,20])
    for r, row in prob_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        qp_ws.write(off2+2+r, 0, row["Winning Probability"], fmt_alt if alt else fmt_text)
        qp_ws.write_number(off2+2+r, 1, row["Count"], fmt_alt_num if alt else fmt_num)
        qp_ws.write_number(off2+2+r, 2, row["Gross"], fmt_alt_num if alt else fmt_num)
        qp_ws.write_number(off2+2+r, 3, row["Net"],   fmt_alt_num if alt else fmt_num)
    # Probability TOTAL row
    _pb_r1 = off2 + 3; _pb_rN = off2 + 2 + len(prob_df)
    qp_ws.write(off2+2+len(prob_df), 0, "TOTAL", tot_lbl)
    qp_ws.write_formula(off2+2+len(prob_df), 1, f"=SUM(B{_pb_r1}:B{_pb_rN})", tot_fmt)
    qp_ws.write_formula(off2+2+len(prob_df), 2, f"=SUM(C{_pb_r1}:C{_pb_rN})", tot_fmt)
    qp_ws.write_formula(off2+2+len(prob_df), 3, f"=SUM(D{_pb_r1}:D{_pb_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 — FORECAST
    # ════════════════════════════════════════════════════════════════════════
    fw = wb.add_worksheet("Forecast")
    fw.set_zoom(90)
    fw.set_tab_color("#228B22")
    fw.merge_range("A1:J1", f"Forecasted Deals — {TODAY.strftime('%d %B %Y')}", fmt_title)
    fw.set_row(0, 28)

    fore_cols = ["Account Name","Opportunity","Stage","Account Manager","Sector",
                 "Gross (QAR)","Net (QAR)","Win Probability","Quarter","Est. Close Date"]
    fore_widths = [30,36,22,24,18,18,18,14,10,16]
    write_header_row(fw, 1, fore_cols, fore_widths)

    for r, row in fore_df.reset_index(drop=True).iterrows():
        fw.write(2+r, 0, str(row["Account Name"])  if pd.notna(row["Account Name"])  else "", fmt_grn)
        fw.write(2+r, 1, str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "", fmt_grn)
        fw.write(2+r, 2, str(row["Stage"])          if pd.notna(row["Stage"])          else "", fmt_grn)
        fw.write(2+r, 3, str(row["Account Manager"])if pd.notna(row["Account Manager"])else "", fmt_grn)
        fw.write(2+r, 4, str(row["Sector"])         if pd.notna(row["Sector"])         else "", fmt_grn)
        fw.write_number(2+r, 5, row["Total Gross"], fmt_grn_num)
        fw.write_number(2+r, 6, row["Total Net"],   fmt_grn_num)
        fw.write(2+r, 7, str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "", fmt_grn)
        fw.write(2+r, 8, str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "", fmt_grn)
        if pd.notna(row["Est. Close Date"]):
            fw.write_datetime(2+r, 9, row["Est. Close Date"].to_pydatetime(), fmt_grn_dt)
        else:
            fw.write_blank(2+r, 9, None, fmt_grn_dt)

    # TOTAL row with SUM formulas
    t2 = 2 + len(fore_df)
    _fc_r1 = 3; _fc_rN = 2 + len(fore_df)
    fw.write(t2, 0, "TOTAL", tot_lbl)
    fw.write_formula(t2, 5, f"=SUM(F{_fc_r1}:F{_fc_rN})", tot_fmt)
    fw.write_formula(t2, 6, f"=SUM(G{_fc_r1}:G{_fc_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 — OVERDUE
    # ════════════════════════════════════════════════════════════════════════
    ow = wb.add_worksheet("Overdue Deals")
    ow.set_zoom(90)
    ow.set_tab_color("#FF0000")
    ow.merge_range("A1:H1", f"Overdue Deals (Close Date Passed) — {TODAY.strftime('%d %B %Y')}", fmt_title)
    ow.set_row(0, 28)

    ov_cols   = ["Account Name","Opportunity","Stage","Account Manager",
                 "Net (QAR)","Est. Close Date","Win Probability","Quarter"]
    ov_widths = [30,36,22,24,18,16,14,10]
    write_header_row(ow, 1, ov_cols, ov_widths)

    for r, row in overdue_df.reset_index(drop=True).iterrows():
        ow.write(2+r, 0, str(row["Account Name"])  if pd.notna(row["Account Name"])  else "", fmt_red)
        ow.write(2+r, 1, str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "", fmt_red)
        ow.write(2+r, 2, str(row["Stage"])          if pd.notna(row["Stage"])          else "", fmt_red)
        ow.write(2+r, 3, str(row["Account Manager"])if pd.notna(row["Account Manager"])else "", fmt_red)
        ow.write_number(2+r, 4, row["Total Net"],   fmt_red_num)
        if pd.notna(row["Est. Close Date"]):
            ow.write_datetime(2+r, 5, row["Est. Close Date"].to_pydatetime(), fmt_red_dt)
        else:
            ow.write_blank(2+r, 5, None, fmt_red_dt)
        ow.write(2+r, 6, str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "", fmt_red)
        ow.write(2+r, 7, str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "", fmt_red)
    # Overdue TOTAL row
    _ov_r1 = 3; _ov_rN = 2 + len(overdue_df)
    ow.write(2+len(overdue_df), 0, "TOTAL", tot_lbl)
    ow.write_formula(2+len(overdue_df), 4, f"=SUM(E{_ov_r1}:E{_ov_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 7 — FULL PIPELINE
    # ════════════════════════════════════════════════════════════════════════
    pw = wb.add_worksheet("Full Pipeline")
    pw.set_zoom(85)
    pw.set_tab_color("#6495ED")
    pw.merge_range("A1:Q1", "Full Pipeline — All Opportunities", fmt_title)
    pw.set_row(0, 28)
    pw.freeze_panes(2, 0)

    full_cols = ["#","Account Name","Opportunity","Stage","Account Manager","Sector",
                 "BU","DU","Gross (QAR)","Net (QAR)","Win Prob","Forecasted",
                 "Strategic","Quarter","Est. Close Date","Source","Overdue"]
    full_widths = [5,28,36,22,22,16,30,36,18,18,12,12,10,10,16,16,8]
    write_header_row(pw, 1, full_cols, full_widths)
    pw.autofilter(1, 0, 1+len(full_df), len(full_cols)-1)

    for r, row in full_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        is_overdue = bool(row["Overdue"])
        t_fmt     = fmt_red     if is_overdue else (fmt_alt     if alt else fmt_text)
        n_fmt     = fmt_red_num if is_overdue else (fmt_alt_num if alt else fmt_num)
        d_fmt     = fmt_red_dt  if is_overdue else fmt_date

        pw.write(2+r, 0,  str(row["SNo."]) if pd.notna(row["SNo."]) else "", t_fmt)
        pw.write(2+r, 1,  str(row["Account Name"])   if pd.notna(row["Account Name"])   else "", t_fmt)
        pw.write(2+r, 2,  str(row["Lead/Opp Name"])  if pd.notna(row["Lead/Opp Name"])  else "", t_fmt)
        pw.write(2+r, 3,  str(row["Stage"])           if pd.notna(row["Stage"])           else "", t_fmt)
        pw.write(2+r, 4,  str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "", t_fmt)
        pw.write(2+r, 5,  str(row["Sector"])          if pd.notna(row["Sector"])          else "", t_fmt)
        pw.write(2+r, 6,  str(row["BU"])              if pd.notna(row["BU"])              else "", t_fmt)
        pw.write(2+r, 7,  str(row["DU"])              if pd.notna(row["DU"])              else "", t_fmt)
        pw.write_number(2+r, 8,  row["Total Gross"], n_fmt)
        pw.write_number(2+r, 9,  row["Total Net"],   n_fmt)
        pw.write(2+r, 10, str(row["Winning Probability"])  if pd.notna(row["Winning Probability"])  else "", t_fmt)
        pw.write(2+r, 11, str(row["Forecasted"])           if pd.notna(row["Forecasted"])           else "", t_fmt)
        pw.write(2+r, 12, str(row["Strategic Opportunity"])if pd.notna(row["Strategic Opportunity"])else "", t_fmt)
        pw.write(2+r, 13, str(row["Closure Due Quarter"])  if pd.notna(row["Closure Due Quarter"])  else "", t_fmt)
        if pd.notna(row["Est. Close Date"]):
            pw.write_datetime(2+r, 14, row["Est. Close Date"].to_pydatetime(), d_fmt)
        else:
            pw.write_blank(2+r, 14, None, d_fmt)
        pw.write(2+r, 15, str(row["Source of Opportunity"]) if pd.notna(row["Source of Opportunity"]) else "", t_fmt)
        pw.write(2+r, 16, "YES" if is_overdue else "", t_fmt)
    # Full Pipeline TOTAL row
    _fp_r1 = 3; _fp_rN = 2 + len(full_df)
    pw.write(_fp_rN, 0, "TOTAL", tot_lbl)
    pw.write_formula(_fp_rN, 8, f"=SUM(I{_fp_r1}:I{_fp_rN})", tot_fmt)
    pw.write_formula(_fp_rN, 9, f"=SUM(J{_fp_r1}:J{_fp_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 8 — PIPELINE BREAKDOWN (styled, one row per DU per opportunity)
    # ════════════════════════════════════════════════════════════════════════
    CLR = {
        "title_bg":    "1F3864",
        "hdr_deal":    "1F3864",
        "hdr_du":      "17375E",
        "hdr_finance": "1F4E79",
        "hdr_other":   "2E5FA3",
        "bu_fill":     "EDF2F9",
        "du_fill":     "E4ECF7",
        "num_fill":    "EBF5FB",
        "tot_fill":    "D5E8F5",
        "date_fill":   "FFF2CC",
        "alt_a":       "F5F8FF",
        "alt_b":       "FFFFFF",
    }
    fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_du      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_du"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fmt_title_bd = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+CLR["title_bg"],"align":"center","valign":"vcenter"})

    def _bd_fmt(wb, bg, top, num_fmt=None, extra=None):
        d = {"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
        if num_fmt:
            d["num_format"] = num_fmt
            d["align"] = "right"
        if extra:
            d.update(extra)
        return wb.add_format(d)

    ft_a_first  = _bd_fmt(wb, CLR["alt_a"], 2)
    ft_a_next   = _bd_fmt(wb, CLR["alt_a"], 1)
    ft_b_first  = _bd_fmt(wb, CLR["alt_b"], 2)
    ft_b_next   = _bd_fmt(wb, CLR["alt_b"], 1)
    fn_a_first  = _bd_fmt(wb, CLR["alt_a"], 2, "#,##0")
    fn_a_next   = _bd_fmt(wb, CLR["alt_a"], 1, "#,##0")
    fn_b_first  = _bd_fmt(wb, CLR["alt_b"], 2, "#,##0")
    fn_b_next   = _bd_fmt(wb, CLR["alt_b"], 1, "#,##0")
    fb_a_first  = _bd_fmt(wb, CLR["bu_fill"], 2)
    fb_a_next   = _bd_fmt(wb, CLR["bu_fill"], 1)
    fb_b_first  = _bd_fmt(wb, CLR["bu_fill"], 2)
    fb_b_next   = _bd_fmt(wb, CLR["bu_fill"], 1)
    fd_a_first  = _bd_fmt(wb, CLR["du_fill"], 2)
    fd_a_next   = _bd_fmt(wb, CLR["du_fill"], 1)
    fd_b_first  = _bd_fmt(wb, CLR["du_fill"], 2)
    fd_b_next   = _bd_fmt(wb, CLR["du_fill"], 1)
    fx_a_first  = _bd_fmt(wb, CLR["num_fill"], 2, "#,##0")
    fx_a_next   = _bd_fmt(wb, CLR["num_fill"], 1, "#,##0")
    fx_b_first  = _bd_fmt(wb, CLR["num_fill"], 2, "#,##0")
    fx_b_next   = _bd_fmt(wb, CLR["num_fill"], 1, "#,##0")
    ft_tot      = wb.add_format({"bg_color":"#"+CLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
    ft_tot_blank= wb.add_format({"bg_color":"#"+CLR["num_fill"],"border":1,"font_size":9})
    ft_date     = wb.add_format({"bg_color":"#"+CLR["date_fill"],"border":1,"align":"center","num_format":"DD-MMM-YYYY","font_size":9})
    ft_date_blank = wb.add_format({"bg_color":"#"+CLR["alt_a"],"border":1,"font_size":9})

    bw = wb.add_worksheet("Pipeline Breakdown")
    bw.set_zoom(85)
    bw.set_tab_color("#9370DB")

    bd_output_cols = [
        ("SNo.",                  6,  "deal"),
        ("Account Name",         24,  "deal"),
        ("Lead/Opp Name",        36,  "deal"),
        ("BU",                   36,  "du"),
        ("DU",                   34,  "du"),
        ("Gross (breakdown)",    16,  "finance"),
        ("Net (breakdown)",      16,  "finance"),
        ("Total Gross",          15,  "finance"),
        ("Total Net",            15,  "finance"),
        ("Stage",                36,  "other"),
        ("Account Manager",      22,  "other"),
        ("Sector",               16,  "other"),
        ("Closure Due Quarter",   9,  "other"),
        ("Winning Probability",  10,  "other"),
        ("Forecasted",           10,  "other"),
        ("Strategic Opportunity",10,  "other"),
        ("Est. Close Date",      14,  "other"),
    ]
    hdr_fmt_map = {"deal": fh_deal, "du": fh_du, "finance": fh_finance, "other": fh_other}
    ncols = len(bd_output_cols)
    bw.merge_range(0, 0, 0, ncols - 1, "Pipeline — Expanded by Delivery Unit", fmt_title_bd)
    bw.set_row(0, 28)
    for c, (col_name, col_w, col_type) in enumerate(bd_output_cols):
        bw.write(1, c, col_name, hdr_fmt_map[col_type])
        bw.set_column(c, c, col_w)
    bw.set_row(1, 28)
    bw.freeze_panes(2, 0)

    bd_exp = _expand_deals(df, "Lead/Opp Name")
    bw.autofilter(1, 0, 1 + len(bd_exp), ncols - 1)

    col_map = {name: idx for idx, (name, _, __) in enumerate(bd_output_cols)}
    g_col = col_map["Gross (breakdown)"]   # column index for SUM formula
    n_col = col_map["Net (breakdown)"]

    # Pre-compute Excel row range per deal (xl_r = 2 + r_pos, Excel row = xl_r + 1)
    deal_rows = {}
    for r_pos, (_, row) in enumerate(bd_exp.iterrows()):
        didx = row["_deal_idx"]
        xl_r = 2 + r_pos
        if didx not in deal_rows:
            deal_rows[didx] = [xl_r, xl_r]
        else:
            deal_rows[didx][1] = xl_r

    prev_deal_idx = None
    alt_toggle = False
    for r_pos, (_, row) in enumerate(bd_exp.iterrows()):
        didx     = row["_deal_idx"]
        is_first = bool(row["_is_first"])
        if didx != prev_deal_idx:
            alt_toggle = not alt_toggle
            prev_deal_idx = didx
        alt = alt_toggle
        ft  = (ft_a_first if is_first else ft_a_next) if alt else (ft_b_first if is_first else ft_b_next)
        fbu = (fb_a_first if is_first else fb_a_next) if alt else (fb_b_first if is_first else fb_b_next)
        fdu = (fd_a_first if is_first else fd_a_next) if alt else (fd_b_first if is_first else fd_b_next)
        fxn = (fx_a_first if is_first else fx_a_next) if alt else (fx_b_first if is_first else fx_b_next)
        xl_r = 2 + r_pos

        def _ws(col_idx, val, fmt):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                bw.write_blank(xl_r, col_idx, None, fmt)
            else:
                bw.write(xl_r, col_idx, str(val), fmt)

        def _wn(col_idx, val, fmt):
            v = _parse_num(val) if not isinstance(val, (int, float)) else val
            if v is None or (isinstance(v, float) and pd.isna(v)):
                bw.write_blank(xl_r, col_idx, None, fmt)
            else:
                bw.write_number(xl_r, col_idx, v, fmt)

        # SNo. — first row only (sequence number)
        _ws(col_map["SNo."], row.get("SNo.") if is_first else None, ft)
        # Deal-level cols — replicated on every DU row
        _ws(col_map["Account Name"],  row.get("Account Name"),  ft)
        _ws(col_map["Lead/Opp Name"], row.get("Lead/Opp Name"), ft)
        # BU / DU — always
        _ws(col_map["BU"], row.get("BU_exp"), fbu)
        _ws(col_map["DU"], row.get("DU_exp"), fdu)
        # Breakdown numbers — always
        _wn(col_map["Gross (breakdown)"], row.get("Gross_exp"), fxn)
        _wn(col_map["Net (breakdown)"],   row.get("Net_exp"),   fxn)
        # Total Gross / Net — SUM formula on first row, blank on others
        if is_first:
            r0, r1 = deal_rows[didx]
            gc = chr(65 + g_col); nc = chr(65 + n_col)
            bw.write_formula(xl_r, col_map["Total Gross"], f"=SUM({gc}{r0+1}:{gc}{r1+1})", ft_tot)
            bw.write_formula(xl_r, col_map["Total Net"],   f"=SUM({nc}{r0+1}:{nc}{r1+1})", ft_tot)
        else:
            bw.write_blank(xl_r, col_map["Total Gross"], None, ft_tot_blank)
            bw.write_blank(xl_r, col_map["Total Net"],   None, ft_tot_blank)
        # Other deal-level cols — replicated on every DU row
        _ws(col_map["Stage"],                 row.get("Stage"),                 ft)
        _ws(col_map["Account Manager"],       row.get("Account Manager"),       ft)
        _ws(col_map["Sector"],                row.get("Sector"),                ft)
        _ws(col_map["Closure Due Quarter"],   row.get("Closure Due Quarter"),   ft)
        _ws(col_map["Winning Probability"],   row.get("Winning Probability"),   ft)
        _ws(col_map["Forecasted"],            row.get("Forecasted"),            ft)
        _ws(col_map["Strategic Opportunity"], row.get("Strategic Opportunity"), ft)
        # Est. Close Date — replicated on every row
        cd_idx = col_map["Est. Close Date"]
        cd_val = row.get("Est. Close Date")
        if pd.notna(cd_val):
            try:
                bw.write_datetime(xl_r, cd_idx, pd.Timestamp(cd_val).to_pydatetime(), ft_date)
            except Exception:
                bw.write(xl_r, cd_idx, str(cd_val), ft_date)
        else:
            bw.write_blank(xl_r, cd_idx, None, ft_date_blank)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 9 — BOOK3 MAPPING
    # ════════════════════════════════════════════════════════════════════════
    fmt_matched   = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"})
    fmt_matched_n = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"})
    fmt_partial   = wb.add_format({"bg_color":"#FFF2CC","border":1,"align":"left"})
    fmt_partial_n = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right"})
    fmt_nomatch   = wb.add_format({"bg_color":"#FFE0E0","border":1,"align":"left"})
    fmt_nomatch_n = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right"})
    fmt_month_hdr = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E5FA3",
                                   "border":1,"align":"center","text_wrap":True})
    fmt_neg_num   = wb.add_format({"num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
    fmt_neg_grn   = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
    fmt_neg_yel   = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
    fmt_neg_red   = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})

    mw = wb.add_worksheet("Book3 Mapping")
    mw.set_zoom(80)
    mw.set_tab_color("#FF8C00")
    mw.freeze_panes(3, 4)

    title_cols = 4 + len(MONTHS) + 1 + 6 + 6  # BU/Type/Name/Total + months + pipeline cols + awarded cols
    mw.merge_range(0, 0, 0, title_cols-1,
                   f"Book3 ↔ Pipeline & Awarded Mapping — {TODAY.strftime('%d %B %Y')}", fmt_title)
    mw.set_row(0, 28)

    # Legend row
    mw.write(1, 0, "🟢 Strong match (≥0.70)", fmt_matched)
    mw.write(1, 1, "🟡 Partial match (0.55–0.69)", fmt_partial)
    mw.write(1, 2, "🔴 No match (<0.55)", fmt_nomatch)
    mw.write(1, 3, "", fmt_text)
    mw.set_row(1, 18)

    # Headers — row 2
    book3_cols  = ["Book3 BU","Book3 Project Type","Book3 Project Name","Book3 Grand Total (QAR)"]
    month_cols  = [f"{m}" for m in MONTHS]
    pipe_cols   = ["Pipeline Match","Match Score","Pipeline Gross","Pipeline Net","Pipeline Stage","Pipeline AM"]
    award_cols  = ["Awarded Match","Match Score","Awarded Gross","Awarded Net","Awarded Stage","Awarded AM"]
    all_cols    = book3_cols + month_cols + pipe_cols + award_cols
    col_widths  = [36,22,42,20] + [10]*len(MONTHS) + [36,10,18,18,22,22] + [36,10,18,18,22,22]

    for c, (col, w) in enumerate(zip(all_cols, col_widths)):
        if col in month_cols:
            mw.write(2, c, col, fmt_month_hdr)
        else:
            mw.write(2, c, col, fmt_header)
        mw.set_column(c, c, w)
    mw.set_row(2, 20)
    mw.autofilter(2, 0, 2+len(mapping_df), len(all_cols)-1)

    for r, row in mapping_df.reset_index(drop=True).iterrows():
        pipe_score  = row["Pipeline Score"]
        award_score = row["Awarded Score"]
        best_score  = max(pipe_score, award_score)

        if best_score >= 0.70:
            ft = fmt_matched;  fn = fmt_matched_n;  fn_neg = fmt_neg_grn
        elif best_score >= 0.55:
            ft = fmt_partial;  fn = fmt_partial_n;  fn_neg = fmt_neg_yel
        else:
            ft = fmt_nomatch;  fn = fmt_nomatch_n;  fn_neg = fmt_neg_red

        c = 0
        mw.write(3+r, c, str(row["Book3 BU"])           or "", ft);  c+=1
        mw.write(3+r, c, str(row["Book3 Project Type"])  or "", ft);  c+=1
        mw.write(3+r, c, str(row["Book3 Project Name"])  or "", ft);  c+=1
        mw.write_number(3+r, c, row["Book3 Grand Total"], fn_neg);    c+=1
        for m in MONTHS:
            val = row.get(f"Book3 {m}", 0)
            mw.write_number(3+r, c, val, fn_neg); c+=1
        # Pipeline
        mw.write(3+r, c, str(row["Pipeline Match"])      or "", ft);  c+=1
        mw.write_number(3+r, c, row["Pipeline Score"],   fn);         c+=1
        mw.write_number(3+r, c, row["Pipeline Gross (QAR)"], fn);     c+=1
        mw.write_number(3+r, c, row["Pipeline Net (QAR)"],   fn);     c+=1
        mw.write(3+r, c, str(row["Pipeline Stage"])      or "", ft);  c+=1
        mw.write(3+r, c, str(row["Pipeline AM"])         or "", ft);  c+=1
        # Awarded
        mw.write(3+r, c, str(row["Awarded Match"])       or "", ft);  c+=1
        mw.write_number(3+r, c, row["Awarded Score"],    fn);         c+=1
        mw.write_number(3+r, c, row["Awarded Gross (QAR)"], fn);      c+=1
        mw.write_number(3+r, c, row["Awarded Net (QAR)"],   fn);      c+=1
        mw.write(3+r, c, str(row["Awarded Stage"])       or "", ft);  c+=1
        mw.write(3+r, c, str(row["Awarded AM"])          or "", ft);  c+=1

    # Write all sheets to the writer (register them)
    for name, ws_obj in [
        ("Summary", ws), ("DU Breakdown", du_ws), ("Forecast per DU", fd_ws),
        ("Sector & AM", sa_ws), ("Quarterly & Probability", qp_ws),
        ("Forecast", fw), ("Overdue Deals", ow), ("Full Pipeline", pw),
        ("Pipeline Breakdown", bw), ("Book3 Mapping", mw),
    ]:
        writer.sheets[name] = ws_obj

print(f"\nReport saved to:\n  {OUT_FILE}\n")
