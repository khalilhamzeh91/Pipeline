"""
Account Manager Pipeline Report Generator
Based on data (2) (1).xlsx — Capability Sales / AM view with monthly breakdown.
Usage: python generate_am_pipeline_report.py
"""

import pandas as pd
import re
import warnings
from datetime import date
warnings.filterwarnings("ignore")

INPUT_FILE = r"C:\Users\khali\Downloads\data (2) (1).xlsx"
COA_FILE   = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
TODAY      = date.today()
OUT_FILE   = rf"C:\Users\khali\Downloads\AM_Pipeline_Report_{TODAY}.xlsx"

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

STAGE_SHORT = {
    "Stage 1: Assessment & Qualification":                    "S1 - Assessment",
    "Stage 2: Discovery & Scoping":                           "S2 - Discovery",
    "Stage 3.1: RFP & BID Qualification":                     "S3.1 - RFP",
    "Stage 3.2: Solution Development & Proposal Submission":  "S3.2 - Solution Dev",
    "Stage 4: Technical Evaluation By Customer":              "S4 - Tech Eval",
    "Stage 5: Resolution/Financial Negotiation":              "S5 - Negotiation",
}

# ── LOAD & CLEAN ─────────────────────────────────────────────────────────────
df = pd.read_excel(INPUT_FILE, sheet_name="Export")
df.columns = df.columns.str.strip()
df = df.dropna(subset=["Stage"])
df["Total Gross"]    = pd.to_numeric(df["Total Gross"],  errors="coerce").fillna(0)
df["Total Net"]      = pd.to_numeric(df["Total Net"],    errors="coerce").fillna(0)
df["Est. Close Date"] = pd.to_datetime(df["Est. Close Date"], errors="coerce")
df["Overdue"]         = (df["Est. Close Date"] < pd.Timestamp(TODAY))
df["Stage_Short"]     = df["Stage"].map(STAGE_SHORT).fillna(df["Stage"])

for m in MONTHS:
    if m in df.columns:
        df[m] = pd.to_numeric(df[m], errors="coerce").fillna(0)
    else:
        df[m] = 0

# ── CHARTER OF ACCOUNTS — DU → BU MAPPING ────────────────────────────────────
coa = pd.read_excel(COA_FILE)
coa.columns = coa.columns.str.strip()
coa["_code"] = coa["DU"].str.extract(r"(\d{6})")
DU_TO_BU = coa.dropna(subset=["_code"]).set_index("_code")["BU"].to_dict()

def du_to_bu(du_str):
    m = re.match(r"(\d{6})", str(du_str).strip())
    if m:
        return DU_TO_BU.get(m.group(1), "Unknown")
    return "Unknown"

# ── HELPERS ───────────────────────────────────────────────────────────────────
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

def _clean_am_list(value):
    """Deduplicate and normalize AM names from newline-separated field."""
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return []
    seen = {}
    for name in str(value).split("\n"):
        name = name.strip()
        if not name:
            continue
        if "khalil" in name.lower():   name = "Khalil Hamzeh"
        elif "yazan" in name.lower():  name = "Yazan Al Razem"
        seen[name] = None
    return list(seen.keys())

# ── EXPLODE BY AM (Capability Sales) ─────────────────────────────────────────
am_rows = []
for deal_idx, row in df.iterrows():
    ams = _clean_am_list(row.get("Capability Sales", ""))
    if not ams:
        ams = ["Unassigned"]
    for i, am in enumerate(ams):
        nr = {c: row[c] for c in df.columns}
        nr["AM_exp"]    = am
        nr["_is_first"] = (i == 0)
        nr["_am_count"] = len(ams)
        nr["_deal_idx"] = deal_idx
        am_rows.append(nr)
am_exp = pd.DataFrame(am_rows)

# ── EXPLODE BY DU (for DU breakdown) ─────────────────────────────────────────
du_rows = []
for _, row in df.iterrows():
    dus   = str(row["DU"]).split("\n")                                   if pd.notna(row["DU"])                   else ["Unknown"]
    gross = str(row["Gross (Breakdown)"]).replace(",","").split("\n")    if pd.notna(row.get("Gross (Breakdown)","")) else ["0"]
    net   = str(row["Net (breakdown)"]).replace(",","").split("\n")      if pd.notna(row.get("Net (breakdown)",""))  else ["0"]
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
            "BU": du_to_bu(du), "DU": du,
            "Gross": g_val, "Net": n_val,
            "Forecasted":          str(row.get("Forecasted","")).strip(),
            "Stage":               row.get("Stage",""),
            "Sector":              row.get("Sector",""),
            "Closure Due Quarter": row.get("Closure Due Quarter",""),
            "Account Name":        row.get("Account Name",""),
            "Lead/Opp Name":       row.get("Lead/Opp Name",""),
            "Winning Probability": row.get("Winning Probability",""),
        })
du_exp_df = pd.DataFrame(du_rows)

# ── SUMMARY TABLES ────────────────────────────────────────────────────────────
# 1. KPIs
total_gross    = df["Total Gross"].sum()
total_net      = df["Total Net"].sum()
fore_net       = df[df["Forecasted"] == "Yes"]["Total Net"].sum()
fore_gross     = df[df["Forecasted"] == "Yes"]["Total Gross"].sum()
strategic_cnt  = int((df["Strategic Opportunity"] == "Yes").sum())
overdue_cnt    = int(df["Overdue"].sum())

# 2. By Stage
stage_df = (
    df.groupby("Stage_Short")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().rename(columns={"Stage_Short":"Stage"})
)
stage_order = list(STAGE_SHORT.values())
stage_df["_ord"] = stage_df["Stage"].map({s:i for i,s in enumerate(stage_order)})
stage_df = stage_df.sort_values("_ord").drop(columns="_ord")

# 3. By AM (Capability Sales) — exploded
am_agg = (
    am_exp.groupby("AM_exp")
    .agg(Count=("Lead/Opp Name","nunique"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().rename(columns={"AM_exp":"Account Manager"})
    .sort_values("Net", ascending=False)
)
# Forecasted per AM
fore_am_agg = (
    am_exp[am_exp["Forecasted"]=="Yes"]
    .groupby("AM_exp")["Total Net"].sum().reset_index()
    .rename(columns={"AM_exp":"Account Manager","Total Net":"Forecasted Net"})
)
am_agg = am_agg.merge(fore_am_agg, on="Account Manager", how="left").fillna({"Forecasted Net":0})

# 4. By Quarter
q_df = (
    df.groupby("Closure Due Quarter")
    .agg(Count=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().sort_values("Closure Due Quarter")
)
fore_q = df[df["Forecasted"]=="Yes"].groupby("Closure Due Quarter")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
q_df = q_df.merge(fore_q, on="Closure Due Quarter", how="left").fillna({"Forecasted Net":0})

# 5. Monthly Pipeline (sum of month columns across all deals)
monthly_totals = {m: float(df[m].sum()) for m in MONTHS}

# 6. DU Breakdown
du_totals = (
    du_exp_df.groupby(["BU","DU"])[["Gross","Net"]]
    .sum().reset_index().sort_values(["BU","Net"], ascending=[True,False])
)
fore_du = du_exp_df[du_exp_df["Forecasted"]=="Yes"].groupby("DU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"})
du_totals = du_totals.merge(fore_du, on="DU", how="left").fillna({"Forecasted Net":0})

# 7. Full pipeline
full_df = df[[
    "SNo.","Account Name","Lead/Opp Name","Stage_Short","Capability Sales","Sector",
    "BU","DU","Total Gross","Total Net","Winning Probability","Forecasted",
    "Strategic Opportunity","Closure Due Quarter","Est. Close Date","Source of Opportunity","Overdue"
]].sort_values("Total Net", ascending=False).rename(columns={"Stage_Short":"Stage"})

# PRE-COMPUTE for formula refs
fp_last_row = 2 + len(full_df)   # Full Pipeline last data Excel row (1-based)

# ── WRITE EXCEL ───────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as writer:
    wb = writer.book

    # ── FORMATS ──────────────────────────────────────────────────────────────
    fmt_title   = wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF",
                                  "bg_color":"#1a3a6b","align":"center","valign":"vcenter"})
    fmt_header  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1a3a6b",
                                  "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    fmt_kpi_lbl = wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"})
    fmt_kpi_val = wb.add_format({"bold":True,"bg_color":"#EBF0FB","border":1,
                                  "num_format":"#,##0","align":"right"})
    fmt_num     = wb.add_format({"num_format":"#,##0","border":1,"align":"right"})
    fmt_text    = wb.add_format({"border":1,"align":"left"})
    fmt_date    = wb.add_format({"num_format":"dd-mmm-yyyy","border":1,"align":"center"})
    fmt_alt     = wb.add_format({"bg_color":"#F2F5FB","border":1,"align":"left"})
    fmt_alt_num = wb.add_format({"bg_color":"#F2F5FB","num_format":"#,##0","border":1,"align":"right"})
    fmt_red     = wb.add_format({"bg_color":"#FFE0E0","border":1,"align":"left"})
    fmt_red_num = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right"})
    fmt_red_dt  = wb.add_format({"bg_color":"#FFE0E0","num_format":"dd-mmm-yyyy","border":1,"align":"center"})
    fmt_grn     = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"})
    fmt_grn_num = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"})
    fmt_bu_lbl  = wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"})
    fmt_bu_num  = wb.add_format({"bold":True,"bg_color":"#D9E1F2","num_format":"#,##0","border":1,"align":"right"})
    fmt_bu_hdr  = wb.add_format({"bold":True,"bg_color":"#2E5FA3","font_color":"#FFFFFF",
                                  "border":1,"align":"left","font_size":11})
    fmt_opp     = wb.add_format({"italic":True,"bg_color":"#FAFAFA","border":1,"align":"left","indent":2})
    fmt_opp_num = wb.add_format({"italic":True,"bg_color":"#FAFAFA","num_format":"#,##0","border":1,"align":"right"})
    fmt_month   = wb.add_format({"num_format":"#,##0","border":1,"align":"right","bg_color":"#EFF7F0"})
    fmt_month_hdr = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E7D32",
                                    "border":1,"align":"center","text_wrap":True})
    fmt_month_tot = wb.add_format({"bold":True,"bg_color":"#1B5E20","font_color":"#FFFFFF",
                                    "num_format":"#,##0","border":1,"align":"right"})
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

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 1 — SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.add_worksheet("Summary")
    ws.set_zoom(90)
    ws.set_tab_color("#1a3a6b")
    ws.merge_range("A1:C1", f"AM Pipeline Summary — {TODAY.strftime('%d %B %Y')}", fmt_title)
    ws.set_row(0, 28)

    ws.write(2, 0, "Metric", fmt_header)
    ws.write(2, 1, "Value", fmt_header)
    ws.set_column(0, 0, 34); ws.set_column(1, 1, 22)

    _fp = "'Full Pipeline'"
    _n  = fp_last_row
    kpi_formulas = [
        ("Total Opportunities",        f"=COUNTA({_fp}!C3:C{_n})",                                        fmt_kpi_val),
        ("Total Gross Pipeline (QAR)", f"=SUM({_fp}!I3:I{_n})",                                           fmt_kpi_val),
        ("Total Net Pipeline (QAR)",   f"=SUM({_fp}!J3:J{_n})",                                           fmt_kpi_val),
        ("Forecasted Gross (QAR)",     f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!I3:I{_n})',                   fmt_kpi_val),
        ("Forecasted Net (QAR)",       f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!J3:J{_n})',                   fmt_kpi_val),
        ("Strategic Opportunities",    f'=COUNTIF({_fp}!M3:M{_n},"Yes")',                                 fmt_kpi_val),
        ("Overdue Deals",              f'=COUNTIF({_fp}!Q3:Q{_n},"YES")',                                  fmt_kpi_val),
    ]
    for i, (label, formula, fmt) in enumerate(kpi_formulas):
        ws.write(3+i, 0, label, fmt_kpi_lbl)
        ws.write_formula(3+i, 1, formula, fmt)

    # Stage table (right side)
    ws.merge_range("D1:H1", "Pipeline by Stage", fmt_title)
    stage_cols = ["Stage","Count","Gross (QAR)","Net (QAR)"]
    write_header_row(ws, 2, stage_cols)
    ws.set_column(3, 3, 22); ws.set_column(4, 4, 8)
    ws.set_column(5, 5, 18); ws.set_column(6, 6, 18)
    for r, row in stage_df.reset_index(drop=True).iterrows():
        xl1 = 4 + r
        alt = (r % 2 == 1)
        ws.write(3+r, 3, row["Stage"], fmt_alt if alt else fmt_text)
        ws.write_formula(3+r, 4, f'=COUNTIF({_fp}!D3:D{_n},D{xl1})',             fmt_alt_num if alt else fmt_num)
        ws.write_formula(3+r, 5, f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!I3:I{_n})',fmt_alt_num if alt else fmt_num)
        ws.write_formula(3+r, 6, f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!J3:J{_n})',fmt_alt_num if alt else fmt_num)
    _sg = len(stage_df)
    ws.write(3+_sg, 3, "TOTAL", tot_lbl)
    ws.write_formula(3+_sg, 4, f"=SUM(E4:E{3+_sg})", tot_fmt)
    ws.write_formula(3+_sg, 5, f"=SUM(F4:F{3+_sg})", tot_fmt)
    ws.write_formula(3+_sg, 6, f"=SUM(G4:G{3+_sg})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 — BY ACCOUNT MANAGER
    # ════════════════════════════════════════════════════════════════════════
    am_ws = wb.add_worksheet("By Account Manager")
    am_ws.set_zoom(90)
    am_ws.set_tab_color("#FF8C00")
    am_ws.merge_range("A1:F1", f"Pipeline by Account Manager (Capability Sales) — {TODAY.strftime('%d %B %Y')}", fmt_title)
    am_ws.set_row(0, 28)

    am_cols   = ["Account Manager","Deals","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    am_widths = [34, 8, 20, 20, 22]
    write_header_row(am_ws, 1, am_cols, am_widths)

    _am_data_start = 3   # 1-based first data row
    for r, row in am_agg.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        am_ws.write(2+r, 0, row["Account Manager"], fmt_alt if alt else fmt_text)
        am_ws.write_number(2+r, 1, row["Count"],         fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+r, 2, row["Gross"],         fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+r, 3, row["Net"],           fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+r, 4, row["Forecasted Net"],fmt_alt_num if alt else fmt_num)
    _am_rN = 2 + len(am_agg)
    am_ws.write(_am_rN, 0, "TOTAL", tot_lbl)
    am_ws.write_formula(_am_rN, 1, f"=SUM(B{_am_data_start}:B{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 2, f"=SUM(C{_am_data_start}:C{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 3, f"=SUM(D{_am_data_start}:D{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 4, f"=SUM(E{_am_data_start}:E{_am_rN})", tot_fmt)

    # Per-AM deal detail below (grouped by AM)
    det_off = _am_rN + 2
    am_ws.merge_range(det_off, 0, det_off, 5, "Deal Detail by Account Manager", fmt_bu_hdr)
    am_ws.set_row(det_off, 22)
    det_cols   = ["Account Manager","Account Name","Opportunity","Stage","Quarter","Gross (QAR)","Net (QAR)"]
    det_widths = [30, 28, 40, 22, 12, 18, 18]
    write_header_row(am_ws, det_off+1, det_cols, det_widths)

    # Pre-compute deal detail layout for formula-based subtotals
    _det_layout = []
    _det_pos = det_off + 2
    for am_name, am_grp in am_exp.groupby("AM_exp", sort=False):
        bu_xl = _det_pos; _det_pos += 1
        deal_rows_xl = []
        for _ in am_grp.itertuples():
            deal_rows_xl.append(_det_pos); _det_pos += 1
        _det_layout.append((am_name, bu_xl, am_grp, deal_rows_xl))

    _det_bu_rows = []
    for am_name, bu_xl, am_grp, deal_rows_xl in _det_layout:
        _dF = "+".join("F" + str(r+1) for r in deal_rows_xl)
        _dG = "+".join("G" + str(r+1) for r in deal_rows_xl)
        am_ws.write(bu_xl, 0, am_name, fmt_bu_lbl)
        for cc in range(1, 7):
            am_ws.write(bu_xl, cc, "", fmt_bu_lbl)
        am_ws.write_formula(bu_xl, 5, "=" + _dF, fmt_bu_num)
        am_ws.write_formula(bu_xl, 6, "=" + _dG, fmt_bu_num)
        _det_bu_rows.append(bu_xl + 1)

        for r_pos, (_, row) in zip(deal_rows_xl, am_grp.iterrows()):
            alt = (r_pos % 2 == 1)
            am_ws.write(r_pos, 0, am_name,                fmt_alt if alt else fmt_text)
            am_ws.write(r_pos, 1, str(row["Account Name"]) if pd.notna(row["Account Name"]) else "", fmt_alt if alt else fmt_text)
            am_ws.write(r_pos, 2, str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "", fmt_alt if alt else fmt_text)
            am_ws.write(r_pos, 3, str(row["Stage_Short"]) if pd.notna(row["Stage_Short"]) else "", fmt_alt if alt else fmt_text)
            am_ws.write(r_pos, 4, str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "", fmt_alt if alt else fmt_text)
            am_ws.write_number(r_pos, 5, row["Total Gross"], fmt_alt_num if alt else fmt_num)
            am_ws.write_number(r_pos, 6, row["Total Net"],   fmt_alt_num if alt else fmt_num)

    # Grand Total detail
    _gt_F = "+".join("F" + str(r) for r in _det_bu_rows)
    _gt_G = "+".join("G" + str(r) for r in _det_bu_rows)
    _gt_row = _det_pos
    am_ws.write(_gt_row, 0, "GRAND TOTAL", tot_lbl)
    for cc in range(1, 5):
        am_ws.write(_gt_row, cc, "", tot_lbl)
    am_ws.write_formula(_gt_row, 5, "=" + _gt_F, tot_fmt)
    am_ws.write_formula(_gt_row, 6, "=" + _gt_G, tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — MONTHLY PIPELINE
    # ════════════════════════════════════════════════════════════════════════
    mw = wb.add_worksheet("Monthly Pipeline")
    mw.set_zoom(90)
    mw.set_tab_color("#228B22")
    mw.merge_range(0, 0, 0, 13, f"Monthly Pipeline Breakdown — {TODAY.strftime('%d %B %Y')}", fmt_title)
    mw.set_row(0, 28)

    # Header
    mw.write(1, 0, "Opportunity / Account", fmt_header)
    mw.set_column(0, 0, 40)
    mw.write(1, 1, "Account Name", fmt_header)
    mw.set_column(1, 1, 28)
    for ci, m in enumerate(MONTHS):
        mw.write(1, 2+ci, m, fmt_month_hdr)
        mw.set_column(2+ci, 2+ci, 14)
    mw.write(1, 14, "Total Net", fmt_month_hdr)
    mw.set_column(14, 14, 16)
    mw.freeze_panes(2, 0)

    _month_data_start = 3   # 1-based
    for r, row in df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        mw.write(2+r, 0, str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "", fmt_alt if alt else fmt_text)
        mw.write(2+r, 1, str(row["Account Name"])  if pd.notna(row["Account Name"])  else "", fmt_alt if alt else fmt_text)
        for ci, m in enumerate(MONTHS):
            v = row[m]
            if v and v != 0:
                mw.write_number(2+r, 2+ci, v, fmt_month if alt else fmt_num)
            else:
                mw.write_blank(2+r, 2+ci, None, fmt_alt if alt else fmt_text)
        mw.write_number(2+r, 14, row["Total Net"], fmt_alt_num if alt else fmt_num)

    # TOTAL row
    _m_rN = 2 + len(df)
    mw.write(_m_rN, 0, "TOTAL", tot_lbl)
    mw.write(_m_rN, 1, "", tot_lbl)
    for ci, m in enumerate(MONTHS):
        col_letter = chr(67 + ci)   # C=Jan, D=Feb, ...
        mw.write_formula(_m_rN, 2+ci, f"=SUM({col_letter}{_month_data_start}:{col_letter}{_m_rN})", fmt_month_tot)
    mw.write_formula(_m_rN, 14, f"=SUM(O{_month_data_start}:O{_m_rN})", fmt_month_tot)

    # Monthly summary row (sum per month across all deals)
    mw.merge_range(_m_rN+2, 0, _m_rN+2, 14, "Monthly Totals Summary", fmt_bu_hdr)
    mw.set_row(_m_rN+2, 22)
    write_header_row(mw, _m_rN+3, ["Month"] + MONTHS + ["Total"], [14]+[14]*12+[16])
    mw.write(_m_rN+4, 0, "Net (QAR)", fmt_bu_lbl)
    for ci, m in enumerate(MONTHS):
        col_letter = chr(67 + ci)
        mw.write_formula(_m_rN+4, 1+ci, f"={col_letter}{_m_rN+1}", fmt_bu_num)
    mw.write_formula(_m_rN+4, 13, f"=O{_m_rN+1}", fmt_bu_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 — DU BREAKDOWN
    # ════════════════════════════════════════════════════════════════════════
    du_ws = wb.add_worksheet("DU Breakdown")
    du_ws.set_zoom(90)
    du_ws.set_tab_color("#FF8C00")
    du_ws.merge_range("A1:F1", "Gross & Net Breakdown by BU / Delivery Unit", fmt_title)
    du_ws.set_row(0, 28)

    du_cols   = ["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    du_widths = [42, 52, 20, 20, 22]
    write_header_row(du_ws, 1, du_cols, du_widths)

    # Pre-compute layout
    _du_layout = []
    _du_pos = 0
    for bu_name, bu_grp in du_totals.groupby("BU"):
        bu_r0 = 2 + _du_pos; _du_pos += 1
        du_list = []
        for _, drow in bu_grp.iterrows():
            du_r0 = 2 + _du_pos; _du_pos += 1
            du_deals = du_exp_df[du_exp_df["DU"] == drow["DU"]].copy()
            opp_rows = []
            for _ in du_deals.itertuples():
                opp_rows.append(2 + _du_pos); _du_pos += 1
            du_list.append((drow, du_r0, opp_rows))
        _du_layout.append((bu_name, bu_r0, bu_grp, du_list))
    _du_grand_r = 2 + _du_pos

    _du_bu_rows = []
    for bu_name, bu_r0, bu_grp, du_list in _du_layout:
        _du_r1s = [dr0+1 for (_, dr0, _) in du_list]
        _bu_C = "+".join("C" + str(r) for r in _du_r1s)
        _bu_D = "+".join("D" + str(r) for r in _du_r1s)
        _bu_E = "+".join("E" + str(r) for r in _du_r1s)
        du_ws.write(bu_r0, 0, bu_name, fmt_bu_lbl)
        du_ws.write(bu_r0, 1, "", fmt_bu_lbl)
        du_ws.write_formula(bu_r0, 2, "=" + _bu_C, fmt_bu_num)
        du_ws.write_formula(bu_r0, 3, "=" + _bu_D, fmt_bu_num)
        du_ws.write_formula(bu_r0, 4, "=" + _bu_E, fmt_bu_num)
        _du_bu_rows.append(bu_r0 + 1)

        for drow, du_r0, opp_rows in du_list:
            alt = (du_r0 % 2 == 1)
            if opp_rows:
                _r1 = min(opp_rows)+1; _rN = max(opp_rows)+1
                _os = lambda c: f"=SUM({c}{_r1}:{c}{_rN})"
            else:
                _os = lambda c: "=0"
            du_ws.write(du_r0, 0, "", fmt_alt if alt else fmt_text)
            du_ws.write(du_r0, 1, drow["DU"], fmt_alt if alt else fmt_text)
            du_ws.write_formula(du_r0, 2, _os("C"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 3, _os("D"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 4, _os("E"), fmt_alt_num if alt else fmt_num)

            du_deals = du_exp_df[du_exp_df["DU"] == drow["DU"]].copy()
            for opp_r0, (_, deal) in zip(opp_rows, du_deals.iterrows()):
                opp_label = f"  ↳  {deal['Lead/Opp Name']}"
                fore_net = deal["Net"] if str(deal.get("Forecasted","")).strip() == "Yes" else 0
                du_ws.write(opp_r0, 0, "", fmt_opp)
                du_ws.write(opp_r0, 1, opp_label, fmt_opp)
                du_ws.write_number(opp_r0, 2, deal["Gross"], fmt_opp_num)
                du_ws.write_number(opp_r0, 3, deal["Net"],   fmt_opp_num)
                du_ws.write_number(opp_r0, 4, fore_net,      fmt_opp_num)

    _gt_C = "+".join("C" + str(r) for r in _du_bu_rows)
    _gt_D = "+".join("D" + str(r) for r in _du_bu_rows)
    _gt_E = "+".join("E" + str(r) for r in _du_bu_rows)
    t = _du_grand_r
    du_ws.write(t, 0, "GRAND TOTAL", tot_lbl); du_ws.write(t, 1, "", tot_lbl)
    du_ws.write_formula(t, 2, "=" + _gt_C, tot_fmt)
    du_ws.write_formula(t, 3, "=" + _gt_D, tot_fmt)
    du_ws.write_formula(t, 4, "=" + _gt_E, tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 — QUARTERLY PLAN
    # ════════════════════════════════════════════════════════════════════════
    qw = wb.add_worksheet("Quarterly Plan")
    qw.set_zoom(90)
    qw.set_tab_color("#DAA520")
    qw.merge_range("A1:E1", f"Quarterly Close Plan — {TODAY.strftime('%d %B %Y')}", fmt_title)
    qw.set_row(0, 28)

    q_cols   = ["Quarter","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"]
    write_header_row(qw, 1, q_cols, [14, 8, 20, 20, 22])
    for r, row in q_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        qw.write(2+r, 0, row["Closure Due Quarter"], fmt_alt if alt else fmt_text)
        qw.write_number(2+r, 1, row["Count"],          fmt_alt_num if alt else fmt_num)
        qw.write_number(2+r, 2, row["Gross"],          fmt_alt_num if alt else fmt_num)
        qw.write_number(2+r, 3, row["Net"],            fmt_alt_num if alt else fmt_num)
        qw.write_number(2+r, 4, row["Forecasted Net"], fmt_alt_num if alt else fmt_num)
    _q_r1 = 3; _q_rN = 2 + len(q_df)
    qw.write(_q_rN, 0, "TOTAL", tot_lbl)
    qw.write_formula(_q_rN, 1, f"=SUM(B{_q_r1}:B{_q_rN})", tot_fmt)
    qw.write_formula(_q_rN, 2, f"=SUM(C{_q_r1}:C{_q_rN})", tot_fmt)
    qw.write_formula(_q_rN, 3, f"=SUM(D{_q_r1}:D{_q_rN})", tot_fmt)
    qw.write_formula(_q_rN, 4, f"=SUM(E{_q_r1}:E{_q_rN})", tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 — AM BREAKDOWN (styled, one row per AM per opportunity)
    # ════════════════════════════════════════════════════════════════════════
    CLR = {
        "title_bg":    "1F3864", "hdr_deal": "1F3864", "hdr_am":  "17375E",
        "hdr_finance": "1F4E79", "hdr_other":"2E5FA3",
        "am_fill":     "EDF2F9", "num_fill": "EBF5FB",
        "tot_fill":    "D5E8F5", "date_fill":"FFF2CC",
        "alt_a":       "F5F8FF", "alt_b":    "FFFFFF",
    }
    def _bd_fmt(bg, top, num_fmt=None):
        d = {"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
        if num_fmt:
            d["num_format"] = num_fmt; d["align"] = "right"
        return wb.add_format(d)

    fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_am      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_am"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    fmt_title_bd = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+CLR["title_bg"],"align":"center","valign":"vcenter"})

    ft_a1 = _bd_fmt(CLR["alt_a"], 2); ft_an = _bd_fmt(CLR["alt_a"], 1)
    ft_b1 = _bd_fmt(CLR["alt_b"], 2); ft_bn = _bd_fmt(CLR["alt_b"], 1)
    fn_a1 = _bd_fmt(CLR["alt_a"], 2, "#,##0"); fn_an = _bd_fmt(CLR["alt_a"], 1, "#,##0")
    fn_b1 = _bd_fmt(CLR["alt_b"], 2, "#,##0"); fn_bn = _bd_fmt(CLR["alt_b"], 1, "#,##0")
    fam_f = _bd_fmt(CLR["am_fill"], 2); fam_n = _bd_fmt(CLR["am_fill"], 1)
    fxn_f = _bd_fmt(CLR["num_fill"], 2, "#,##0"); fxn_n = _bd_fmt(CLR["num_fill"], 1, "#,##0")
    bd_ftot    = wb.add_format({"bg_color":"#"+CLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
    bd_ftotbl  = wb.add_format({"bg_color":"#"+CLR["num_fill"],"border":1,"font_size":9})
    bd_fdate   = wb.add_format({"bg_color":"#"+CLR["date_fill"],"num_format":"dd-mmm-yyyy","border":1,"align":"center","font_size":9})
    bd_fdatebl = wb.add_format({"bg_color":"#"+CLR["alt_a"],"border":1,"font_size":9})

    bw = wb.add_worksheet("AM Breakdown")
    bw.set_zoom(85); bw.set_tab_color("#9370DB")

    bd_ocols = [
        ("SNo.",              6,  "deal"),
        ("Account Name",     24,  "deal"),
        ("Lead/Opp Name",    36,  "deal"),
        ("Account Manager",  30,  "am"),
        ("Stage",            22,  "other"),
        ("Account Manager",  24,  "other"),
        ("Sector",           16,  "other"),
        ("Quarter",          10,  "other"),
        ("Win Prob",         12,  "other"),
        ("Forecasted",       10,  "other"),
        ("Strategic",        10,  "other"),
        ("Gross (breakdown)",16,  "finance"),
        ("Net (breakdown)",  16,  "finance"),
        ("Total Gross",      15,  "finance"),
        ("Total Net",        15,  "finance"),
        ("Est. Close Date",  14,  "other"),
    ]
    bd_hfmt = {"deal":fh_deal,"am":fh_am,"finance":fh_finance,"other":fh_other}
    bd_nc   = len(bd_ocols)
    bw.merge_range(0, 0, 0, bd_nc-1, "AM Pipeline — Expanded by Account Manager", fmt_title_bd)
    bw.set_row(0, 28)
    for c, (cn, cw, ct) in enumerate(bd_ocols):
        bw.write(1, c, cn, bd_hfmt[ct])
        bw.set_column(c, c, cw)
    bw.set_row(1, 28); bw.freeze_panes(2, 0)

    bd_cmap = {name: idx for idx, (name, _, __) in enumerate(bd_ocols)}
    g_col = 11   # Gross (breakdown) col index
    n_col = 12   # Net (breakdown) col index

    # Pre-compute deal row ranges for SUM formulas
    _bd_deal_rows = {}
    for rp, (_, row) in enumerate(am_exp.iterrows()):
        didx = row["_deal_idx"]; xl_r = 2 + rp
        if didx not in _bd_deal_rows:
            _bd_deal_rows[didx] = [xl_r, xl_r]
        else:
            _bd_deal_rows[didx][1] = xl_r

    bw.autofilter(1, 0, 1+len(am_exp), bd_nc-1)
    prev_deal = None; alt_tog = False
    for rp, (_, row) in enumerate(am_exp.iterrows()):
        didx = row["_deal_idx"]; isf = bool(row["_is_first"])
        if didx != prev_deal:
            alt_tog = not alt_tog; prev_deal = didx
        ft  = (ft_a1 if isf else ft_an) if alt_tog else (ft_b1 if isf else ft_bn)
        fn  = (fn_a1 if isf else fn_an) if alt_tog else (fn_b1 if isf else fn_bn)
        fam = fam_f if isf else fam_n
        fxn = fxn_f if isf else fxn_n
        xl_r = 2 + rp

        def _bws(ci, val, fmt):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                bw.write_blank(xl_r, ci, None, fmt)
            else:
                bw.write(xl_r, ci, str(val), fmt)

        def _bwn(ci, val, fmt):
            v = _parse_num(val) if not isinstance(val, (int, float)) else val
            if v is None or (isinstance(v, float) and pd.isna(v)):
                bw.write_blank(xl_r, ci, None, fmt)
            else:
                bw.write_number(xl_r, ci, v, fmt)

        _bws(0,  row.get("SNo.")               if isf else None, ft)
        _bws(1,  row.get("Account Name"),       ft)
        _bws(2,  row.get("Lead/Opp Name"),      ft)
        _bws(3,  row.get("AM_exp"),             fam)   # AM expanded
        _bws(4,  row.get("Stage_Short"),        ft)
        _bws(5,  row.get("Capability Sales"),   ft)    # all AMs on deal
        _bws(6,  row.get("Sector"),             ft)
        _bws(7,  row.get("Closure Due Quarter"),ft)
        _bws(8,  row.get("Winning Probability"),ft)
        _bws(9,  row.get("Forecasted"),         ft)
        _bws(10, row.get("Strategic Opportunity"), ft)
        _bwn(11, _parse_num(str(row.get("Gross (Breakdown)","")).split("\n")[0] if isf else None), fxn)
        _bwn(12, _parse_num(str(row.get("Net (breakdown)","")).split("\n")[0]   if isf else None), fxn)
        if isf:
            r0, r1 = _bd_deal_rows[didx]
            gc = chr(65 + g_col); nc = chr(65 + n_col)
            bw.write_formula(xl_r, 13, f"=SUM({gc}{r0+1}:{gc}{r1+1})", bd_ftot)
            bw.write_formula(xl_r, 14, f"=SUM({nc}{r0+1}:{nc}{r1+1})", bd_ftot)
        else:
            bw.write_blank(xl_r, 13, None, bd_ftotbl)
            bw.write_blank(xl_r, 14, None, bd_ftotbl)
        cd_v = row.get("Est. Close Date")
        if isf and pd.notna(cd_v):
            try: bw.write_datetime(xl_r, 15, pd.Timestamp(cd_v).to_pydatetime(), bd_fdate)
            except Exception: bw.write(xl_r, 15, str(cd_v), bd_fdate)
        else:
            bw.write_blank(xl_r, 15, None, bd_fdate if isf else bd_fdatebl)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 7 — FULL PIPELINE
    # ════════════════════════════════════════════════════════════════════════
    pw = wb.add_worksheet("Full Pipeline")
    pw.set_zoom(85); pw.set_tab_color("#6495ED")
    pw.merge_range("A1:Q1", "Full AM Pipeline — All Opportunities", fmt_title)
    pw.set_row(0, 28); pw.freeze_panes(2, 0)

    full_cols   = ["#","Account Name","Opportunity","Stage","Capability Sales","Sector",
                   "BU","DU","Gross (QAR)","Net (QAR)","Win Prob","Forecasted",
                   "Strategic","Quarter","Est. Close Date","Source","Overdue"]
    full_widths = [5, 28, 36, 22, 28, 16, 30, 36, 18, 18, 12, 12, 10, 10, 16, 16, 8]
    write_header_row(pw, 1, full_cols, full_widths)
    pw.autofilter(1, 0, 1+len(full_df), len(full_cols)-1)

    for r, row in full_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        is_overdue = bool(row["Overdue"])
        t_fmt = fmt_red     if is_overdue else (fmt_alt     if alt else fmt_text)
        n_fmt = fmt_red_num if is_overdue else (fmt_alt_num if alt else fmt_num)
        d_fmt = fmt_red_dt  if is_overdue else fmt_date

        pw.write(2+r, 0,  str(row["SNo."]) if pd.notna(row["SNo."]) else "", t_fmt)
        pw.write(2+r, 1,  str(row["Account Name"])   if pd.notna(row["Account Name"])   else "", t_fmt)
        pw.write(2+r, 2,  str(row["Lead/Opp Name"])  if pd.notna(row["Lead/Opp Name"])  else "", t_fmt)
        pw.write(2+r, 3,  str(row["Stage"])           if pd.notna(row["Stage"])           else "", t_fmt)
        pw.write(2+r, 4,  str(row["Capability Sales"])if pd.notna(row["Capability Sales"])else "", t_fmt)
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

    # TOTAL row
    _fp_r1 = 3; _fp_rN = 2 + len(full_df)
    pw.write(_fp_rN, 0, "TOTAL", tot_lbl)
    pw.write_formula(_fp_rN, 8, f"=SUM(I{_fp_r1}:I{_fp_rN})", tot_fmt)
    pw.write_formula(_fp_rN, 9, f"=SUM(J{_fp_r1}:J{_fp_rN})", tot_fmt)

print(f"Done! Saved: {OUT_FILE}")
