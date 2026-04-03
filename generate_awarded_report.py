"""
Awarded Deals Report Generator
Combines 2025 and 2026 awarded deals into a formatted Excel workbook.
Usage: python generate_awarded_report.py
Output: C:/Users/khali/Downloads/Awarded_Report_YYYY-MM-DD.xlsx
"""

import pandas as pd
import re
import warnings
from datetime import date
warnings.filterwarnings("ignore")

INPUT_2026 = r"C:\Users\khali\Downloads\data (1).xlsx"
INPUT_2025 = r"C:\Users\khali\Downloads\data (4).xlsx"   # update path if file was renamed
COA_FILE   = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
TODAY      = date.today()
OUT_FILE   = rf"C:\Users\khali\Downloads\Awarded_Report_{TODAY}.xlsx"

# ── CHARTER OF ACCOUNTS — DU → BU MAPPING ────────────────────────────────────
coa = pd.read_excel(COA_FILE)
coa.columns = coa.columns.str.strip()
coa["_code"] = coa["DU"].str.extract(r"(\d{6})")
DU_TO_BU = coa.dropna(subset=["_code"]).set_index("_code")["BU"].to_dict()

def du_to_bu(du_str):
    m = re.match(r"(\d{6})", str(du_str).strip())
    return DU_TO_BU.get(m.group(1), "Unknown") if m else "Unknown"

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

# ── LOAD & CLEAN ──────────────────────────────────────────────────────────────
def load_awarded(path, year_label):
    df = pd.read_excel(path, sheet_name="Export")
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    df["Year"]           = year_label
    df["Total Gross"]    = pd.to_numeric(df["Total Gross"],   errors="coerce").fillna(0)
    df["Total Net"]      = pd.to_numeric(df["Total Net"],     errors="coerce").fillna(0)
    df["Project Value"]  = pd.to_numeric(df["Project value (as per the contract value)"], errors="coerce").fillna(0)
    df["Client Commitment"] = pd.to_numeric(df["Client Commitment/WOs Net"], errors="coerce").fillna(0)
    def simplify_nr(val):
        if pd.isna(val): return "Unknown"
        vals = set(v.strip() for v in str(val).split("\n"))
        if vals == {"New"}:   return "New"
        if vals == {"Renew"}: return "Renew"
        return "Mixed"
    df["Type"] = df["New/Renew"].apply(simplify_nr)
    return df

import os
frames = []
if os.path.exists(INPUT_2025):
    frames.append(load_awarded(INPUT_2025, "2025"))
else:
    print(f"WARNING: 2025 file not found, skipping: {INPUT_2025}")
if os.path.exists(INPUT_2026):
    frames.append(load_awarded(INPUT_2026, "2026"))
else:
    print(f"WARNING: 2026 file not found, skipping: {INPUT_2026}")
if not frames:
    raise FileNotFoundError("No awarded deal files found. Check INPUT_2025 and INPUT_2026 paths.")
df = pd.concat(frames, ignore_index=True)

# ── DU EXPLOSION ──────────────────────────────────────────────────────────────
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
            "Year":            row["Year"],
            "BU":              du_to_bu(du),
            "DU":              du,
            "Gross":           g_val,
            "Net":             n_val,
            "Account Manager": row.get("Account Manager", ""),
            "Stage":           row.get("Stage", ""),
            "Award Quarter":   row.get("Award Quarter", ""),
            "Contracted":      str(row.get("Contracted", "")).strip(),
            "Account Name":    row.get("Account Name", ""),
            "Opportunity":     row.get("Opportunity Name", ""),
        })
du_exp = pd.DataFrame(du_rows)

# ── BUILD SUMMARY TABLES ───────────────────────────────────────────────────────

# KPIs
total_gross   = df["Total Gross"].sum()
total_net     = df["Total Net"].sum()
total_pv      = df["Project Value"].sum()
contracted    = len(df[df["Contracted"] == "Yes"])
loa_count     = len(df[df["Stage"].str.contains("Letter Of Award", na=False)])
signed_count  = len(df[df["Stage"].str.contains("Contracting", na=False)])
new_count     = len(df[df["Type"] == "New"])
renew_count   = len(df[df["Type"] == "Renew"])
count_2025    = len(df[df["Year"] == "2025"])
count_2026    = len(df[df["Year"] == "2026"])

kpi_df = pd.DataFrame([
    {"Metric": "Total Awarded Deals",           "Value": len(df)},
    {"Metric": "  — 2025",                      "Value": count_2025},
    {"Metric": "  — 2026",                      "Value": count_2026},
    {"Metric": "Total Gross (QAR)",             "Value": total_gross},
    {"Metric": "Total Net (QAR)",               "Value": total_net},
    {"Metric": "Total Contract Value (QAR)",    "Value": total_pv},
    {"Metric": "Contracted (Signed)",           "Value": signed_count},
    {"Metric": "LOA (Not Yet Signed)",          "Value": loa_count},
    {"Metric": "New Deals",                     "Value": new_count},
    {"Metric": "Renew Deals",                   "Value": renew_count},
])

# By Year summary
year_df = (
    df.groupby("Year")
    .agg(Count=("Opportunity Name","count"), Gross=("Total Gross","sum"),
         Net=("Total Net","sum"), PV=("Project Value","sum"))
    .reset_index().sort_values("Year")
)

# By Stage
stage_df = (
    df.groupby(["Year","Stage"])
    .agg(Count=("Opportunity Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().sort_values(["Year","Stage"])
)

# By Account Manager
am_df = (
    df.groupby("Account Manager")
    .agg(Count=("Opportunity Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().sort_values("Net", ascending=False)
)
am_25 = df[df["Year"]=="2025"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Net 2025"})
am_26 = df[df["Year"]=="2026"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Net 2026"})
am_df = am_df.merge(am_25, on="Account Manager", how="left").fillna({"Net 2025":0})
am_df = am_df.merge(am_26, on="Account Manager", how="left").fillna({"Net 2026":0})

# By Award Quarter
q_df = (
    df.groupby(["Year","Award Quarter"])
    .agg(Count=("Opportunity Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().sort_values(["Year","Award Quarter"])
)

# New vs Renew
nr_df = (
    df.groupby(["Year","Type"])
    .agg(Count=("Opportunity Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
    .reset_index().sort_values(["Year","Type"])
)

# DU Breakdown
du_totals = (
    du_exp.groupby(["BU","DU"])[["Gross","Net"]]
    .sum().reset_index().sort_values(["BU","Net"], ascending=[True,False])
)
du_by_year = (
    du_exp.groupby(["BU","DU","Year"])[["Gross","Net"]]
    .sum().reset_index()
)

# DU × Year
du_year_pivot = (
    du_by_year.pivot_table(index=["BU","DU"], columns="Year", values="Net", aggfunc="sum", fill_value=0)
    .reset_index()
)
du_totals = du_totals.merge(du_year_pivot, on=["BU","DU"], how="left")
for yr in ["2025","2026"]:
    if yr not in du_totals.columns:
        du_totals[yr] = 0

# Full table
full_df = df[[
    "Year","SNo.","Account Name","Opportunity Name","Stage","Account Manager",
    "Type","Total Gross","Total Net","Project Value",
    "Award Quarter","Contracted","Contract Signed Quarter","ORF Number","Project Duration"
]].sort_values(["Year","Total Net"], ascending=[True,False])

# ── WRITE EXCEL ───────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as writer:
    wb = writer.book

    # ── FORMATS ──────────────────────────────────────────────────────────────
    fmt_title   = wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF",
                                  "bg_color":"#1a3a6b","align":"center","valign":"vcenter"})
    fmt_header  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1a3a6b",
                                  "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    fmt_hdr_25  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E5FA3",
                                  "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    fmt_hdr_26  = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1a6b3a",
                                  "border":1,"align":"center","valign":"vcenter","text_wrap":True})
    fmt_kpi_lbl = wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"})
    fmt_kpi_val = wb.add_format({"bold":True,"bg_color":"#EBF0FB","border":1,
                                  "num_format":"#,##0","align":"right"})
    fmt_num     = wb.add_format({"num_format":"#,##0","border":1,"align":"right"})
    fmt_text    = wb.add_format({"border":1,"align":"left"})
    fmt_date    = wb.add_format({"num_format":"dd-mmm-yyyy","border":1,"align":"center"})
    fmt_alt     = wb.add_format({"bg_color":"#F2F5FB","border":1,"align":"left"})
    fmt_alt_num = wb.add_format({"bg_color":"#F2F5FB","num_format":"#,##0","border":1,"align":"right"})
    fmt_grn     = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"})
    fmt_grn_num = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"})
    fmt_opp     = wb.add_format({"italic":True,"bg_color":"#FAFAFA","border":1,"align":"left","indent":2})
    fmt_opp_num = wb.add_format({"italic":True,"bg_color":"#FAFAFA","num_format":"#,##0","border":1,"align":"right"})
    fmt_bu_hdr  = wb.add_format({"bold":True,"bg_color":"#2E5FA3","font_color":"#FFFFFF",
                                  "border":1,"align":"left","font_size":11})
    fmt_bu_lbl  = wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"})
    fmt_bu_num  = wb.add_format({"bold":True,"bg_color":"#D9E1F2","num_format":"#,##0","border":1,"align":"right"})
    fmt_tot_lbl = wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF",
                                  "border":1,"align":"left"})
    fmt_tot_num = wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF",
                                  "num_format":"#,##0","border":1,"align":"right"})
    fmt_y25     = wb.add_format({"bg_color":"#DAE8FC","border":1,"align":"left"})
    fmt_y25_num = wb.add_format({"bg_color":"#DAE8FC","num_format":"#,##0","border":1,"align":"right"})
    fmt_y26     = wb.add_format({"bg_color":"#D5E8D4","border":1,"align":"left"})
    fmt_y26_num = wb.add_format({"bg_color":"#D5E8D4","num_format":"#,##0","border":1,"align":"right"})

    def write_header_row(ws, row, cols, widths=None):
        for c, col in enumerate(cols):
            ws.write(row, c, col, fmt_header)
        if widths:
            for c, w in enumerate(widths):
                ws.set_column(c, c, w)

    def write_data_rows(ws, start_row, df_in, col_types):
        for r, row_data in df_in.iterrows():
            alt   = (r % 2 == 1)
            yr    = str(row_data.get("Year","")) if "Year" in df_in.columns else ""
            is_25 = yr == "2025"
            is_26 = yr == "2026"
            for c, (col, typ) in enumerate(col_types):
                val = row_data[col]
                if typ == "num":
                    if is_25: f = fmt_y25_num
                    elif is_26: f = fmt_y26_num
                    else: f = fmt_alt_num if alt else fmt_num
                    ws.write_number(start_row+r, c, val if pd.notna(val) else 0, f)
                elif typ == "date":
                    f = fmt_date
                    if pd.notna(val):
                        ws.write_datetime(start_row+r, c, val.to_pydatetime(), f)
                    else:
                        ws.write_blank(start_row+r, c, None, f)
                else:
                    if is_25: f = fmt_y25
                    elif is_26: f = fmt_y26
                    else: f = fmt_alt if alt else fmt_text
                    ws.write(start_row+r, c, str(val) if pd.notna(val) else "", f)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 1 — SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    ws = wb.add_worksheet("Summary")
    ws.set_zoom(90)
    ws.set_tab_color("#1a3a6b")
    ws.merge_range("A1:C1", f"Awarded Deals Summary — {TODAY.strftime('%d %B %Y')}", fmt_title)
    ws.set_row(0, 28)

    ws.write(2, 0, "Metric",  fmt_header)
    ws.write(2, 1, "Value",   fmt_header)
    ws.set_column(0, 0, 36)
    ws.set_column(1, 1, 22)

    for i, row in kpi_df.iterrows():
        ws.write(3+i, 0, row["Metric"], fmt_kpi_lbl)
        ws.write_number(3+i, 1, row["Value"], fmt_kpi_val)

    # Year comparison (offset right)
    ws.merge_range("D1:H1", "Summary by Year", fmt_title)
    yr_cols = ["Year","Deals","Gross (QAR)","Net (QAR)","Contract Value (QAR)"]
    for c, col in enumerate(yr_cols):
        ws.write(2, 3+c, col, fmt_header)
    ws.set_column(3, 3, 8)
    ws.set_column(4, 4, 8)
    ws.set_column(5, 5, 20)
    ws.set_column(6, 6, 20)
    ws.set_column(7, 7, 22)
    for i, row in year_df.reset_index(drop=True).iterrows():
        f_txt = fmt_y25 if row["Year"] == "2025" else fmt_y26
        f_num = fmt_y25_num if row["Year"] == "2025" else fmt_y26_num
        ws.write(3+i, 3, row["Year"],    f_txt)
        ws.write_number(3+i, 4, row["Count"], f_num)
        ws.write_number(3+i, 5, row["Gross"], f_num)
        ws.write_number(3+i, 6, row["Net"],   f_num)
        ws.write_number(3+i, 7, row["PV"],    f_num)
    # Totals row
    tr = 3 + len(year_df)
    ws.write(tr, 3, "TOTAL", fmt_tot_lbl)
    ws.write_number(tr, 4, year_df["Count"].sum(), fmt_tot_num)
    ws.write_number(tr, 5, year_df["Gross"].sum(), fmt_tot_num)
    ws.write_number(tr, 6, year_df["Net"].sum(),   fmt_tot_num)
    ws.write_number(tr, 7, year_df["PV"].sum(),    fmt_tot_num)

    # Stage table
    off_s = len(kpi_df) + 5
    ws.merge_range(off_s, 0, off_s, 3, "By Stage", fmt_title)
    stg_cols = ["Year","Stage","Deals","Net (QAR)"]
    for c, col in enumerate(stg_cols):
        ws.write(off_s+1, c, col, fmt_header)
    ws.set_column(0, 0, 8)
    ws.set_column(1, 1, 36)
    for i, row in stage_df.reset_index(drop=True).iterrows():
        f_txt = fmt_y25 if row["Year"] == "2025" else fmt_y26
        f_num = fmt_y25_num if row["Year"] == "2025" else fmt_y26_num
        ws.write(off_s+2+i, 0, row["Year"],  f_txt)
        ws.write(off_s+2+i, 1, row["Stage"], f_txt)
        ws.write_number(off_s+2+i, 2, row["Count"], f_num)
        ws.write_number(off_s+2+i, 3, row["Net"],   f_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 — DU BREAKDOWN
    # ════════════════════════════════════════════════════════════════════════
    du_ws = wb.add_worksheet("DU Breakdown")
    du_ws.set_zoom(90)
    du_ws.set_tab_color("#FF8C00")
    du_ws.merge_range("A1:G1", "Gross & Net Breakdown by BU / Delivery Unit", fmt_title)
    du_ws.set_row(0, 28)

    du_cols   = ["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Net 2025 (QAR)","Net 2026 (QAR)"]
    du_widths = [42, 52, 20, 20, 20, 20]
    for c, col in enumerate(du_cols):
        du_ws.write(1, c, col, fmt_header)
        du_ws.set_column(c, c, du_widths[c])

    r_out = 0
    for bu_name, bu_grp in du_totals.groupby("BU"):
        # BU subtotal row
        du_ws.write(2+r_out, 0, bu_name, fmt_bu_lbl)
        du_ws.write(2+r_out, 1, "", fmt_bu_lbl)
        du_ws.write_number(2+r_out, 2, bu_grp["Gross"].sum(),         fmt_bu_num)
        du_ws.write_number(2+r_out, 3, bu_grp["Net"].sum(),           fmt_bu_num)
        du_ws.write_number(2+r_out, 4, bu_grp.get("2025",pd.Series([0])).sum(), fmt_bu_num)
        du_ws.write_number(2+r_out, 5, bu_grp.get("2026",pd.Series([0])).sum(), fmt_bu_num)
        r_out += 1
        for _, row in bu_grp.iterrows():
            alt = (r_out % 2 == 1)
            du_ws.write(2+r_out, 0, "", fmt_alt if alt else fmt_text)
            du_ws.write(2+r_out, 1, row["DU"], fmt_alt if alt else fmt_text)
            du_ws.write_number(2+r_out, 2, row["Gross"], fmt_alt_num if alt else fmt_num)
            du_ws.write_number(2+r_out, 3, row["Net"],   fmt_alt_num if alt else fmt_num)
            du_ws.write_number(2+r_out, 4, row.get("2025", 0), fmt_alt_num if alt else fmt_num)
            du_ws.write_number(2+r_out, 5, row.get("2026", 0), fmt_alt_num if alt else fmt_num)
            r_out += 1
            # Opportunity sub-rows for this DU
            du_deals = du_exp[du_exp["DU"] == row["DU"]].copy()
            for _, deal in du_deals.iterrows():
                opp_label = f"  ↳  {deal['Opportunity']}"
                du_ws.write(2+r_out, 0, "", fmt_opp)
                du_ws.write(2+r_out, 1, opp_label, fmt_opp)
                du_ws.write_number(2+r_out, 2, deal["Gross"], fmt_opp_num)
                du_ws.write_number(2+r_out, 3, deal["Net"],   fmt_opp_num)
                net_25 = deal["Net"] if str(deal.get("Year","")) == "2025" else 0
                net_26 = deal["Net"] if str(deal.get("Year","")) == "2026" else 0
                du_ws.write_number(2+r_out, 4, net_25, fmt_opp_num)
                du_ws.write_number(2+r_out, 5, net_26, fmt_opp_num)
                r_out += 1

    # Grand total
    t = 2 + r_out
    du_ws.write(t, 0, "GRAND TOTAL", fmt_tot_lbl)
    du_ws.write(t, 1, "", fmt_tot_lbl)
    du_ws.write_number(t, 2, du_totals["Gross"].sum(), fmt_tot_num)
    du_ws.write_number(t, 3, du_totals["Net"].sum(),   fmt_tot_num)
    du_ws.write_number(t, 4, du_totals["2025"].sum() if "2025" in du_totals.columns else 0, fmt_tot_num)
    du_ws.write_number(t, 5, du_totals["2026"].sum() if "2026" in du_totals.columns else 0, fmt_tot_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — ACCOUNT MANAGER
    # ════════════════════════════════════════════════════════════════════════
    am_ws = wb.add_worksheet("Account Manager")
    am_ws.set_zoom(90)
    am_ws.set_tab_color("#228B22")
    am_ws.merge_range("A1:F1", f"Awarded Deals by Account Manager — {TODAY.strftime('%d %B %Y')}", fmt_title)
    am_ws.set_row(0, 28)

    am_cols   = ["Account Manager","Total Deals","Gross (QAR)","Net (QAR)","Net 2025 (QAR)","Net 2026 (QAR)"]
    am_widths = [30, 12, 20, 20, 20, 20]
    for c, col in enumerate(am_cols):
        am_ws.write(1, c, col, fmt_header)
        am_ws.set_column(c, c, am_widths[c])

    for i, row in am_df.reset_index(drop=True).iterrows():
        alt = (i % 2 == 1)
        am_ws.write(2+i, 0, row["Account Manager"], fmt_alt if alt else fmt_text)
        am_ws.write_number(2+i, 1, row["Count"],      fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+i, 2, row["Gross"],      fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+i, 3, row["Net"],        fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+i, 4, row["Net 2025"],   fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2+i, 5, row["Net 2026"],   fmt_alt_num if alt else fmt_num)

    # Totals
    tr = 2 + len(am_df)
    am_ws.write(tr, 0, "GRAND TOTAL", fmt_tot_lbl)
    am_ws.write_number(tr, 1, am_df["Count"].sum(),    fmt_tot_num)
    am_ws.write_number(tr, 2, am_df["Gross"].sum(),    fmt_tot_num)
    am_ws.write_number(tr, 3, am_df["Net"].sum(),      fmt_tot_num)
    am_ws.write_number(tr, 4, am_df["Net 2025"].sum(), fmt_tot_num)
    am_ws.write_number(tr, 5, am_df["Net 2026"].sum(), fmt_tot_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 — AWARD QUARTER
    # ════════════════════════════════════════════════════════════════════════
    aq_ws = wb.add_worksheet("Award Quarter")
    aq_ws.set_zoom(90)
    aq_ws.set_tab_color("#DAA520")
    aq_ws.merge_range("A1:E1", f"Net by Award Quarter — {TODAY.strftime('%d %B %Y')}", fmt_title)
    aq_ws.set_row(0, 28)

    aq_cols   = ["Year","Award Quarter","Deals","Gross (QAR)","Net (QAR)"]
    aq_widths = [8, 14, 8, 20, 20]
    for c, col in enumerate(aq_cols):
        aq_ws.write(1, c, col, fmt_header)
        aq_ws.set_column(c, c, aq_widths[c])

    for i, row in q_df.reset_index(drop=True).iterrows():
        f_txt = fmt_y25 if row["Year"] == "2025" else fmt_y26
        f_num = fmt_y25_num if row["Year"] == "2025" else fmt_y26_num
        aq_ws.write(2+i, 0, row["Year"],         f_txt)
        aq_ws.write(2+i, 1, row["Award Quarter"],f_txt)
        aq_ws.write_number(2+i, 2, row["Count"], f_num)
        aq_ws.write_number(2+i, 3, row["Gross"], f_num)
        aq_ws.write_number(2+i, 4, row["Net"],   f_num)

    tr = 2 + len(q_df)
    aq_ws.write(tr, 0, "GRAND TOTAL", fmt_tot_lbl)
    aq_ws.write(tr, 1, "", fmt_tot_lbl)
    aq_ws.write_number(tr, 2, q_df["Count"].sum(), fmt_tot_num)
    aq_ws.write_number(tr, 3, q_df["Gross"].sum(), fmt_tot_num)
    aq_ws.write_number(tr, 4, q_df["Net"].sum(),   fmt_tot_num)

    # New vs Renew section below
    nr_off = tr + 2
    aq_ws.merge_range(nr_off, 0, nr_off, 4, "New vs Renew Breakdown", fmt_title)
    for c, col in enumerate(["Year","Type","Deals","Gross (QAR)","Net (QAR)"]):
        aq_ws.write(nr_off+1, c, col, fmt_header)
    for i, row in nr_df.reset_index(drop=True).iterrows():
        f_txt = fmt_y25 if row["Year"] == "2025" else fmt_y26
        f_num = fmt_y25_num if row["Year"] == "2025" else fmt_y26_num
        aq_ws.write(nr_off+2+i, 0, row["Year"],  f_txt)
        aq_ws.write(nr_off+2+i, 1, row["Type"],  f_txt)
        aq_ws.write_number(nr_off+2+i, 2, row["Count"], f_num)
        aq_ws.write_number(nr_off+2+i, 3, row["Gross"], f_num)
        aq_ws.write_number(nr_off+2+i, 4, row["Net"],   f_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 — FULL AWARDED DEALS
    # ════════════════════════════════════════════════════════════════════════
    fd_ws = wb.add_worksheet("All Awarded Deals")
    fd_ws.set_zoom(85)
    fd_ws.set_tab_color("#6495ED")
    fd_ws.merge_range("A1:O1", f"All Awarded Deals — {TODAY.strftime('%d %B %Y')}", fmt_title)
    fd_ws.set_row(0, 28)

    fd_cols   = ["Year","SNo.","Account Name","Opportunity Name","Stage","Account Manager",
                 "Type","Total Gross","Total Net","Project Value",
                 "Award Quarter","Contracted","Contract Signed Quarter","ORF Number","Project Duration"]
    fd_widths = [7, 5, 32, 38, 36, 24, 8, 18, 18, 20, 14, 12, 24, 14, 16]
    for c, col in enumerate(fd_cols):
        fd_ws.write(1, c, col, fmt_header)
        fd_ws.set_column(c, c, fd_widths[c])
    fd_ws.freeze_panes(2, 0)

    for i, row in full_df.reset_index(drop=True).iterrows():
        f_txt = fmt_y25 if row["Year"] == "2025" else fmt_y26
        f_num = fmt_y25_num if row["Year"] == "2025" else fmt_y26_num
        for c, col in enumerate(fd_cols):
            val = row[col]
            if col in ("Total Gross","Total Net","Project Value","SNo.","Project Duration"):
                fd_ws.write_number(2+i, c, val if pd.notna(val) else 0, f_num)
            else:
                fd_ws.write(2+i, c, str(val) if pd.notna(val) else "", f_txt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 — AWARDED BREAKDOWN (styled, one row per DU per opportunity)
    # ════════════════════════════════════════════════════════════════════════
    CLR_AW = {
        "title_bg":    "1F3864",
        "hdr_deal":    "1F3864",
        "hdr_du":      "17375E",
        "hdr_finance": "1F4E79",
        "hdr_other":   "2E5FA3",
        "bu_fill":     "EDF2F9",
        "du_fill":     "E4ECF7",
        "num_fill":    "EBF5FB",
        "tot_fill":    "D5E8F5",
        "alt_a":       "F5F8FF",
        "alt_b":       "FFFFFF",
    }
    aw_fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR_AW["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    aw_fh_du      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR_AW["hdr_du"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    aw_fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR_AW["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    aw_fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+CLR_AW["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
    aw_fmt_title  = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+CLR_AW["title_bg"],"align":"center","valign":"vcenter"})

    def _aw_fmt(bg, top, num_fmt=None):
        d = {"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
        if num_fmt:
            d["num_format"] = num_fmt
            d["align"] = "right"
        return wb.add_format(d)

    aw_ft_a_first = _aw_fmt(CLR_AW["alt_a"], 2)
    aw_ft_a_next  = _aw_fmt(CLR_AW["alt_a"], 1)
    aw_ft_b_first = _aw_fmt(CLR_AW["alt_b"], 2)
    aw_ft_b_next  = _aw_fmt(CLR_AW["alt_b"], 1)
    aw_fn_a_first = _aw_fmt(CLR_AW["alt_a"], 2, "#,##0")
    aw_fn_a_next  = _aw_fmt(CLR_AW["alt_a"], 1, "#,##0")
    aw_fn_b_first = _aw_fmt(CLR_AW["alt_b"], 2, "#,##0")
    aw_fn_b_next  = _aw_fmt(CLR_AW["alt_b"], 1, "#,##0")
    aw_fbu_first  = _aw_fmt(CLR_AW["bu_fill"], 2)
    aw_fbu_next   = _aw_fmt(CLR_AW["bu_fill"], 1)
    aw_fdu_first  = _aw_fmt(CLR_AW["du_fill"], 2)
    aw_fdu_next   = _aw_fmt(CLR_AW["du_fill"], 1)
    aw_fxn_first  = _aw_fmt(CLR_AW["num_fill"], 2, "#,##0")
    aw_fxn_next   = _aw_fmt(CLR_AW["num_fill"], 1, "#,##0")
    aw_ft_tot     = wb.add_format({"bg_color":"#"+CLR_AW["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
    aw_ft_tot_blank = wb.add_format({"bg_color":"#"+CLR_AW["num_fill"],"border":1,"font_size":9})

    aw_bw = wb.add_worksheet("Awarded Breakdown")
    aw_bw.set_zoom(85)
    aw_bw.set_tab_color("#9370DB")

    aw_output_cols = [
        ("SNo.",               6,  "deal"),
        ("Account Name",      24,  "deal"),
        ("Opportunity Name",  36,  "deal"),
        ("BU",                36,  "du"),
        ("DU",                34,  "du"),
        ("Gross (breakdown)", 16,  "finance"),
        ("Net (breakdown)",   16,  "finance"),
        ("Total Gross",       15,  "finance"),
        ("Total Net",         15,  "finance"),
        ("Stage",             28,  "other"),
        ("Account Manager",   22,  "other"),
        ("Award Quarter",     10,  "other"),
        ("Contracted",        10,  "other"),
        ("Year",               8,  "other"),
    ]
    aw_hdr_fmt_map = {"deal": aw_fh_deal, "du": aw_fh_du, "finance": aw_fh_finance, "other": aw_fh_other}
    aw_ncols = len(aw_output_cols)
    aw_bw.merge_range(0, 0, 0, aw_ncols - 1, "Awarded Deals — Expanded by Delivery Unit", aw_fmt_title)
    aw_bw.set_row(0, 28)
    for c, (col_name, col_w, col_type) in enumerate(aw_output_cols):
        aw_bw.write(1, c, col_name, aw_hdr_fmt_map[col_type])
        aw_bw.set_column(c, c, col_w)
    aw_bw.set_row(1, 28)
    aw_bw.freeze_panes(2, 0)

    aw_exp = _expand_deals(df, "Opportunity Name")
    aw_bw.autofilter(1, 0, 1 + len(aw_exp), aw_ncols - 1)

    aw_col_map = {name: idx for idx, (name, _, __) in enumerate(aw_output_cols)}

    # Pre-compute Excel row range per deal for SUM formulas
    aw_deal_rows = {}
    for r_pos, (_, row) in enumerate(aw_exp.iterrows()):
        didx = row["_deal_idx"]
        xl_r = 2 + r_pos
        if didx not in aw_deal_rows:
            aw_deal_rows[didx] = [xl_r, xl_r]
        else:
            aw_deal_rows[didx][1] = xl_r
    aw_g_col = aw_col_map["Gross (breakdown)"]
    aw_n_col = aw_col_map["Net (breakdown)"]

    prev_deal_idx = None
    alt_toggle = False
    for r_pos, (_, row) in enumerate(aw_exp.iterrows()):
        didx     = row["_deal_idx"]
        is_first = bool(row["_is_first"])
        if didx != prev_deal_idx:
            alt_toggle = not alt_toggle
            prev_deal_idx = didx
        alt = alt_toggle
        ft  = (aw_ft_a_first if is_first else aw_ft_a_next) if alt else (aw_ft_b_first if is_first else aw_ft_b_next)
        fn  = (aw_fn_a_first if is_first else aw_fn_a_next) if alt else (aw_fn_b_first if is_first else aw_fn_b_next)
        fbu = aw_fbu_first if is_first else aw_fbu_next
        fdu = aw_fdu_first if is_first else aw_fdu_next
        fxn = aw_fxn_first if is_first else aw_fxn_next
        xl_r = 2 + r_pos

        def _aws(col_idx, val, fmt):
            if val is None or (isinstance(val, float) and pd.isna(val)):
                aw_bw.write_blank(xl_r, col_idx, None, fmt)
            else:
                aw_bw.write(xl_r, col_idx, str(val), fmt)

        def _awn(col_idx, val, fmt):
            v = _parse_num(val) if not isinstance(val, (int, float)) else val
            if v is None or (isinstance(v, float) and pd.isna(v)):
                aw_bw.write_blank(xl_r, col_idx, None, fmt)
            else:
                aw_bw.write_number(xl_r, col_idx, v, fmt)

        # SNo. only on first row; deal-level fields replicated on all rows
        _aws(aw_col_map["SNo."],             row.get("SNo.") if is_first else None, ft)
        _aws(aw_col_map["Account Name"],      row.get("Account Name"),      ft)
        _aws(aw_col_map["Opportunity Name"],  row.get("Opportunity Name"),  ft)
        _aws(aw_col_map["BU"],  row.get("BU_exp"),  fbu)
        _aws(aw_col_map["DU"],  row.get("DU_exp"),  fdu)
        _awn(aw_col_map["Gross (breakdown)"], row.get("Gross_exp"), fxn)
        _awn(aw_col_map["Net (breakdown)"],   row.get("Net_exp"),   fxn)
        if is_first:
            r0, r1 = aw_deal_rows[didx]
            gc = chr(65 + aw_g_col); nc = chr(65 + aw_n_col)
            aw_bw.write_formula(xl_r, aw_col_map["Total Gross"], f"=SUM({gc}{r0+1}:{gc}{r1+1})", aw_ft_tot)
            aw_bw.write_formula(xl_r, aw_col_map["Total Net"],   f"=SUM({nc}{r0+1}:{nc}{r1+1})", aw_ft_tot)
        else:
            aw_bw.write_blank(xl_r, aw_col_map["Total Gross"], None, aw_ft_tot_blank)
            aw_bw.write_blank(xl_r, aw_col_map["Total Net"],   None, aw_ft_tot_blank)
        _aws(aw_col_map["Stage"],          row.get("Stage"),          ft)
        _aws(aw_col_map["Account Manager"], row.get("Account Manager"), ft)
        _aws(aw_col_map["Award Quarter"],  row.get("Award Quarter"),  ft)
        _aws(aw_col_map["Contracted"],     row.get("Contracted"),     ft)
        _aws(aw_col_map["Year"],           row.get("Year"),           ft)

print(f"Done! Saved: {OUT_FILE}")
