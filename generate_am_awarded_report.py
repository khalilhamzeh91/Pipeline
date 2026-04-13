"""
AM Awarded Deals Report Generator
Based on data (3) (1).xlsx — Capability Sales / AM view for awarded deals.
Usage: python generate_am_awarded_report.py
"""

import pandas as pd
import re
import warnings
from datetime import date
warnings.filterwarnings("ignore")

INPUT_FILE = r"C:\Users\khali\Downloads\data (3) (1).xlsx"
COA_FILE   = r"C:\Users\khali\Downloads\Charter of Accounts 1 (1)2.xlsx"
TODAY      = date.today()
OUT_FILE   = rf"C:\Users\khali\Downloads\AM_Awarded_Report_{TODAY}.xlsx"

# ── CHARTER OF ACCOUNTS — DU → BU MAPPING ────────────────────────────────────
coa = pd.read_excel(COA_FILE)
coa.columns = coa.columns.str.strip()
coa["_code"] = coa["DU"].str.extract(r"(\d{6})")
DU_TO_BU = coa.dropna(subset=["_code"]).set_index("_code")["BU"].to_dict()

def du_to_bu(du_str):
    m = re.match(r"(\d{6})", str(du_str).strip())
    return DU_TO_BU.get(m.group(1), "Unknown") if m else "Unknown"

# ── HELPERS ───────────────────────────────────────────────────────────────────
def _split_field(value):
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return []
    return [x.strip() for x in re.split(r"\r\n|\r|\n", str(value)) if x.strip()]

def _clean_am_list(value):
    """Deduplicate and normalize AM names from newline-separated field."""
    if pd.isna(value) or str(value).strip() in ("", "nan"):
        return []
    seen = {}
    for name in re.split(r"\r\n|\r|\n", str(value)):
        name = name.strip()
        if not name:
            continue
        if "khalil" in name.lower():  name = "Khalil Hamzeh"
        elif "yazan" in name.lower(): name = "Yazan Al Razem"
        seen[name] = None
    return list(seen.keys())

def _normalize_am_cell(value):
    names = _clean_am_list(value)
    return "\n".join(names) if names else (value if pd.notna(value) else value)

# ── LOAD & CLEAN ──────────────────────────────────────────────────────────────
df = pd.read_excel(INPUT_FILE, sheet_name="Export")
df.columns = df.columns.str.strip()
# Drop summary/filter rows at the bottom (no Opportunity Name)
df = df.dropna(subset=["Opportunity Name"])
df = df[df["Opportunity Name"].astype(str).str.strip().ne("")]

df["Total Gross"] = pd.to_numeric(df["Total Gross"], errors="coerce").fillna(0)
df["Total Net"]   = pd.to_numeric(df["Total Net"],   errors="coerce").fillna(0)

pv_col = next((c for c in df.columns if "project value" in c.lower() or "contract value" in c.lower()), None)
df["Project Value"] = pd.to_numeric(df[pv_col], errors="coerce").fillna(0) if pv_col else 0

cc_col = next((c for c in df.columns if "client commitment" in c.lower()), None)
df["Client Commitment"] = pd.to_numeric(df[cc_col], errors="coerce").fillna(0) if cc_col else 0

df["Contracted"] = df["Contracted"].astype(str).str.strip()
df["Award Quarter"] = df["Award Quarter"].astype(str).str.strip()
df["New/Renew"] = df["New/Renew"].astype(str).str.strip()

# Normalize Capability Sales column at source
if "Capability Sales" in df.columns:
    df["Capability Sales"] = df["Capability Sales"].apply(_normalize_am_cell)

df = df.reset_index(drop=True)

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

# ── EXPLODE BY DU ─────────────────────────────────────────────────────────────
du_rows = []
for _, row in df.iterrows():
    dus   = _split_field(row.get("DU", ""))   or [str(row.get("DU", "Unknown"))]
    gross = _split_field(str(row.get("Gross (breakdown)", "")).replace(",", "")) or ["0"]
    net   = _split_field(str(row.get("Net (breakdown)", "")).replace(",", ""))   or ["0"]
    n = max(len(dus), len(gross), len(net))
    for i in range(n):
        du = dus[i] if i < len(dus) else dus[-1]
        try:   g_val = float(gross[i].replace(",", "") if i < len(gross) else "0")
        except: g_val = 0.0
        try:   n_val = float(net[i].replace(",", "") if i < len(net) else "0")
        except: n_val = 0.0
        du_rows.append({
            "BU": du_to_bu(du), "DU": du,
            "Gross": g_val, "Net": n_val,
            "Contracted":      row.get("Contracted", ""),
            "Award Quarter":   row.get("Award Quarter", ""),
            "New/Renew":       row.get("New/Renew", ""),
            "Account Name":    row.get("Account Name", ""),
            "Opportunity Name": row.get("Opportunity Name", ""),
        })
du_exp = pd.DataFrame(du_rows)

# ── SUMMARY TABLES ────────────────────────────────────────────────────────────
total_gross  = df["Total Gross"].sum()
total_net    = df["Total Net"].sum()
total_pv     = df["Project Value"].sum()
contracted_n = int((df["Contracted"] == "Yes").sum())
new_count    = int((df["New/Renew"] == "New").sum())
renew_count  = int((df["New/Renew"] == "Renew").sum())
total_deals  = len(df)

# By AM
am_agg = (
    am_exp.groupby("AM_exp")
    .agg(Count=("Opportunity Name", "nunique"),
         Gross=("Total Gross", "sum"),
         Net=("Total Net", "sum"),
         PV=("Project Value", "sum"))
    .reset_index().rename(columns={"AM_exp": "Account Manager"})
    .sort_values("Net", ascending=False)
)
contracted_am = (
    am_exp[am_exp["Contracted"] == "Yes"]
    .groupby("AM_exp")["Total Net"].sum().reset_index()
    .rename(columns={"AM_exp": "Account Manager", "Total Net": "Contracted Net"})
)
am_agg = am_agg.merge(contracted_am, on="Account Manager", how="left").fillna({"Contracted Net": 0})

# By Quarter
q_df = (
    df.groupby("Award Quarter")
    .agg(Count=("Opportunity Name", "count"),
         Gross=("Total Gross", "sum"),
         Net=("Total Net", "sum"))
    .reset_index().sort_values("Award Quarter")
)
contracted_q = (
    df[df["Contracted"] == "Yes"]
    .groupby("Award Quarter")["Total Net"].sum().reset_index()
    .rename(columns={"Total Net": "Contracted Net"})
)
q_df = q_df.merge(contracted_q, on="Award Quarter", how="left").fillna({"Contracted Net": 0})

# By New/Renew
nr_df = (
    df.groupby("New/Renew")
    .agg(Count=("Opportunity Name", "count"),
         Gross=("Total Gross", "sum"),
         Net=("Total Net", "sum"))
    .reset_index().sort_values("Net", ascending=False)
)

# By DU
du_totals = (
    du_exp.groupby(["BU", "DU"])[["Gross", "Net"]]
    .sum().reset_index().sort_values(["BU", "Net"], ascending=[True, False])
)
contracted_du = (
    du_exp[du_exp["Contracted"] == "Yes"]
    .groupby("DU")["Net"].sum().reset_index()
    .rename(columns={"Net": "Contracted Net"})
)
du_totals = du_totals.merge(contracted_du, on="DU", how="left").fillna({"Contracted Net": 0})

# Full table
full_df = df[[
    "SNo.", "Account Name", "Opportunity Name", "Capability Sales",
    "BU", "DU", "Total Gross", "Total Net", "Project Value",
    "New/Renew", "Award Quarter", "Contracted"
]].sort_values("Total Net", ascending=False)

fp_last_row = 2 + len(full_df)

# ── WRITE EXCEL ───────────────────────────────────────────────────────────────
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as writer:
    wb = writer.book

    # ── FORMATS ──────────────────────────────────────────────────────────────
    fmt_title   = wb.add_format({"bold": True, "font_size": 14, "font_color": "#FFFFFF",
                                  "bg_color": "#1a3a6b", "align": "center", "valign": "vcenter"})
    fmt_header  = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1a3a6b",
                                  "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_kpi_lbl = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "left"})
    fmt_kpi_val = wb.add_format({"bold": True, "bg_color": "#EBF0FB", "border": 1,
                                  "num_format": "#,##0", "align": "right"})
    fmt_num     = wb.add_format({"num_format": "#,##0", "border": 1, "align": "right"})
    fmt_text    = wb.add_format({"border": 1, "align": "left"})
    fmt_alt     = wb.add_format({"bg_color": "#F2F5FB", "border": 1, "align": "left"})
    fmt_alt_num = wb.add_format({"bg_color": "#F2F5FB", "num_format": "#,##0", "border": 1, "align": "right"})
    fmt_grn     = wb.add_format({"bg_color": "#E2EFDA", "border": 1, "align": "left"})
    fmt_grn_num = wb.add_format({"bg_color": "#E2EFDA", "num_format": "#,##0", "border": 1, "align": "right"})
    fmt_bu_lbl  = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "left"})
    fmt_bu_num  = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "num_format": "#,##0", "border": 1, "align": "right"})
    fmt_bu_hdr  = wb.add_format({"bold": True, "bg_color": "#2E5FA3", "font_color": "#FFFFFF",
                                  "border": 1, "align": "left", "font_size": 11})
    fmt_opp     = wb.add_format({"italic": True, "bg_color": "#FAFAFA", "border": 1, "align": "left", "indent": 2})
    fmt_opp_num = wb.add_format({"italic": True, "bg_color": "#FAFAFA", "num_format": "#,##0", "border": 1, "align": "right"})
    tot_fmt     = wb.add_format({"bold": True, "bg_color": "#1a3a6b", "font_color": "#FFFFFF",
                                  "num_format": "#,##0", "border": 1, "align": "right"})
    tot_lbl     = wb.add_format({"bold": True, "bg_color": "#1a3a6b", "font_color": "#FFFFFF",
                                  "border": 1, "align": "left"})

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
    ws.merge_range("A1:C1", f"AM Awarded Deals Summary — {TODAY.strftime('%d %B %Y')}", fmt_title)
    ws.set_row(0, 28)
    ws.write(2, 0, "Metric", fmt_header)
    ws.write(2, 1, "Value", fmt_header)
    ws.set_column(0, 0, 34)
    ws.set_column(1, 1, 22)

    _fp = "'Full Awarded'"
    _n  = fp_last_row
    kpi_rows = [
        ("Total Awarded Deals",         f"=COUNTA({_fp}!C3:C{_n})"),
        ("Total Gross (QAR)",            f"=SUM({_fp}!G3:G{_n})"),
        ("Total Net (QAR)",              f"=SUM({_fp}!H3:H{_n})"),
        ("Total Project Value (QAR)",    f"=SUM({_fp}!I3:I{_n})"),
        ("Contracted Deals",             f'=COUNTIF({_fp}!L3:L{_n},"Yes")'),
        ("New Deals",                    f'=COUNTIF({_fp}!J3:J{_n},"New")'),
        ("Renew Deals",                  f'=COUNTIF({_fp}!J3:J{_n},"Renew")'),
    ]
    for i, (label, formula) in enumerate(kpi_rows):
        ws.write(3 + i, 0, label, fmt_kpi_lbl)
        ws.write_formula(3 + i, 1, formula, fmt_kpi_val)

    # Quarter table (right side)
    ws.merge_range("D1:H1", "Pipeline by Award Quarter", fmt_title)
    write_header_row(ws, 2, ["Quarter", "Count", "Gross (QAR)", "Net (QAR)", "Contracted Net (QAR)"])
    ws.set_column(3, 3, 14); ws.set_column(4, 4, 8)
    ws.set_column(5, 5, 18); ws.set_column(6, 6, 18); ws.set_column(7, 7, 22)
    for r, row in q_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        ws.write(3 + r, 3, row["Award Quarter"],    fmt_alt if alt else fmt_text)
        ws.write_number(3 + r, 4, row["Count"],     fmt_alt_num if alt else fmt_num)
        ws.write_number(3 + r, 5, row["Gross"],     fmt_alt_num if alt else fmt_num)
        ws.write_number(3 + r, 6, row["Net"],       fmt_alt_num if alt else fmt_num)
        ws.write_number(3 + r, 7, row["Contracted Net"], fmt_alt_num if alt else fmt_num)
    _qr = 3 + len(q_df)
    ws.write(_qr, 3, "TOTAL", tot_lbl)
    ws.write_formula(_qr, 4, f"=SUM(E4:E{_qr})", tot_fmt)
    ws.write_formula(_qr, 5, f"=SUM(F4:F{_qr})", tot_fmt)
    ws.write_formula(_qr, 6, f"=SUM(G4:G{_qr})", tot_fmt)
    ws.write_formula(_qr, 7, f"=SUM(H4:H{_qr})", tot_fmt)

    # New vs Renew table below quarter table
    _nr_off = _qr + 2
    ws.merge_range(_nr_off, 3, _nr_off, 7, "New vs Renew Breakdown", fmt_bu_hdr)
    write_header_row(ws, _nr_off + 1, ["Type", "Count", "Gross (QAR)", "Net (QAR)", ""])
    for r, row in nr_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        ws.write(_nr_off + 2 + r, 3, row["New/Renew"], fmt_alt if alt else fmt_text)
        ws.write_number(_nr_off + 2 + r, 4, row["Count"], fmt_alt_num if alt else fmt_num)
        ws.write_number(_nr_off + 2 + r, 5, row["Gross"], fmt_alt_num if alt else fmt_num)
        ws.write_number(_nr_off + 2 + r, 6, row["Net"],   fmt_alt_num if alt else fmt_num)
        ws.write(_nr_off + 2 + r, 7, "", fmt_alt if alt else fmt_text)
    _nrr = _nr_off + 2 + len(nr_df)
    ws.write(_nrr, 3, "TOTAL", tot_lbl)
    ws.write_formula(_nrr, 4, f"=SUM(E{_nr_off+3}:E{_nrr})", tot_fmt)
    ws.write_formula(_nrr, 5, f"=SUM(F{_nr_off+3}:F{_nrr})", tot_fmt)
    ws.write_formula(_nrr, 6, f"=SUM(G{_nr_off+3}:G{_nrr})", tot_fmt)
    ws.write(_nrr, 7, "", tot_lbl)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 2 — BY ACCOUNT MANAGER
    # ════════════════════════════════════════════════════════════════════════
    am_ws = wb.add_worksheet("By Account Manager")
    am_ws.set_zoom(90)
    am_ws.set_tab_color("#FF8C00")
    am_ws.merge_range("A1:F1",
        f"Awarded Deals by Account Manager (Capability Sales) — {TODAY.strftime('%d %B %Y')}", fmt_title)
    am_ws.set_row(0, 28)

    am_cols   = ["Account Manager", "Deals", "Gross (QAR)", "Net (QAR)", "Project Value (QAR)", "Contracted Net (QAR)"]
    am_widths = [34, 8, 20, 20, 22, 22]
    write_header_row(am_ws, 1, am_cols, am_widths)

    _am_data_start = 3
    for r, row in am_agg.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        am_ws.write(2 + r, 0, row["Account Manager"], fmt_alt if alt else fmt_text)
        am_ws.write_number(2 + r, 1, row["Count"],          fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2 + r, 2, row["Gross"],          fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2 + r, 3, row["Net"],            fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2 + r, 4, row["PV"],             fmt_alt_num if alt else fmt_num)
        am_ws.write_number(2 + r, 5, row["Contracted Net"], fmt_alt_num if alt else fmt_num)
    _am_rN = 2 + len(am_agg)
    am_ws.write(_am_rN, 0, "TOTAL", tot_lbl)
    am_ws.write_formula(_am_rN, 1, f"=SUM(B{_am_data_start}:B{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 2, f"=SUM(C{_am_data_start}:C{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 3, f"=SUM(D{_am_data_start}:D{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 4, f"=SUM(E{_am_data_start}:E{_am_rN})", tot_fmt)
    am_ws.write_formula(_am_rN, 5, f"=SUM(F{_am_data_start}:F{_am_rN})", tot_fmt)

    # Deal detail by AM
    det_off = _am_rN + 2
    am_ws.merge_range(det_off, 0, det_off, 6, "Deal Detail by Account Manager", fmt_bu_hdr)
    am_ws.set_row(det_off, 22)
    det_cols   = ["Account Manager", "Account Name", "Opportunity", "Quarter", "Type", "Gross (QAR)", "Net (QAR)"]
    det_widths = [30, 28, 44, 12, 10, 18, 18]
    write_header_row(am_ws, det_off + 1, det_cols, det_widths)

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
        _dF = "+".join("F" + str(r + 1) for r in deal_rows_xl)
        _dG = "+".join("G" + str(r + 1) for r in deal_rows_xl)
        am_ws.write(bu_xl, 0, am_name, fmt_bu_lbl)
        for cc in range(1, 7):
            am_ws.write(bu_xl, cc, "", fmt_bu_lbl)
        am_ws.write_formula(bu_xl, 5, "=" + _dF, fmt_bu_num)
        am_ws.write_formula(bu_xl, 6, "=" + _dG, fmt_bu_num)
        _det_bu_rows.append(bu_xl + 1)

        for r_pos, (_, row) in zip(deal_rows_xl, am_grp.iterrows()):
            alt = (r_pos % 2 == 1)
            tf  = fmt_alt if alt else fmt_text
            nf  = fmt_alt_num if alt else fmt_num
            am_ws.write(r_pos, 0, am_name,                                                   tf)
            am_ws.write(r_pos, 1, str(row["Account Name"])    if pd.notna(row["Account Name"])    else "", tf)
            am_ws.write(r_pos, 2, str(row["Opportunity Name"]) if pd.notna(row["Opportunity Name"]) else "", tf)
            am_ws.write(r_pos, 3, str(row["Award Quarter"])   if pd.notna(row["Award Quarter"])   else "", tf)
            am_ws.write(r_pos, 4, str(row["New/Renew"])       if pd.notna(row["New/Renew"])       else "", tf)
            am_ws.write_number(r_pos, 5, row["Total Gross"],  nf)
            am_ws.write_number(r_pos, 6, row["Total Net"],    nf)

    _gt_F = "+".join("F" + str(r) for r in _det_bu_rows)
    _gt_G = "+".join("G" + str(r) for r in _det_bu_rows)
    _gt_row = _det_pos
    am_ws.write(_gt_row, 0, "GRAND TOTAL", tot_lbl)
    for cc in range(1, 5):
        am_ws.write(_gt_row, cc, "", tot_lbl)
    am_ws.write_formula(_gt_row, 5, "=" + _gt_F, tot_fmt)
    am_ws.write_formula(_gt_row, 6, "=" + _gt_G, tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 3 — DU BREAKDOWN
    # ════════════════════════════════════════════════════════════════════════
    du_ws = wb.add_worksheet("DU Breakdown")
    du_ws.set_zoom(90)
    du_ws.set_tab_color("#FF8C00")
    du_ws.merge_range("A1:E1", "Gross & Net Breakdown by BU / Delivery Unit", fmt_title)
    du_ws.set_row(0, 28)

    du_cols   = ["BU", "Delivery Unit / Opportunity", "Gross (QAR)", "Net (QAR)", "Contracted Net (QAR)"]
    du_widths = [42, 52, 20, 20, 22]
    write_header_row(du_ws, 1, du_cols, du_widths)

    _du_layout = []
    _du_pos = 0
    for bu_name, bu_grp in du_totals.groupby("BU"):
        bu_r0 = 2 + _du_pos; _du_pos += 1
        du_list = []
        for _, drow in bu_grp.iterrows():
            du_r0 = 2 + _du_pos; _du_pos += 1
            du_deals = du_exp[du_exp["DU"] == drow["DU"]].copy()
            opp_rows = []
            for _ in du_deals.itertuples():
                opp_rows.append(2 + _du_pos); _du_pos += 1
            du_list.append((drow, du_r0, opp_rows))
        _du_layout.append((bu_name, bu_r0, bu_grp, du_list))
    _du_grand_r = 2 + _du_pos

    _du_bu_rows = []
    for bu_name, bu_r0, bu_grp, du_list in _du_layout:
        _du_r1s = [dr0 + 1 for (_, dr0, _) in du_list]
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
                _r1 = min(opp_rows) + 1; _rN = max(opp_rows) + 1
                def _os(c, r1=_r1, rN=_rN): return f"=SUM({c}{r1}:{c}{rN})"
            else:
                def _os(c): return "=0"
            du_ws.write(du_r0, 0, "", fmt_alt if alt else fmt_text)
            du_ws.write(du_r0, 1, drow["DU"], fmt_alt if alt else fmt_text)
            du_ws.write_formula(du_r0, 2, _os("C"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 3, _os("D"), fmt_alt_num if alt else fmt_num)
            du_ws.write_formula(du_r0, 4, _os("E"), fmt_alt_num if alt else fmt_num)

            du_deals = du_exp[du_exp["DU"] == drow["DU"]].copy()
            for opp_r0, (_, deal) in zip(opp_rows, du_deals.iterrows()):
                opp_label = f"  ↳  {deal['Opportunity Name']}"
                c_net = deal["Net"] if str(deal.get("Contracted", "")).strip() == "Yes" else 0
                du_ws.write(opp_r0, 0, "", fmt_opp)
                du_ws.write(opp_r0, 1, opp_label, fmt_opp)
                du_ws.write_number(opp_r0, 2, deal["Gross"], fmt_opp_num)
                du_ws.write_number(opp_r0, 3, deal["Net"],   fmt_opp_num)
                du_ws.write_number(opp_r0, 4, c_net,         fmt_opp_num)

    _gt_C = "+".join("C" + str(r) for r in _du_bu_rows)
    _gt_D = "+".join("D" + str(r) for r in _du_bu_rows)
    _gt_E = "+".join("E" + str(r) for r in _du_bu_rows)
    t = _du_grand_r
    du_ws.write(t, 0, "GRAND TOTAL", tot_lbl); du_ws.write(t, 1, "", tot_lbl)
    du_ws.write_formula(t, 2, "=" + _gt_C, tot_fmt)
    du_ws.write_formula(t, 3, "=" + _gt_D, tot_fmt)
    du_ws.write_formula(t, 4, "=" + _gt_E, tot_fmt)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 4 — QUARTERLY PLAN
    # ════════════════════════════════════════════════════════════════════════
    qw = wb.add_worksheet("Quarterly Plan")
    qw.set_zoom(90)
    qw.set_tab_color("#DAA520")
    qw.merge_range("A1:E1", f"Awarded Deals by Quarter — {TODAY.strftime('%d %B %Y')}", fmt_title)
    qw.set_row(0, 28)

    write_header_row(qw, 1,
        ["Quarter", "Count", "Gross (QAR)", "Net (QAR)", "Contracted Net (QAR)"],
        [14, 8, 20, 20, 22])
    for r, row in q_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        qw.write(2 + r, 0, row["Award Quarter"],    fmt_alt if alt else fmt_text)
        qw.write_number(2 + r, 1, row["Count"],     fmt_alt_num if alt else fmt_num)
        qw.write_number(2 + r, 2, row["Gross"],     fmt_alt_num if alt else fmt_num)
        qw.write_number(2 + r, 3, row["Net"],       fmt_alt_num if alt else fmt_num)
        qw.write_number(2 + r, 4, row["Contracted Net"], fmt_alt_num if alt else fmt_num)
    _qr2 = 2 + len(q_df)
    qw.write(_qr2, 0, "TOTAL", tot_lbl)
    qw.write_formula(_qr2, 1, f"=SUM(B3:B{_qr2})", tot_fmt)
    qw.write_formula(_qr2, 2, f"=SUM(C3:C{_qr2})", tot_fmt)
    qw.write_formula(_qr2, 3, f"=SUM(D3:D{_qr2})", tot_fmt)
    qw.write_formula(_qr2, 4, f"=SUM(E3:E{_qr2})", tot_fmt)

    # New vs Renew section below
    _nr2_off = _qr2 + 2
    qw.merge_range(_nr2_off, 0, _nr2_off, 4, "New vs Renew", fmt_bu_hdr)
    write_header_row(qw, _nr2_off + 1, ["Type", "Count", "Gross (QAR)", "Net (QAR)", ""])
    for r, row in nr_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        qw.write(_nr2_off + 2 + r, 0, row["New/Renew"], fmt_alt if alt else fmt_text)
        qw.write_number(_nr2_off + 2 + r, 1, row["Count"], fmt_alt_num if alt else fmt_num)
        qw.write_number(_nr2_off + 2 + r, 2, row["Gross"], fmt_alt_num if alt else fmt_num)
        qw.write_number(_nr2_off + 2 + r, 3, row["Net"],   fmt_alt_num if alt else fmt_num)
        qw.write(_nr2_off + 2 + r, 4, "", fmt_alt if alt else fmt_text)
    _nrr2 = _nr2_off + 2 + len(nr_df)
    qw.write(_nrr2, 0, "TOTAL", tot_lbl)
    qw.write_formula(_nrr2, 1, f"=SUM(B{_nr2_off+3}:B{_nrr2})", tot_fmt)
    qw.write_formula(_nrr2, 2, f"=SUM(C{_nr2_off+3}:C{_nrr2})", tot_fmt)
    qw.write_formula(_nrr2, 3, f"=SUM(D{_nr2_off+3}:D{_nrr2})", tot_fmt)
    qw.write(_nrr2, 4, "", tot_lbl)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 5 — AM BREAKDOWN (styled, one row per AM per deal)
    # ════════════════════════════════════════════════════════════════════════
    # Color formats for AM Breakdown
    am_fmt_title = wb.add_format({"bold": True, "font_size": 13, "font_color": "#FFFFFF",
                                   "bg_color": "#1a3a6b", "align": "center", "valign": "vcenter"})
    am_fh_deal    = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1a3a6b",
                                    "border": 1, "align": "center", "text_wrap": True})
    am_fh_am      = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#FF8C00",
                                    "border": 1, "align": "center", "text_wrap": True})
    am_fh_finance = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#2E7D32",
                                    "border": 1, "align": "center", "text_wrap": True})
    am_fh_other   = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#4472C4",
                                    "border": 1, "align": "center", "text_wrap": True})

    am_fd_deal    = wb.add_format({"border": 1, "align": "left",  "bg_color": "#DEEAF1"})
    am_fd_am      = wb.add_format({"border": 1, "align": "left",  "bg_color": "#FCE4D6"})
    am_fd_finance = wb.add_format({"border": 1, "align": "right", "bg_color": "#E2EFDA", "num_format": "#,##0"})
    am_fd_other   = wb.add_format({"border": 1, "align": "left",  "bg_color": "#EBF0FB"})

    am_tot_lbl    = wb.add_format({"bold": True, "bg_color": "#1a3a6b", "font_color": "#FFFFFF",
                                    "border": 1, "align": "left"})
    am_tot_num    = wb.add_format({"bold": True, "bg_color": "#1a3a6b", "font_color": "#FFFFFF",
                                    "border": 1, "align": "right", "num_format": "#,##0"})

    am_bw = wb.add_worksheet("AM Breakdown")
    am_bw.set_zoom(90)
    am_bw.set_tab_color("#FF8C00")

    am_bd_cols = [
        ("SNo.",            6,  "deal"),
        ("Account Name",   28,  "deal"),
        ("Opportunity",    44,  "deal"),
        ("Capability Sales", 30, "am"),
        ("New/Renew",      12,  "other"),
        ("Quarter",        10,  "other"),
        ("Contracted",     12,  "other"),
        ("All Cap. Sales", 28,  "other"),
        ("Gross (QAR)",    18,  "finance"),
        ("Net (QAR)",      18,  "finance"),
        ("Project Value (QAR)", 20, "finance"),
    ]
    am_nc = len(am_bd_cols)
    am_hfmt = {"deal": am_fh_deal, "am": am_fh_am, "finance": am_fh_finance, "other": am_fh_other}

    am_bw.merge_range(0, 0, 0, am_nc - 1,
        "AM Awarded Deals — Expanded by Capability Sales", am_fmt_title)
    am_bw.set_row(0, 28)
    for c, (cn, cw, ct) in enumerate(am_bd_cols):
        am_bw.write(1, c, cn, am_hfmt[ct])
        am_bw.set_column(c, c, cw)
    am_bw.set_row(1, 28)
    am_bw.freeze_panes(2, 0)

    am_cmap = {cn: c for c, (cn, _, _) in enumerate(am_bd_cols)}

    _bw_data_start = 3
    _bw_gross_rows = []
    _bw_net_rows   = []

    for xl_r, (_, row) in enumerate(am_exp.iterrows(), start=2):
        def _amws(col_name, val, fmt):
            c = am_cmap[col_name]
            if val is None or (isinstance(val, float) and pd.isna(val)):
                am_bw.write_blank(xl_r, c, None, fmt)
            elif isinstance(val, (int, float)):
                am_bw.write_number(xl_r, c, val, fmt)
            else:
                am_bw.write(xl_r, c, str(val), fmt)

        ft  = am_fd_deal
        fam = am_fd_am
        ff  = am_fd_finance
        fo  = am_fd_other

        _amws("SNo.",            row.get("SNo."),            ft)
        _amws("Account Name",    row.get("Account Name"),    ft)
        _amws("Opportunity",     row.get("Opportunity Name"), ft)
        _amws("Capability Sales", row.get("AM_exp"),         fam)
        _amws("New/Renew",       row.get("New/Renew"),       fo)
        _amws("Quarter",         row.get("Award Quarter"),   fo)
        _amws("Contracted",      row.get("Contracted"),      fo)
        _amws("All Cap. Sales",  row.get("Capability Sales"), fo)
        _amws("Gross (QAR)",     row.get("Total Gross"),     ff)
        _amws("Net (QAR)",       row.get("Total Net"),       ff)
        _amws("Project Value (QAR)", row.get("Project Value"), ff)
        _bw_gross_rows.append(xl_r + 1)
        _bw_net_rows.append(xl_r + 1)

    _bw_last = 2 + len(am_exp)
    am_bw.write(_bw_last, am_cmap["SNo."],        "TOTAL", am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["Account Name"],"",      am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["Opportunity"], "",      am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["Capability Sales"], "", am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["New/Renew"],   "",      am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["Quarter"],     "",      am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["Contracted"],  "",      am_tot_lbl)
    am_bw.write(_bw_last, am_cmap["All Cap. Sales"], "",   am_tot_lbl)
    gc = chr(65 + am_cmap["Gross (QAR)"])
    nc = chr(65 + am_cmap["Net (QAR)"])
    pc = chr(65 + am_cmap["Project Value (QAR)"])
    am_bw.write_formula(_bw_last, am_cmap["Gross (QAR)"],
        f"=SUM({gc}{_bw_data_start}:{gc}{_bw_last})", am_tot_num)
    am_bw.write_formula(_bw_last, am_cmap["Net (QAR)"],
        f"=SUM({nc}{_bw_data_start}:{nc}{_bw_last})", am_tot_num)
    am_bw.write_formula(_bw_last, am_cmap["Project Value (QAR)"],
        f"=SUM({pc}{_bw_data_start}:{pc}{_bw_last})", am_tot_num)

    # ════════════════════════════════════════════════════════════════════════
    # SHEET 6 — FULL AWARDED
    # ════════════════════════════════════════════════════════════════════════
    pw = wb.add_worksheet("Full Awarded")
    pw.set_zoom(90)
    pw.set_tab_color("#1a3a6b")
    pw.merge_range("A1:L1", "Full AM Awarded Deals — All Opportunities", fmt_title)
    pw.set_row(0, 28)
    pw.freeze_panes(2, 0)

    full_cols   = ["#", "Account Name", "Opportunity", "Capability Sales",
                   "BU", "DU", "Gross (QAR)", "Net (QAR)", "Project Value (QAR)",
                   "New/Renew", "Quarter", "Contracted"]
    full_widths = [5, 28, 44, 28, 30, 36, 18, 18, 20, 12, 10, 12]
    write_header_row(pw, 1, full_cols, full_widths)
    pw.autofilter(1, 0, 1 + len(full_df), len(full_cols) - 1)

    for r, row in full_df.reset_index(drop=True).iterrows():
        alt = (r % 2 == 1)
        tf  = fmt_alt if alt else fmt_text
        nf  = fmt_alt_num if alt else fmt_num
        pw.write(2 + r, 0,  row.get("SNo.", r + 1),               tf)
        pw.write(2 + r, 1,  str(row["Account Name"])    if pd.notna(row["Account Name"])    else "", tf)
        pw.write(2 + r, 2,  str(row["Opportunity Name"]) if pd.notna(row["Opportunity Name"]) else "", tf)
        pw.write(2 + r, 3,  str(row["Capability Sales"]) if pd.notna(row["Capability Sales"]) else "", tf)
        pw.write(2 + r, 4,  str(row["BU"])              if pd.notna(row["BU"])              else "", tf)
        pw.write(2 + r, 5,  str(row["DU"])              if pd.notna(row["DU"])              else "", tf)
        pw.write_number(2 + r, 6,  row["Total Gross"],    nf)
        pw.write_number(2 + r, 7,  row["Total Net"],      nf)
        pw.write_number(2 + r, 8,  row["Project Value"],  nf)
        pw.write(2 + r, 9,  str(row["New/Renew"])        if pd.notna(row["New/Renew"])       else "", tf)
        pw.write(2 + r, 10, str(row["Award Quarter"])    if pd.notna(row["Award Quarter"])   else "", tf)
        pw.write(2 + r, 11, str(row["Contracted"])       if pd.notna(row["Contracted"])      else "", tf)

    _fp_last = 2 + len(full_df)
    pw.write(_fp_last, 0, "TOTAL", tot_lbl)
    for cc in range(1, 6):
        pw.write(_fp_last, cc, "", tot_lbl)
    pw.write_formula(_fp_last, 6,  f"=SUM(G3:G{_fp_last})", tot_fmt)
    pw.write_formula(_fp_last, 7,  f"=SUM(H3:H{_fp_last})", tot_fmt)
    pw.write_formula(_fp_last, 8,  f"=SUM(I3:I{_fp_last})", tot_fmt)
    for cc in range(9, 12):
        pw.write(_fp_last, cc, "", tot_lbl)

print(f"\nDone! Report saved to:\n{OUT_FILE}")
