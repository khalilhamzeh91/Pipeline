"""
Weekly Pipeline Review Dashboard
Drop your Excel file and get instant analysis.
Run: streamlit run pipeline_dashboard.py
"""

import re
import io
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date
from difflib import SequenceMatcher
import warnings
warnings.filterwarnings("ignore")

BOOK3_MONTHS = ["Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
BOOK3_SKIP   = ["Total","Grand Total","Existing Renewal Total",
                "Opportunity (ORF) Total","Opportunity Pipeline Total"]

@st.cache_data
def load_book3(file):
    raw = pd.read_excel(file, header=None)
    cols = ["_","BU","Project Type","Project Name"] + BOOK3_MONTHS + ["Grand Total"]
    raw.columns = cols[:len(raw.columns)]
    raw = raw.iloc[2:].reset_index(drop=True)
    current_bu, current_type = None, None
    rows = []
    for _, r in raw.iterrows():
        bu    = str(r["BU"]).strip()           if pd.notna(r["BU"])           else ""
        ptype = str(r["Project Type"]).strip() if pd.notna(r["Project Type"]) else ""
        name  = str(r["Project Name"]).strip() if pd.notna(r["Project Name"]) else ""
        if bu and bu != "nan" and not any(k in bu for k in BOOK3_SKIP):
            current_bu = bu
        if ptype and ptype != "nan" and not any(k in ptype for k in BOOK3_SKIP):
            current_type = ptype
        if not name or name == "nan" or any(k in name for k in BOOK3_SKIP):
            continue
        if any(k in (ptype or "") for k in BOOK3_SKIP):
            continue
        row_data = {"BU": current_bu, "Project Type": current_type, "Project Name": name}
        for m in BOOK3_MONTHS:
            val = r.get(m, None)
            try: row_data[m] = float(val) if pd.notna(val) else 0.0
            except: row_data[m] = 0.0
        try: row_data["Grand Total"] = float(r.get("Grand Total", 0)) if pd.notna(r.get("Grand Total")) else 0.0
        except: row_data["Grand Total"] = 0.0
        rows.append(row_data)
    return pd.DataFrame(rows)

def _book3_best_match(name, candidates, threshold=0.55):
    def _clean(s): return set(re.sub(r"[^a-z0-9]"," ",str(s).lower()).split())
    tokens = _clean(name)
    best, best_score = None, 0.0
    for cand in candidates:
        cand_tokens = _clean(cand)
        overlap = len(tokens & cand_tokens) / max(len(tokens | cand_tokens), 1)
        seq = SequenceMatcher(None, name.lower(), cand.lower()).ratio()
        score = max(overlap, seq)
        if score > best_score:
            best, best_score = cand, score
    return (best, round(best_score, 2)) if best_score >= threshold else (None, 0.0)

def build_book3_mapping(book3_df, pipeline_df, awarded_df):
    pipe_names  = pipeline_df["Lead/Opp Name"].dropna().tolist() if pipeline_df is not None else []
    award_names = awarded_df["Opportunity Name"].dropna().tolist() if awarded_df is not None and not awarded_df.empty else []
    rows = []
    for _, b3 in book3_df.iterrows():
        pipe_match,  pipe_score  = _book3_best_match(b3["Project Name"], pipe_names)
        award_match, award_score = _book3_best_match(b3["Project Name"], award_names)
        pipe_row  = pipeline_df[pipeline_df["Lead/Opp Name"]==pipe_match].iloc[0]  if pipe_match  and pipeline_df is not None else None
        award_row = awarded_df[awarded_df["Opportunity Name"]==award_match].iloc[0] if award_match and awarded_df is not None and not awarded_df.empty else None
        row = {"Book3 BU": b3["BU"], "Book3 Project Type": b3["Project Type"],
               "Book3 Project Name": b3["Project Name"], "Book3 Grand Total": b3["Grand Total"]}
        for m in BOOK3_MONTHS: row[f"Book3 {m}"] = b3[m]
        row["Pipeline Match"]        = pipe_match  or ""
        row["Pipeline Score"]        = pipe_score
        row["Pipeline Account"]      = str(pipe_row["Account Name"])       if pipe_row is not None else ""
        row["Pipeline Gross (QAR)"]  = float(pipe_row["Total Gross"])      if pipe_row is not None else 0.0
        row["Pipeline Net (QAR)"]    = float(pipe_row["Total Net"])        if pipe_row is not None else 0.0
        row["Pipeline Stage"]        = str(pipe_row["Stage"])              if pipe_row is not None else ""
        row["Pipeline AM"]           = str(pipe_row["Account Manager"])    if pipe_row is not None else ""
        row["Pipeline Quarter"]      = str(pipe_row["Closure Due Quarter"])if pipe_row is not None else ""
        row["Awarded Match"]         = award_match or ""
        row["Awarded Score"]         = award_score
        row["Awarded Account"]       = str(award_row["Account Name"])      if award_row is not None else ""
        row["Awarded Gross (QAR)"]   = float(award_row["Total Gross"])     if award_row is not None else 0.0
        row["Awarded Net (QAR)"]     = float(award_row["Total Net"])       if award_row is not None else 0.0
        row["Awarded Stage"]         = str(award_row["Stage"])             if award_row is not None else ""
        row["Awarded AM"]            = str(award_row["Account Manager"])   if award_row is not None else ""
        row["Awarded Year"]          = str(award_row["Year"])              if award_row is not None else ""
        rows.append(row)
    return pd.DataFrame(rows)

def _export_mapping_excel(map_df, today):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_title     = wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF","bg_color":"#1a3a6b","align":"center","valign":"vcenter"})
        fmt_header    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1a3a6b","border":1,"align":"center","valign":"vcenter","text_wrap":True})
        fmt_mhdr      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E5FA3","border":1,"align":"center","text_wrap":True})
        fmt_text      = wb.add_format({"border":1,"align":"left"})
        fmt_matched   = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"})
        fmt_matched_n = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"})
        fmt_matched_s = wb.add_format({"bg_color":"#E2EFDA","num_format":"0.00","border":1,"align":"center"})
        fmt_partial   = wb.add_format({"bg_color":"#FFF2CC","border":1,"align":"left"})
        fmt_partial_n = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right"})
        fmt_partial_s = wb.add_format({"bg_color":"#FFF2CC","num_format":"0.00","border":1,"align":"center"})
        fmt_nomatch   = wb.add_format({"bg_color":"#FFE0E0","border":1,"align":"left"})
        fmt_nomatch_n = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right"})
        fmt_nomatch_s = wb.add_format({"bg_color":"#FFE0E0","num_format":"0.00","border":1,"align":"center"})
        fmt_neg_grn   = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
        fmt_neg_yel   = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
        fmt_neg_red   = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})

        ws = wb.add_worksheet("Book3 Mapping")
        writer.sheets["Book3 Mapping"] = ws
        ws.set_zoom(80); ws.set_tab_color("#FF8C00"); ws.freeze_panes(3, 4)
        total_cols = 4 + len(BOOK3_MONTHS) + 16
        ws.merge_range(0,0,0,total_cols-1,f"Book3 ↔ Pipeline & Awarded Mapping — {today.strftime('%d %B %Y')}",fmt_title); ws.set_row(0,28)
        ws.write(1,0,"🟢 Strong match (score ≥ 0.70)",fmt_matched); ws.write(1,1,"🟡 Partial match (score 0.55–0.69)",fmt_partial); ws.write(1,2,"🔴 No match (score < 0.55)",fmt_nomatch); ws.write(1,3,"",fmt_text); ws.set_row(1,18)
        headers = (["Book3 BU","Book3 Project Type","Book3 Project Name","Grand Total (QAR)"] + BOOK3_MONTHS +
                   ["Pipeline Match","Score","Account","Gross (QAR)","Net (QAR)","Stage","AM","Quarter"] +
                   ["Awarded Match","Score","Account","Gross (QAR)","Net (QAR)","Stage","AM","Year"])
        widths  = [36,22,42,18]+[10]*len(BOOK3_MONTHS)+[38,8,28,16,16,20,22,10]+[38,8,28,16,16,20,22,8]
        for c,(h,w) in enumerate(zip(headers,widths)):
            ws.write(2,c,h,fmt_mhdr if h in BOOK3_MONTHS else fmt_header); ws.set_column(c,c,w)
        ws.set_row(2,20); ws.autofilter(2,0,2+len(map_df),len(headers)-1)
        for i, row in map_df.reset_index(drop=True).iterrows():
            sc = max(row["Pipeline Score"], row["Awarded Score"])
            if sc >= 0.70:   ft,fn,fs,fn_neg = fmt_matched,fmt_matched_n,fmt_matched_s,fmt_neg_grn
            elif sc >= 0.55: ft,fn,fs,fn_neg = fmt_partial,fmt_partial_n,fmt_partial_s,fmt_neg_yel
            else:            ft,fn,fs,fn_neg = fmt_nomatch,fmt_nomatch_n,fmt_nomatch_s,fmt_neg_red
            c=0
            ws.write(3+i,c,str(row["Book3 BU"]) or "",ft);          c+=1
            ws.write(3+i,c,str(row["Book3 Project Type"]) or "",ft); c+=1
            ws.write(3+i,c,str(row["Book3 Project Name"]) or "",ft); c+=1
            ws.write_number(3+i,c,row["Book3 Grand Total"],fn_neg);  c+=1
            for m in BOOK3_MONTHS:
                ws.write_number(3+i,c,row.get(f"Book3 {m}",0),fn_neg); c+=1
            ws.write(3+i,c,str(row["Pipeline Match"]) or "",ft);     c+=1
            ws.write_number(3+i,c,row["Pipeline Score"],fs);          c+=1
            ws.write(3+i,c,str(row["Pipeline Account"]) or "",ft);   c+=1
            ws.write_number(3+i,c,row["Pipeline Gross (QAR)"],fn);   c+=1
            ws.write_number(3+i,c,row["Pipeline Net (QAR)"],fn);     c+=1
            ws.write(3+i,c,str(row["Pipeline Stage"]) or "",ft);     c+=1
            ws.write(3+i,c,str(row["Pipeline AM"]) or "",ft);        c+=1
            ws.write(3+i,c,str(row["Pipeline Quarter"]) or "",ft);   c+=1
            ws.write(3+i,c,str(row["Awarded Match"]) or "",ft);      c+=1
            ws.write_number(3+i,c,row["Awarded Score"],fs);           c+=1
            ws.write(3+i,c,str(row["Awarded Account"]) or "",ft);    c+=1
            ws.write_number(3+i,c,row["Awarded Gross (QAR)"],fn);    c+=1
            ws.write_number(3+i,c,row["Awarded Net (QAR)"],fn);      c+=1
            ws.write(3+i,c,str(row["Awarded Stage"]) or "",ft);      c+=1
            ws.write(3+i,c,str(row["Awarded AM"]) or "",ft);         c+=1
            ws.write(3+i,c,str(row["Awarded Year"]) or "",ft);       c+=1
    output.seek(0)
    return output.read()

st.set_page_config(
    page_title="Sales Dashboard",
    page_icon="📊",
    layout="wide",
)

# ── STAGE ORDER ────────────────────────────────────────────────────────────────
STAGE_ORDER = [
    "Stage 1: Assessment & Qualification",
    "Stage 2: Discovery & Scoping",
    "Stage 3.1: RFP & BID Qualification",
    "Stage 3.2: Solution Development & Proposal Submission",
    "Stage 4: Technical Evaluation By Customer",
    "Stage 5: Resolution/Financial Negotiation",
]

STAGE_SHORT = {
    "Stage 1: Assessment & Qualification":                    "S1 - Assessment",
    "Stage 2: Discovery & Scoping":                           "S2 - Discovery",
    "Stage 3.1: RFP & BID Qualification":                     "S3.1 - RFP",
    "Stage 3.2: Solution Development & Proposal Submission":  "S3.2 - Solution Dev",
    "Stage 4: Technical Evaluation By Customer":              "S4 - Tech Eval",
    "Stage 5: Resolution/Financial Negotiation":              "S5 - Negotiation",
}

STAGE_COLORS = {
    "S1 - Assessment":    "#B0C4DE",
    "S2 - Discovery":     "#6495ED",
    "S3.1 - RFP":         "#4169E1",
    "S3.2 - Solution Dev":"#DAA520",
    "S4 - Tech Eval":     "#FF8C00",
    "S5 - Negotiation":   "#228B22",
}

COA_FILE = "charter_of_accounts.xlsx"

# ── HELPERS ────────────────────────────────────────────────────────────────────
def fmt_m(val):
    if pd.isna(val): return "—"
    return f"QAR {val/1_000_000:.1f}M"

@st.cache_data
def load_coa():
    coa = pd.read_excel(COA_FILE)
    coa.columns = coa.columns.str.strip()
    coa["_code"] = coa["DU"].str.extract(r"(\d{6})")
    return coa.dropna(subset=["_code"]).set_index("_code")["BU"].to_dict()

def du_to_bu(du_str, mapping):
    m = re.match(r"(\d{6})", str(du_str).strip())
    return mapping.get(m.group(1), "Unknown") if m else "Unknown"

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

def _expand_deals(df_in, opp_col=None, mapping=None):  # noqa: ARG001
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
            nr["BU_exp"]    = du_to_bu(du, mapping) if mapping is not None else str(du)
            nr["DU_exp"]    = du
            nr["Gross_exp"] = _parse_num(gr[i])
            nr["Net_exp"]   = _parse_num(nt[i])
            nr["_is_first"] = (i == 0)
            nr["_du_count"] = n
            nr["_deal_idx"] = deal_idx
            rows.append(nr)
    return pd.DataFrame(rows)

@st.cache_data
def load_data(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    df["Stage_Short"] = df["Stage"].map(STAGE_SHORT).fillna(df["Stage"])
    df["Total Gross"] = pd.to_numeric(df["Total Gross"], errors="coerce").fillna(0)
    df["Total Net"]   = pd.to_numeric(df["Total Net"],   errors="coerce").fillna(0)
    df["Est. Close Date"] = pd.to_datetime(df["Est. Close Date"], errors="coerce")
    today = pd.Timestamp(date.today())
    df["Overdue"] = (df["Est. Close Date"] < today) & (~df["Stage"].str.contains("Won|Lost", na=False))
    return df

@st.cache_data
def build_du_breakdown(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    mapping = load_coa()
    rows = []
    for _, row in df.iterrows():
        dus   = str(row["DU"]).split("\n")   if pd.notna(row["DU"])               else ["Unknown"]
        gross = str(row["Gross (breakdown)"]).replace(",", "").split("\n") if pd.notna(row["Gross (breakdown)"]) else ["0"]
        net   = str(row["Net (breakdown)"]).replace(",", "").split("\n")   if pd.notna(row["Net (breakdown)"])   else ["0"]
        n = max(len(dus), len(gross), len(net))
        for i in range(n):
            du = dus[i].strip()   if i < len(dus)   else dus[-1].strip()
            g  = gross[i].strip() if i < len(gross) else "0"
            nt = net[i].strip()   if i < len(net)   else "0"
            try: g_val = float(g)
            except: g_val = 0.0
            try: n_val = float(nt)
            except: n_val = 0.0
            rows.append({
                "BU":                  du_to_bu(du, mapping),
                "DU":                  du,
                "Gross":               g_val,
                "Net":                 n_val,
                "Forecasted":          str(row.get("Forecasted", "")).strip(),
                "Account Manager":     row.get("Account Manager", ""),
                "Stage":               row.get("Stage", ""),
                "Sector":              row.get("Sector", ""),
                "Closure Due Quarter": row.get("Closure Due Quarter", ""),
                "Account Name":        row.get("Account Name", ""),
                "Lead/Opp Name":       row.get("Lead/Opp Name", ""),
                "Winning Probability": row.get("Winning Probability", ""),
                "Est. Close Date":     row.get("Est. Close Date", pd.NaT),
            })
    return pd.DataFrame(rows)

# ── LOAD COA ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_awarded(file):
    df = pd.read_excel(file, sheet_name="Export")
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    df["Total Gross"]   = pd.to_numeric(df["Total Gross"],   errors="coerce").fillna(0)
    df["Total Net"]     = pd.to_numeric(df["Total Net"],     errors="coerce").fillna(0)
    df["Project Value"] = pd.to_numeric(df["Project value (as per the contract value)"], errors="coerce").fillna(0)
    df["Client Commitment"] = pd.to_numeric(df["Client Commitment/WOs Net"], errors="coerce").fillna(0)
    # Simplify New/Renew (multi-line) to single value
    def simplify_nr(val):
        if pd.isna(val): return "Unknown"
        vals = set(v.strip() for v in str(val).split("\n"))
        if vals == {"New"}: return "New"
        if vals == {"Renew"}: return "Renew"
        return "Mixed"
    df["Type"] = df["New/Renew"].apply(simplify_nr)
    return df

@st.cache_data
def build_awarded_du(file):
    df = pd.read_excel(file, sheet_name="Export")
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    mapping = load_coa()
    rows = []
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
            rows.append({
                "BU":              du_to_bu(du, mapping),
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
    return pd.DataFrame(rows)

# ── EXCEL EXPORT HELPERS ────────────────────────────────────────────────────────
def _xl_formats(wb):
    """Return a dict of named xlsxwriter formats."""
    return {
        "title":    wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF","bg_color":"#1a3a6b","align":"center","valign":"vcenter"}),
        "header":   wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#1a3a6b","border":1,"align":"center","valign":"vcenter","text_wrap":True}),
        "kpi_lbl":  wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"}),
        "kpi_val":  wb.add_format({"bold":True,"bg_color":"#EBF0FB","border":1,"num_format":"#,##0","align":"right"}),
        "num":      wb.add_format({"num_format":"#,##0","border":1,"align":"right"}),
        "text":     wb.add_format({"border":1,"align":"left"}),
        "date":     wb.add_format({"num_format":"dd-mmm-yyyy","border":1,"align":"center"}),
        "alt":      wb.add_format({"bg_color":"#F2F5FB","border":1,"align":"left"}),
        "alt_num":  wb.add_format({"bg_color":"#F2F5FB","num_format":"#,##0","border":1,"align":"right"}),
        "red":      wb.add_format({"bg_color":"#FFE0E0","border":1,"align":"left"}),
        "red_num":  wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right"}),
        "red_dt":   wb.add_format({"bg_color":"#FFE0E0","num_format":"dd-mmm-yyyy","border":1,"align":"center"}),
        "grn":      wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"}),
        "grn_num":  wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"}),
        "grn_dt":   wb.add_format({"bg_color":"#E2EFDA","num_format":"dd-mmm-yyyy","border":1,"align":"center"}),
        "bu_lbl":   wb.add_format({"bold":True,"bg_color":"#D9E1F2","border":1,"align":"left"}),
        "bu_num":   wb.add_format({"bold":True,"bg_color":"#D9E1F2","num_format":"#,##0","border":1,"align":"right"}),
        "tot_lbl":  wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF","border":1,"align":"left"}),
        "tot_num":  wb.add_format({"bold":True,"bg_color":"#1a3a6b","font_color":"#FFFFFF","num_format":"#,##0","border":1,"align":"right"}),
        "sec_hdr":  wb.add_format({"bold":True,"bg_color":"#2E5FA3","font_color":"#FFFFFF","border":1,"align":"left","font_size":11}),
        "y25":      wb.add_format({"bg_color":"#DAE8FC","border":1,"align":"left"}),
        "y25_num":  wb.add_format({"bg_color":"#DAE8FC","num_format":"#,##0","border":1,"align":"right"}),
        "y26":      wb.add_format({"bg_color":"#D5E8D4","border":1,"align":"left"}),
        "y26_num":  wb.add_format({"bg_color":"#D5E8D4","num_format":"#,##0","border":1,"align":"right"}),
        "opp":      wb.add_format({"italic":True,"bg_color":"#FAFAFA","border":1,"align":"left","indent":2}),
        "opp_num":  wb.add_format({"italic":True,"bg_color":"#FAFAFA","num_format":"#,##0","border":1,"align":"right"}),
    }

def _hdr(ws, row, cols, widths, fmt):
    for c, col in enumerate(cols):
        ws.write(row, c, col, fmt)
    if widths:
        for c, w in enumerate(widths):
            ws.set_column(c, c, w)

@st.cache_data
def export_pipeline_excel(file, book3_file=None, awarded_file=None):
    TODAY = date.today()
    mapping = load_coa()
    df = load_data(file)
    du_exp = build_du_breakdown(file)

    STAGE_SHORT_MAP = {
        "Stage 1: Assessment & Qualification":                   "S1 - Assessment",
        "Stage 2: Discovery & Scoping":                          "S2 - Discovery",
        "Stage 3.1: RFP & BID Qualification":                    "S3.1 - RFP",
        "Stage 3.2: Solution Development & Proposal Submission": "S3.2 - Solution Dev",
        "Stage 4: Technical Evaluation By Customer":             "S4 - Tech Eval",
        "Stage 5: Resolution/Financial Negotiation":             "S5 - Negotiation",
    }
    stage_order = list(STAGE_SHORT_MAP.values())

    # ── Build tables ──────────────────────────────────────────────────────────
    kpi_df = pd.DataFrame([
        {"Metric":"Total Opportunities",         "Value": len(df)},
        {"Metric":"Total Gross Pipeline (QAR)",  "Value": df["Total Gross"].sum()},
        {"Metric":"Total Net Pipeline (QAR)",    "Value": df["Total Net"].sum()},
        {"Metric":"Forecasted Gross (QAR)",      "Value": df[df["Forecasted"]=="Yes"]["Total Gross"].sum()},
        {"Metric":"Forecasted Net (QAR)",        "Value": df[df["Forecasted"]=="Yes"]["Total Net"].sum()},
        {"Metric":"Strategic Opportunities",     "Value": len(df[df["Strategic Opportunity"]=="Yes"])},
        {"Metric":"Overdue Deals",               "Value": int(df["Overdue"].sum())},
    ])
    stage_df = (df.groupby("Stage_Short").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().rename(columns={"Stage_Short":"Stage"}))
    stage_df["_ord"] = stage_df["Stage"].map({s:i for i,s in enumerate(stage_order)})
    stage_df = stage_df.sort_values("_ord").drop(columns="_ord")
    sector_df = df.groupby("Sector").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Net",ascending=False)
    am_df = df.groupby("Account Manager").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Net",ascending=False)
    fore_am = df[df["Forecasted"]=="Yes"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
    am_df = am_df.merge(fore_am, on="Account Manager", how="left").fillna({"Forecasted Net":0})
    q_df = df.groupby("Closure Due Quarter").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Closure Due Quarter")
    fore_q = df[df["Forecasted"]=="Yes"].groupby("Closure Due Quarter")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
    q_df = q_df.merge(fore_q, on="Closure Due Quarter", how="left").fillna({"Forecasted Net":0})
    prob_df = df.groupby("Winning Probability").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Net",ascending=False)
    du_totals = du_exp.groupby(["BU","DU"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","Net"],ascending=[True,False])
    fore_du = du_exp[du_exp["Forecasted"]=="Yes"].groupby("DU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"})
    du_totals = du_totals.merge(fore_du, on="DU", how="left").fillna({"Forecasted Net":0})
    fore_du_detail = du_exp[du_exp["Forecasted"]=="Yes"].copy().sort_values(["BU","DU","Net"],ascending=[True,True,False])
    fore_du_summary = fore_du_detail.groupby(["BU","DU","Closure Due Quarter"]).agg(Count=("Lead/Opp Name","count"),Gross=("Gross","sum"),Net=("Net","sum")).reset_index().sort_values(["BU","DU","Closure Due Quarter"])
    fore_df = df[df["Forecasted"]=="Yes"][["Account Name","Lead/Opp Name","Stage_Short","Account Manager","Sector","Total Gross","Total Net","Winning Probability","Closure Due Quarter","Est. Close Date"]].sort_values("Total Net",ascending=False).rename(columns={"Stage_Short":"Stage"})
    overdue_df = df[df["Overdue"]][["Account Name","Lead/Opp Name","Stage_Short","Account Manager","Total Net","Est. Close Date","Winning Probability","Closure Due Quarter"]].sort_values("Est. Close Date").rename(columns={"Stage_Short":"Stage"})
    full_df = df[["SNo.","Account Name","Lead/Opp Name","Stage_Short","Account Manager","Sector","BU","DU","Total Gross","Total Net","Winning Probability","Forecasted","Strategic Opportunity","Closure Due Quarter","Est. Close Date","Source of Opportunity","Overdue"]].sort_values("Total Net",ascending=False).rename(columns={"Stage_Short":"Stage"})
    du_stage = du_exp.groupby(["BU","DU","Stage"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","DU","Net"],ascending=[True,True,False])
    fore_du_stage = du_exp[du_exp["Forecasted"]=="Yes"].groupby(["DU","Stage"])["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"})
    du_stage = du_stage.merge(fore_du_stage, on=["DU","Stage"], how="left").fillna({"Forecasted Net":0})

    # ── Write to buffer ───────────────────────────────────────────────────────
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        F = _xl_formats(wb)

        # SHEET 1 — SUMMARY
        ws = wb.add_worksheet("Summary"); ws.set_zoom(90); ws.set_tab_color("#1a3a6b")
        writer.sheets["Summary"] = ws
        ws.merge_range("A1:C1", f"Weekly Pipeline Summary — {TODAY.strftime('%d %B %Y')}", F["title"]); ws.set_row(0,28)
        _hdr(ws, 2, ["Metric","Value"], [34,22], F["header"])
        for i, row in kpi_df.iterrows():
            ws.write(3+i, 0, row["Metric"], F["kpi_lbl"]); ws.write_number(3+i, 1, row["Value"], F["kpi_val"])
        ws.merge_range("D1:H1", "Pipeline by Stage", F["title"])
        _hdr(ws, 2, ["Stage","Count","Gross (QAR)","Net (QAR)"], None, F["header"])
        ws.set_column(3,3,22); ws.set_column(4,4,8); ws.set_column(5,5,18); ws.set_column(6,6,18)
        for i, row in stage_df.reset_index(drop=True).iterrows():
            alt = i%2==1
            ws.write(3+i,3,row["Stage"],F["alt"] if alt else F["text"]); ws.write_number(3+i,4,row["Count"],F["alt_num"] if alt else F["num"]); ws.write_number(3+i,5,row["Gross"],F["alt_num"] if alt else F["num"]); ws.write_number(3+i,6,row["Net"],F["alt_num"] if alt else F["num"])

        # SHEET 2 — DU BREAKDOWN
        du_ws = wb.add_worksheet("DU Breakdown"); du_ws.set_zoom(90); du_ws.set_tab_color("#FF8C00")
        writer.sheets["DU Breakdown"] = du_ws
        du_ws.merge_range("A1:G1","Gross & Net Breakdown by BU / Delivery Unit",F["title"]); du_ws.set_row(0,28)
        _hdr(du_ws,1,["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[42,52,20,20,22],F["header"])
        r_out=0
        for bu_name, bu_grp in du_totals.groupby("BU"):
            du_ws.write(2+r_out,0,bu_name,F["bu_lbl"]); du_ws.write(2+r_out,1,"",F["bu_lbl"])
            du_ws.write_number(2+r_out,2,bu_grp["Gross"].sum(),F["bu_num"]); du_ws.write_number(2+r_out,3,bu_grp["Net"].sum(),F["bu_num"]); du_ws.write_number(2+r_out,4,bu_grp["Forecasted Net"].sum(),F["bu_num"]); r_out+=1
            for _, row in bu_grp.iterrows():
                alt=r_out%2==1
                du_ws.write(2+r_out,0,"",F["alt"] if alt else F["text"]); du_ws.write(2+r_out,1,row["DU"],F["alt"] if alt else F["text"])
                du_ws.write_number(2+r_out,2,row["Gross"],F["alt_num"] if alt else F["num"]); du_ws.write_number(2+r_out,3,row["Net"],F["alt_num"] if alt else F["num"]); du_ws.write_number(2+r_out,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"]); r_out+=1
                for _, deal in du_exp[du_exp["DU"]==row["DU"]].iterrows():
                    fore_net=deal["Net"] if str(deal.get("Forecasted","")).strip()=="Yes" else 0
                    du_ws.write(2+r_out,0,"",F["opp"]); du_ws.write(2+r_out,1,f"  ↳  {deal['Lead/Opp Name']}",F["opp"])
                    du_ws.write_number(2+r_out,2,deal["Gross"],F["opp_num"]); du_ws.write_number(2+r_out,3,deal["Net"],F["opp_num"]); du_ws.write_number(2+r_out,4,fore_net,F["opp_num"]); r_out+=1
        t=2+r_out
        du_ws.write(t,0,"GRAND TOTAL",F["tot_lbl"]); du_ws.write(t,1,"",F["tot_lbl"])
        du_ws.write_number(t,2,du_totals["Gross"].sum(),F["tot_num"]); du_ws.write_number(t,3,du_totals["Net"].sum(),F["tot_num"]); du_ws.write_number(t,4,du_totals["Forecasted Net"].sum(),F["tot_num"])
        du_ws.merge_range(t+2,0,t+2,5,"BU / DU x Stage Detail",F["title"])
        _hdr(du_ws,t+3,["BU","Delivery Unit","Stage","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],None,F["header"])
        for i, row in du_stage.reset_index(drop=True).iterrows():
            alt=i%2==1
            du_ws.write(t+4+i,0,row["BU"],F["alt"] if alt else F["text"]); du_ws.write(t+4+i,1,row["DU"],F["alt"] if alt else F["text"]); du_ws.write(t+4+i,2,row["Stage"],F["alt"] if alt else F["text"])
            du_ws.write_number(t+4+i,3,row["Gross"],F["alt_num"] if alt else F["num"]); du_ws.write_number(t+4+i,4,row["Net"],F["alt_num"] if alt else F["num"]); du_ws.write_number(t+4+i,5,row["Forecasted Net"],F["alt_num"] if alt else F["num"])

        # SHEET 3 — FORECAST PER DU
        fd_ws = wb.add_worksheet("Forecast per DU"); fd_ws.set_zoom(90); fd_ws.set_tab_color("#228B22")
        writer.sheets["Forecast per DU"] = fd_ws
        fd_ws.merge_range("A1:J1",f"Forecasted Pipeline by BU / Delivery Unit — {TODAY.strftime('%d %B %Y')}",F["title"]); fd_ws.set_row(0,28)
        fd_ws.merge_range("A2:J2","Summary: Forecasted Net by BU / DU / Quarter",F["sec_hdr"]); fd_ws.set_row(2,22)
        _hdr(fd_ws,3,["BU","Delivery Unit","Quarter","Count","Gross (QAR)","Net (QAR)"],[42,38,10,8,20,20],F["header"])
        current_bu=None; r_s=0
        for _, row in fore_du_summary.iterrows():
            if row["BU"]!=current_bu:
                current_bu=row["BU"]; bu_sub=fore_du_summary[fore_du_summary["BU"]==current_bu]
                fd_ws.write(4+r_s,0,current_bu,F["bu_lbl"]); fd_ws.write(4+r_s,1,f"Total: {len(bu_sub)} rows",F["bu_lbl"]); fd_ws.write(4+r_s,2,"",F["bu_lbl"])
                fd_ws.write_number(4+r_s,3,bu_sub["Count"].sum(),F["bu_num"]); fd_ws.write_number(4+r_s,4,bu_sub["Gross"].sum(),F["bu_num"]); fd_ws.write_number(4+r_s,5,bu_sub["Net"].sum(),F["bu_num"]); r_s+=1
            alt=r_s%2==1
            fd_ws.write(4+r_s,0,"",F["alt"] if alt else F["text"]); fd_ws.write(4+r_s,1,row["DU"],F["alt"] if alt else F["text"]); fd_ws.write(4+r_s,2,row["Closure Due Quarter"],F["alt"] if alt else F["text"])
            fd_ws.write_number(4+r_s,3,row["Count"],F["alt_num"] if alt else F["num"]); fd_ws.write_number(4+r_s,4,row["Gross"],F["alt_num"] if alt else F["num"]); fd_ws.write_number(4+r_s,5,row["Net"],F["alt_num"] if alt else F["num"]); r_s+=1
        ts=4+r_s
        fd_ws.write(ts,0,"GRAND TOTAL",F["tot_lbl"])
        for _c in range(1,3): fd_ws.write(ts,_c,"",F["tot_lbl"])
        fd_ws.write_number(ts,3,fore_du_summary["Count"].sum(),F["tot_num"]); fd_ws.write_number(ts,4,fore_du_summary["Gross"].sum(),F["tot_num"]); fd_ws.write_number(ts,5,fore_du_summary["Net"].sum(),F["tot_num"])
        det_start=ts+2
        fd_ws.merge_range(det_start,0,det_start,9,"Detail: Forecasted Deals by BU / DU",F["sec_hdr"]); fd_ws.set_row(det_start,22)
        _hdr(fd_ws,det_start+1,["BU","Delivery Unit","Account Name","Opportunity","Stage","Account Manager","Quarter","Gross (QAR)","Net (QAR)","Win Prob"],[42,38,28,36,22,24,10,18,18,12],F["header"])
        current_bu=None; r_d=0
        for _, row in fore_du_detail.iterrows():
            if row["BU"]!=current_bu:
                current_bu=row["BU"]; bu_grp=fore_du_detail[fore_du_detail["BU"]==current_bu]
                fd_ws.write(det_start+2+r_d,0,current_bu,F["bu_lbl"])
                for _cc in range(1,10): fd_ws.write(det_start+2+r_d,_cc,"",F["bu_lbl"])
                fd_ws.write_number(det_start+2+r_d,7,bu_grp["Gross"].sum(),F["bu_num"]); fd_ws.write_number(det_start+2+r_d,8,bu_grp["Net"].sum(),F["bu_num"]); r_d+=1
            alt=r_d%2==1; f_t=F["grn"] if not alt else F["alt"]; f_n=F["grn_num"] if not alt else F["alt_num"]
            fd_ws.write(det_start+2+r_d,0,"",f_t); fd_ws.write(det_start+2+r_d,1,str(row["DU"]) if pd.notna(row["DU"]) else "",f_t); fd_ws.write(det_start+2+r_d,2,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",f_t); fd_ws.write(det_start+2+r_d,3,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",f_t); fd_ws.write(det_start+2+r_d,4,str(row["Stage"]) if pd.notna(row["Stage"]) else "",f_t); fd_ws.write(det_start+2+r_d,5,str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "",f_t); fd_ws.write(det_start+2+r_d,6,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",f_t)
            fd_ws.write_number(det_start+2+r_d,7,row["Gross"],f_n); fd_ws.write_number(det_start+2+r_d,8,row["Net"],f_n); fd_ws.write(det_start+2+r_d,9,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",f_t); r_d+=1
        td=det_start+2+r_d
        fd_ws.write(td,0,"GRAND TOTAL",F["tot_lbl"])
        for _c in range(1,7): fd_ws.write(td,_c,"",F["tot_lbl"])
        fd_ws.write_number(td,7,fore_du_detail["Gross"].sum(),F["tot_num"]); fd_ws.write_number(td,8,fore_du_detail["Net"].sum(),F["tot_num"]); fd_ws.write(td,9,"",F["tot_lbl"])

        # SHEET 4 — SECTOR & AM
        sa_ws = wb.add_worksheet("Sector & AM"); sa_ws.set_zoom(90); sa_ws.set_tab_color("#228B22")
        writer.sheets["Sector & AM"] = sa_ws
        sa_ws.merge_range("A1:E1","Pipeline by Sector",F["title"]); sa_ws.set_row(0,28)
        _hdr(sa_ws,1,["Sector","Count","Gross (QAR)","Net (QAR)"],[24,8,20,20],F["header"])
        for i, row in sector_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            sa_ws.write(2+i,0,row["Sector"],F["alt"] if alt else F["text"]); sa_ws.write_number(2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(2+i,3,row["Net"],F["alt_num"] if alt else F["num"])
        off=2+len(sector_df)+2
        sa_ws.merge_range(off,0,off,4,"Pipeline by Account Manager",F["title"])
        _hdr(sa_ws,off+1,["Account Manager","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[28,8,20,20,22],F["header"])
        for i, row in am_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            sa_ws.write(off+2+i,0,row["Account Manager"],F["alt"] if alt else F["text"]); sa_ws.write_number(off+2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,3,row["Net"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])

        # SHEET 5 — QUARTERLY & PROBABILITY
        qp_ws = wb.add_worksheet("Quarterly & Probability"); qp_ws.set_zoom(90); qp_ws.set_tab_color("#DAA520")
        writer.sheets["Quarterly & Probability"] = qp_ws
        qp_ws.merge_range("A1:E1","Quarterly Close Plan",F["title"]); qp_ws.set_row(0,28)
        _hdr(qp_ws,1,["Quarter","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[12,8,20,20,22],F["header"])
        for i, row in q_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            qp_ws.write(2+i,0,row["Closure Due Quarter"],F["alt"] if alt else F["text"]); qp_ws.write_number(2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,3,row["Net"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])
        off2=2+len(q_df)+2
        qp_ws.merge_range(off2,0,off2,3,"Pipeline by Winning Probability",F["title"])
        _hdr(qp_ws,off2+1,["Winning Probability","Count","Gross (QAR)","Net (QAR)"],[22,8,20,20],F["header"])
        for i, row in prob_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            qp_ws.write(off2+2+i,0,row["Winning Probability"],F["alt"] if alt else F["text"]); qp_ws.write_number(off2+2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(off2+2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(off2+2+i,3,row["Net"],F["alt_num"] if alt else F["num"])

        # SHEET 6 — FORECAST
        fw = wb.add_worksheet("Forecast"); fw.set_zoom(90); fw.set_tab_color("#228B22")
        writer.sheets["Forecast"] = fw
        fw.merge_range("A1:J1",f"Forecasted Deals — {TODAY.strftime('%d %B %Y')}",F["title"]); fw.set_row(0,28)
        _hdr(fw,1,["Account Name","Opportunity","Stage","Account Manager","Sector","Gross (QAR)","Net (QAR)","Win Probability","Quarter","Est. Close Date"],[30,36,22,24,18,18,18,14,10,16],F["header"])
        for i, row in fore_df.reset_index(drop=True).iterrows():
            fw.write(2+i,0,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",F["grn"]); fw.write(2+i,1,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",F["grn"]); fw.write(2+i,2,str(row["Stage"]) if pd.notna(row["Stage"]) else "",F["grn"]); fw.write(2+i,3,str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "",F["grn"]); fw.write(2+i,4,str(row["Sector"]) if pd.notna(row["Sector"]) else "",F["grn"])
            fw.write_number(2+i,5,row["Total Gross"],F["grn_num"]); fw.write_number(2+i,6,row["Total Net"],F["grn_num"])
            fw.write(2+i,7,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",F["grn"]); fw.write(2+i,8,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",F["grn"])
            if pd.notna(row["Est. Close Date"]):
                fw.write_datetime(2+i,9,row["Est. Close Date"].to_pydatetime(),F["grn_dt"])
            else:
                fw.write_blank(2+i,9,None,F["grn_dt"])
        t2=2+len(fore_df)
        fw.write(t2,0,"TOTAL",F["tot_lbl"]); fw.write_number(t2,5,fore_df["Total Gross"].sum(),F["tot_num"]); fw.write_number(t2,6,fore_df["Total Net"].sum(),F["tot_num"])

        # SHEET 7 — OVERDUE
        ow = wb.add_worksheet("Overdue Deals"); ow.set_zoom(90); ow.set_tab_color("#FF0000")
        writer.sheets["Overdue Deals"] = ow
        ow.merge_range("A1:H1",f"Overdue Deals — {TODAY.strftime('%d %B %Y')}",F["title"]); ow.set_row(0,28)
        _hdr(ow,1,["Account Name","Opportunity","Stage","Account Manager","Net (QAR)","Est. Close Date","Win Probability","Quarter"],[30,36,22,24,18,16,14,10],F["header"])
        for i, row in overdue_df.reset_index(drop=True).iterrows():
            ow.write(2+i,0,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",F["red"]); ow.write(2+i,1,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",F["red"]); ow.write(2+i,2,str(row["Stage"]) if pd.notna(row["Stage"]) else "",F["red"]); ow.write(2+i,3,str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "",F["red"])
            ow.write_number(2+i,4,row["Total Net"],F["red_num"])
            if pd.notna(row["Est. Close Date"]):
                ow.write_datetime(2+i,5,row["Est. Close Date"].to_pydatetime(),F["red_dt"])
            else:
                ow.write_blank(2+i,5,None,F["red_dt"])
            ow.write(2+i,6,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",F["red"]); ow.write(2+i,7,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",F["red"])

        # SHEET 8 — FULL PIPELINE
        pw = wb.add_worksheet("Full Pipeline"); pw.set_zoom(85); pw.set_tab_color("#6495ED")
        writer.sheets["Full Pipeline"] = pw
        pw.merge_range("A1:Q1","Full Pipeline — All Opportunities",F["title"]); pw.set_row(0,28); pw.freeze_panes(2,0)
        _hdr(pw,1,["#","Account Name","Opportunity","Stage","Account Manager","Sector","BU","DU","Gross (QAR)","Net (QAR)","Win Prob","Forecasted","Strategic","Quarter","Est. Close Date","Source","Overdue"],[5,28,36,22,22,16,30,36,18,18,12,12,10,10,16,16,8],F["header"])
        pw.autofilter(1,0,1+len(full_df),16)
        for i, row in full_df.reset_index(drop=True).iterrows():
            alt=i%2==1; is_ov=bool(row["Overdue"])
            ft=F["red"] if is_ov else (F["alt"] if alt else F["text"]); fn=F["red_num"] if is_ov else (F["alt_num"] if alt else F["num"]); fd_=F["red_dt"] if is_ov else F["date"]
            pw.write(2+i,0,str(row["SNo."]) if pd.notna(row["SNo."]) else "",ft); pw.write(2+i,1,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",ft); pw.write(2+i,2,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",ft); pw.write(2+i,3,str(row["Stage"]) if pd.notna(row["Stage"]) else "",ft); pw.write(2+i,4,str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "",ft); pw.write(2+i,5,str(row["Sector"]) if pd.notna(row["Sector"]) else "",ft); pw.write(2+i,6,str(row["BU"]) if pd.notna(row["BU"]) else "",ft); pw.write(2+i,7,str(row["DU"]) if pd.notna(row["DU"]) else "",ft)
            pw.write_number(2+i,8,row["Total Gross"],fn); pw.write_number(2+i,9,row["Total Net"],fn)
            pw.write(2+i,10,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",ft); pw.write(2+i,11,str(row["Forecasted"]) if pd.notna(row["Forecasted"]) else "",ft); pw.write(2+i,12,str(row["Strategic Opportunity"]) if pd.notna(row["Strategic Opportunity"]) else "",ft); pw.write(2+i,13,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",ft)
            if pd.notna(row["Est. Close Date"]):
                pw.write_datetime(2+i,14,row["Est. Close Date"].to_pydatetime(),fd_)
            else:
                pw.write_blank(2+i,14,None,fd_)
            pw.write(2+i,15,str(row["Source of Opportunity"]) if pd.notna(row["Source of Opportunity"]) else "",ft); pw.write(2+i,16,"YES" if is_ov else "",ft)

        # SHEET 9 — PIPELINE BREAKDOWN (styled, one row per DU per opportunity)
        _CLR = {
            "title_bg":"1F3864","hdr_deal":"1F3864","hdr_du":"17375E",
            "hdr_finance":"1F4E79","hdr_other":"2E5FA3",
            "bu_fill":"EDF2F9","du_fill":"E4ECF7","num_fill":"EBF5FB",
            "tot_fill":"D5E8F5","date_fill":"FFF2CC",
            "alt_a":"F5F8FF","alt_b":"FFFFFF",
        }
        def _bdf(bg, top, num_fmt=None):
            d={"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
            if num_fmt: d["num_format"]=num_fmt; d["align"]="right"
            return wb.add_format(d)
        bd_fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_CLR["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        bd_fh_du      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_CLR["hdr_du"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        bd_fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_CLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        bd_fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_CLR["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        bd_fmt_title  = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+_CLR["title_bg"],"align":"center","valign":"vcenter"})
        bd_ft_a1=_bdf(_CLR["alt_a"],2);  bd_ft_an=_bdf(_CLR["alt_a"],1)
        bd_ft_b1=_bdf(_CLR["alt_b"],2);  bd_ft_bn=_bdf(_CLR["alt_b"],1)
        bd_fn_a1=_bdf(_CLR["alt_a"],2,"#,##0"); bd_fn_an=_bdf(_CLR["alt_a"],1,"#,##0")
        bd_fn_b1=_bdf(_CLR["alt_b"],2,"#,##0"); bd_fn_bn=_bdf(_CLR["alt_b"],1,"#,##0")
        bd_fbu1=_bdf(_CLR["bu_fill"],2); bd_fbun=_bdf(_CLR["bu_fill"],1)
        bd_fdu1=_bdf(_CLR["du_fill"],2); bd_fdun=_bdf(_CLR["du_fill"],1)
        bd_fxn1=_bdf(_CLR["num_fill"],2,"#,##0"); bd_fxnn=_bdf(_CLR["num_fill"],1,"#,##0")
        bd_ftot  =wb.add_format({"bg_color":"#"+_CLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
        bd_ftotbl=wb.add_format({"bg_color":"#"+_CLR["num_fill"],"border":1,"font_size":9})
        bd_fdate =wb.add_format({"bg_color":"#"+_CLR["date_fill"],"border":1,"align":"center","num_format":"DD-MMM-YYYY","font_size":9})
        bd_fdatebl=wb.add_format({"bg_color":"#"+_CLR["alt_a"],"border":1,"font_size":9})

        bw = wb.add_worksheet("Pipeline Breakdown"); bw.set_zoom(85); bw.set_tab_color("#9370DB")
        writer.sheets["Pipeline Breakdown"] = bw
        bd_ocols = [
            ("SNo.",6,"deal"),("Account Name",24,"deal"),("Lead/Opp Name",36,"deal"),
            ("BU",36,"du"),("DU",34,"du"),
            ("Gross (breakdown)",16,"finance"),("Net (breakdown)",16,"finance"),
            ("Total Gross",15,"finance"),("Total Net",15,"finance"),
            ("Stage",36,"other"),("Account Manager",22,"other"),("Sector",16,"other"),
            ("Closure Due Quarter",9,"other"),("Winning Probability",10,"other"),
            ("Forecasted",10,"other"),("Strategic Opportunity",10,"other"),
            ("Est. Close Date",14,"other"),
        ]
        bd_hfmt={"deal":bd_fh_deal,"du":bd_fh_du,"finance":bd_fh_finance,"other":bd_fh_other}
        bd_nc=len(bd_ocols)
        bw.merge_range(0,0,0,bd_nc-1,"Pipeline — Expanded by Delivery Unit",bd_fmt_title); bw.set_row(0,28)
        for c,(cn,cw,ct) in enumerate(bd_ocols): bw.write(1,c,cn,bd_hfmt[ct]); bw.set_column(c,c,cw)
        bw.set_row(1,28); bw.freeze_panes(2,0)
        bd_exp = _expand_deals(df, mapping=mapping)
        bw.autofilter(1,0,1+len(bd_exp),bd_nc-1)
        bd_cmap={cn:idx for idx,(cn,_,__) in enumerate(bd_ocols)}
        bd_prev=None; bd_alt=False
        for rp,(_, row) in enumerate(bd_exp.iterrows()):
            didx=row["_deal_idx"]; isf=bool(row["_is_first"])
            if didx!=bd_prev: bd_alt=not bd_alt; bd_prev=didx
            ft=(bd_ft_a1 if isf else bd_ft_an) if bd_alt else (bd_ft_b1 if isf else bd_ft_bn)
            fn=(bd_fn_a1 if isf else bd_fn_an) if bd_alt else (bd_fn_b1 if isf else bd_fn_bn)
            fbu=bd_fbu1 if isf else bd_fbun; fdu=bd_fdu1 if isf else bd_fdun
            fxn=bd_fxn1 if isf else bd_fxnn; xl_r=2+rp
            def _bws(ci,val,fmt):
                if val is None or (isinstance(val,float) and pd.isna(val)): bw.write_blank(xl_r,ci,None,fmt)
                else: bw.write(xl_r,ci,str(val),fmt)
            def _bwn(ci,val,fmt):
                v=_parse_num(val) if not isinstance(val,(int,float)) else val
                if v is None or (isinstance(v,float) and pd.isna(v)): bw.write_blank(xl_r,ci,None,fmt)
                else: bw.write_number(xl_r,ci,v,fmt)
            _bws(bd_cmap["SNo."],              row.get("SNo.")               if isf else None,ft)
            _bws(bd_cmap["Account Name"],       row.get("Account Name")       if isf else None,ft)
            _bws(bd_cmap["Lead/Opp Name"],      row.get("Lead/Opp Name")      if isf else None,ft)
            _bws(bd_cmap["BU"],row.get("BU_exp"),fbu); _bws(bd_cmap["DU"],row.get("DU_exp"),fdu)
            _bwn(bd_cmap["Gross (breakdown)"],  row.get("Gross_exp"),fxn)
            _bwn(bd_cmap["Net (breakdown)"],    row.get("Net_exp"),  fxn)
            if isf:
                _bwn(bd_cmap["Total Gross"],row.get("Total Gross"),bd_ftot)
                _bwn(bd_cmap["Total Net"],  row.get("Total Net"),  bd_ftot)
            else:
                bw.write_blank(xl_r,bd_cmap["Total Gross"],None,bd_ftotbl)
                bw.write_blank(xl_r,bd_cmap["Total Net"],  None,bd_ftotbl)
            _bws(bd_cmap["Stage"],              row.get("Stage")              if isf else None,ft)
            _bws(bd_cmap["Account Manager"],    row.get("Account Manager")    if isf else None,ft)
            _bws(bd_cmap["Sector"],             row.get("Sector")             if isf else None,ft)
            _bws(bd_cmap["Closure Due Quarter"],row.get("Closure Due Quarter")if isf else None,ft)
            _bws(bd_cmap["Winning Probability"],row.get("Winning Probability")if isf else None,ft)
            _bws(bd_cmap["Forecasted"],         row.get("Forecasted")         if isf else None,ft)
            _bws(bd_cmap["Strategic Opportunity"],row.get("Strategic Opportunity") if isf else None,ft)
            cd_i=bd_cmap["Est. Close Date"]
            if isf:
                cd_v=row.get("Est. Close Date")
                if pd.notna(cd_v):
                    try: bw.write_datetime(xl_r,cd_i,pd.Timestamp(cd_v).to_pydatetime(),bd_fdate)
                    except Exception: bw.write(xl_r,cd_i,str(cd_v),bd_fdate)
                else: bw.write_blank(xl_r,cd_i,None,bd_fdate)
            else: bw.write_blank(xl_r,cd_i,None,bd_fdatebl)

        # SHEET 10 — BOOK3 MAPPING (only if Book3 uploaded)
        if book3_file is not None:
            b3_df   = load_book3(book3_file)
            aw_raw2 = load_awarded(awarded_file) if awarded_file is not None else pd.DataFrame()
            map_df  = build_book3_mapping(b3_df, df, aw_raw2)

            fmt_matched   = wb.add_format({"bg_color":"#E2EFDA","border":1,"align":"left"})
            fmt_matched_n = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right"})
            fmt_partial   = wb.add_format({"bg_color":"#FFF2CC","border":1,"align":"left"})
            fmt_partial_n = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right"})
            fmt_nomatch   = wb.add_format({"bg_color":"#FFE0E0","border":1,"align":"left"})
            fmt_nomatch_n = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right"})
            fmt_neg_grn   = wb.add_format({"bg_color":"#E2EFDA","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
            fmt_neg_yel   = wb.add_format({"bg_color":"#FFF2CC","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
            fmt_neg_red   = wb.add_format({"bg_color":"#FFE0E0","num_format":"#,##0","border":1,"align":"right","font_color":"#CC0000"})
            fmt_mhdr      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E5FA3","border":1,"align":"center","text_wrap":True})

            mw = wb.add_worksheet("Book3 Mapping"); mw.set_zoom(80); mw.set_tab_color("#FF8C00")
            writer.sheets["Book3 Mapping"] = mw
            mw.freeze_panes(3, 4)
            total_cols = 4 + len(BOOK3_MONTHS) + 12
            mw.merge_range(0,0,0,total_cols-1,f"Book3 ↔ Pipeline & Awarded Mapping — {TODAY.strftime('%d %B %Y')}",F["title"]); mw.set_row(0,28)
            mw.write(1,0,"🟢 Strong match (≥0.70)",fmt_matched); mw.write(1,1,"🟡 Partial match (0.55–0.69)",fmt_partial); mw.write(1,2,"🔴 No match (<0.55)",fmt_nomatch); mw.write(1,3,"",F["text"]); mw.set_row(1,18)
            all_hdr = ["Book3 BU","Book3 Project Type","Book3 Project Name","Book3 Grand Total (QAR)"] + BOOK3_MONTHS + ["Pipeline Match","Score","Pipeline Gross","Pipeline Net","Pipeline Stage","Pipeline AM","Awarded Match","Score","Awarded Gross","Awarded Net","Awarded Stage","Awarded AM"]
            all_wid = [36,22,42,20]+[10]*len(BOOK3_MONTHS)+[36,8,18,18,22,22,36,8,18,18,22,22]
            for c,(h,w) in enumerate(zip(all_hdr,all_wid)):
                mw.write(2,c,h,fmt_mhdr if h in BOOK3_MONTHS else F["header"]); mw.set_column(c,c,w)
            mw.set_row(2,20); mw.autofilter(2,0,2+len(map_df),len(all_hdr)-1)

            for i, row in map_df.reset_index(drop=True).iterrows():
                sc = max(row["Pipeline Score"], row["Awarded Score"])
                if sc >= 0.70:   ft,fn,fn_neg = fmt_matched,fmt_matched_n,fmt_neg_grn
                elif sc >= 0.55: ft,fn,fn_neg = fmt_partial,fmt_partial_n,fmt_neg_yel
                else:            ft,fn,fn_neg = fmt_nomatch,fmt_nomatch_n,fmt_neg_red
                c=0
                mw.write(3+i,c,str(row["Book3 BU"]) or "",ft);          c+=1
                mw.write(3+i,c,str(row["Book3 Project Type"]) or "",ft); c+=1
                mw.write(3+i,c,str(row["Book3 Project Name"]) or "",ft); c+=1
                mw.write_number(3+i,c,row["Book3 Grand Total"],fn_neg);  c+=1
                for m in BOOK3_MONTHS:
                    mw.write_number(3+i,c,row.get(f"Book3 {m}",0),fn_neg); c+=1
                mw.write(3+i,c,str(row["Pipeline Match"]) or "",ft);     c+=1
                mw.write_number(3+i,c,row["Pipeline Score"],fn);          c+=1
                mw.write_number(3+i,c,row["Pipeline Gross (QAR)"],fn);   c+=1
                mw.write_number(3+i,c,row["Pipeline Net (QAR)"],fn);     c+=1
                mw.write(3+i,c,str(row["Pipeline Stage"]) or "",ft);     c+=1
                mw.write(3+i,c,str(row["Pipeline AM"]) or "",ft);        c+=1
                mw.write(3+i,c,str(row["Awarded Match"]) or "",ft);      c+=1
                mw.write_number(3+i,c,row["Awarded Score"],fn);           c+=1
                mw.write_number(3+i,c,row["Awarded Gross (QAR)"],fn);    c+=1
                mw.write_number(3+i,c,row["Awarded Net (QAR)"],fn);      c+=1
                mw.write(3+i,c,str(row["Awarded Stage"]) or "",ft);      c+=1
                mw.write(3+i,c,str(row["Awarded AM"]) or "",ft);         c+=1

    output.seek(0)
    return output.read()


@st.cache_data
def export_awarded_excel(file26, file25):
    TODAY = date.today()
    parts_raw = []
    if file26 is not None:
        r = load_awarded(file26); r["Year"] = "2026"; parts_raw.append(r)
    if file25 is not None:
        r = load_awarded(file25); r["Year"] = "2025"; parts_raw.append(r)
    df = pd.concat(parts_raw, ignore_index=True)

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
            du_rows.append({"Year":row["Year"],"BU":du_to_bu(du, load_coa()),"DU":du,"Gross":g_val,"Net":n_val,
                            "Account Manager":row.get("Account Manager",""),"Stage":row.get("Stage",""),
                            "Award Quarter":row.get("Award Quarter",""),"Contracted":str(row.get("Contracted","")).strip(),
                            "Account Name":row.get("Account Name",""),"Opportunity":row.get("Opportunity Name","")})
    du_exp = pd.DataFrame(du_rows)

    year_df = df.groupby("Year").agg(Count=("Opportunity Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum"),PV=("Project Value","sum")).reset_index().sort_values("Year")
    stage_df = df.groupby(["Year","Stage"]).agg(Count=("Opportunity Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values(["Year","Stage"])
    am_df = df.groupby("Account Manager").agg(Count=("Opportunity Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Net",ascending=False)
    am_25 = df[df["Year"]=="2025"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Net 2025"})
    am_26 = df[df["Year"]=="2026"].groupby("Account Manager")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Net 2026"})
    am_df = am_df.merge(am_25,on="Account Manager",how="left").fillna({"Net 2025":0}).merge(am_26,on="Account Manager",how="left").fillna({"Net 2026":0})
    q_df = df.groupby(["Year","Award Quarter"]).agg(Count=("Opportunity Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values(["Year","Award Quarter"])
    nr_df = df.groupby(["Year","Type"]).agg(Count=("Opportunity Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values(["Year","Type"])
    du_totals = du_exp.groupby(["BU","DU"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","Net"],ascending=[True,False])
    du_year_pivot = du_exp.groupby(["BU","DU","Year"])[["Net"]].sum().reset_index().pivot_table(index=["BU","DU"],columns="Year",values="Net",aggfunc="sum",fill_value=0).reset_index()
    du_totals = du_totals.merge(du_year_pivot,on=["BU","DU"],how="left")
    for yr in ["2025","2026"]:
        if yr not in du_totals.columns: du_totals[yr] = 0
    kpi_df = pd.DataFrame([
        {"Metric":"Total Awarded Deals",        "Value":len(df)},
        {"Metric":"  — 2025",                   "Value":len(df[df["Year"]=="2025"])},
        {"Metric":"  — 2026",                   "Value":len(df[df["Year"]=="2026"])},
        {"Metric":"Total Gross (QAR)",          "Value":df["Total Gross"].sum()},
        {"Metric":"Total Net (QAR)",            "Value":df["Total Net"].sum()},
        {"Metric":"Total Contract Value (QAR)", "Value":df["Project Value"].sum()},
        {"Metric":"Contracted (Signed)",        "Value":len(df[df["Stage"].str.contains("Contracting",na=False)])},
        {"Metric":"LOA (Not Yet Signed)",       "Value":len(df[df["Stage"].str.contains("Letter Of Award",na=False)])},
        {"Metric":"New Deals",                  "Value":len(df[df["Type"]=="New"])},
        {"Metric":"Renew Deals",                "Value":len(df[df["Type"]=="Renew"])},
    ])
    full_df = df[["Year","SNo.","Account Name","Opportunity Name","Stage","Account Manager","Type","Total Gross","Total Net","Project Value","Award Quarter","Contracted","Contract Signed Quarter","ORF Number","Project Duration"]].sort_values(["Year","Total Net"],ascending=[True,False])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        F = _xl_formats(wb)

        # SHEET 1 — SUMMARY
        ws = wb.add_worksheet("Summary"); ws.set_zoom(90); ws.set_tab_color("#1a3a6b")
        writer.sheets["Summary"] = ws
        ws.merge_range("A1:C1",f"Awarded Deals Summary — {TODAY.strftime('%d %B %Y')}",F["title"]); ws.set_row(0,28)
        _hdr(ws,2,["Metric","Value"],[36,22],F["header"])
        for i, row in kpi_df.iterrows():
            ws.write(3+i,0,row["Metric"],F["kpi_lbl"]); ws.write_number(3+i,1,row["Value"],F["kpi_val"])
        ws.merge_range("D1:H1","Summary by Year",F["title"])
        for c, col in enumerate(["Year","Deals","Gross (QAR)","Net (QAR)","Contract Value (QAR)"]):
            ws.write(2,3+c,col,F["header"])
        ws.set_column(3,3,8); ws.set_column(4,4,8); ws.set_column(5,5,20); ws.set_column(6,6,20); ws.set_column(7,7,22)
        for i, row in year_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            ws.write(3+i,3,row["Year"],ft); ws.write_number(3+i,4,row["Count"],fn); ws.write_number(3+i,5,row["Gross"],fn); ws.write_number(3+i,6,row["Net"],fn); ws.write_number(3+i,7,row["PV"],fn)
        tr=3+len(year_df)
        ws.write(tr,3,"TOTAL",F["tot_lbl"]); ws.write_number(tr,4,year_df["Count"].sum(),F["tot_num"]); ws.write_number(tr,5,year_df["Gross"].sum(),F["tot_num"]); ws.write_number(tr,6,year_df["Net"].sum(),F["tot_num"]); ws.write_number(tr,7,year_df["PV"].sum(),F["tot_num"])
        off_s=len(kpi_df)+5; ws.merge_range(off_s,0,off_s,3,"By Stage",F["title"])
        for c, col in enumerate(["Year","Stage","Deals","Net (QAR)"]): ws.write(off_s+1,c,col,F["header"])
        ws.set_column(0,0,8); ws.set_column(1,1,36)
        for i, row in stage_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            ws.write(off_s+2+i,0,row["Year"],ft); ws.write(off_s+2+i,1,row["Stage"],ft); ws.write_number(off_s+2+i,2,row["Count"],fn); ws.write_number(off_s+2+i,3,row["Net"],fn)

        # SHEET 2 — DU BREAKDOWN
        du_ws = wb.add_worksheet("DU Breakdown"); du_ws.set_zoom(90); du_ws.set_tab_color("#FF8C00")
        writer.sheets["DU Breakdown"] = du_ws
        du_ws.merge_range("A1:F1","Gross & Net Breakdown by BU / Delivery Unit",F["title"]); du_ws.set_row(0,28)
        for c,(col,w) in enumerate(zip(["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Net 2025 (QAR)","Net 2026 (QAR)"],[42,52,20,20,20,20])):
            du_ws.write(1,c,col,F["header"]); du_ws.set_column(c,c,w)
        r_out=0
        for bu_name, bu_grp in du_totals.groupby("BU"):
            du_ws.write(2+r_out,0,bu_name,F["bu_lbl"]); du_ws.write(2+r_out,1,"",F["bu_lbl"])
            du_ws.write_number(2+r_out,2,bu_grp["Gross"].sum(),F["bu_num"]); du_ws.write_number(2+r_out,3,bu_grp["Net"].sum(),F["bu_num"]); du_ws.write_number(2+r_out,4,bu_grp["2025"].sum(),F["bu_num"]); du_ws.write_number(2+r_out,5,bu_grp["2026"].sum(),F["bu_num"]); r_out+=1
            for _, row in bu_grp.iterrows():
                alt=r_out%2==1
                du_ws.write(2+r_out,0,"",F["alt"] if alt else F["text"]); du_ws.write(2+r_out,1,row["DU"],F["alt"] if alt else F["text"])
                du_ws.write_number(2+r_out,2,row["Gross"],F["alt_num"] if alt else F["num"]); du_ws.write_number(2+r_out,3,row["Net"],F["alt_num"] if alt else F["num"]); du_ws.write_number(2+r_out,4,row["2025"],F["alt_num"] if alt else F["num"]); du_ws.write_number(2+r_out,5,row["2026"],F["alt_num"] if alt else F["num"]); r_out+=1
                for _, deal in du_exp[du_exp["DU"]==row["DU"]].iterrows():
                    n25=deal["Net"] if deal["Year"]=="2025" else 0; n26=deal["Net"] if deal["Year"]=="2026" else 0
                    du_ws.write(2+r_out,0,"",F["opp"]); du_ws.write(2+r_out,1,f"  ↳  {deal['Opportunity']}",F["opp"])
                    du_ws.write_number(2+r_out,2,deal["Gross"],F["opp_num"]); du_ws.write_number(2+r_out,3,deal["Net"],F["opp_num"]); du_ws.write_number(2+r_out,4,n25,F["opp_num"]); du_ws.write_number(2+r_out,5,n26,F["opp_num"]); r_out+=1
        t=2+r_out
        du_ws.write(t,0,"GRAND TOTAL",F["tot_lbl"]); du_ws.write(t,1,"",F["tot_lbl"])
        du_ws.write_number(t,2,du_totals["Gross"].sum(),F["tot_num"]); du_ws.write_number(t,3,du_totals["Net"].sum(),F["tot_num"]); du_ws.write_number(t,4,du_totals["2025"].sum(),F["tot_num"]); du_ws.write_number(t,5,du_totals["2026"].sum(),F["tot_num"])

        # SHEET 3 — ACCOUNT MANAGER
        am_ws = wb.add_worksheet("Account Manager"); am_ws.set_zoom(90); am_ws.set_tab_color("#228B22")
        writer.sheets["Account Manager"] = am_ws
        am_ws.merge_range("A1:F1",f"Awarded Deals by Account Manager — {TODAY.strftime('%d %B %Y')}",F["title"]); am_ws.set_row(0,28)
        for c,(col,w) in enumerate(zip(["Account Manager","Total Deals","Gross (QAR)","Net (QAR)","Net 2025 (QAR)","Net 2026 (QAR)"],[30,12,20,20,20,20])):
            am_ws.write(1,c,col,F["header"]); am_ws.set_column(c,c,w)
        for i, row in am_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            am_ws.write(2+i,0,row["Account Manager"],F["alt"] if alt else F["text"]); am_ws.write_number(2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); am_ws.write_number(2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); am_ws.write_number(2+i,3,row["Net"],F["alt_num"] if alt else F["num"]); am_ws.write_number(2+i,4,row["Net 2025"],F["alt_num"] if alt else F["num"]); am_ws.write_number(2+i,5,row["Net 2026"],F["alt_num"] if alt else F["num"])
        tr=2+len(am_df)
        am_ws.write(tr,0,"GRAND TOTAL",F["tot_lbl"]); am_ws.write_number(tr,1,am_df["Count"].sum(),F["tot_num"]); am_ws.write_number(tr,2,am_df["Gross"].sum(),F["tot_num"]); am_ws.write_number(tr,3,am_df["Net"].sum(),F["tot_num"]); am_ws.write_number(tr,4,am_df["Net 2025"].sum(),F["tot_num"]); am_ws.write_number(tr,5,am_df["Net 2026"].sum(),F["tot_num"])

        # SHEET 4 — AWARD QUARTER + NEW/RENEW
        aq_ws = wb.add_worksheet("Award Quarter"); aq_ws.set_zoom(90); aq_ws.set_tab_color("#DAA520")
        writer.sheets["Award Quarter"] = aq_ws
        aq_ws.merge_range("A1:E1",f"Net by Award Quarter — {TODAY.strftime('%d %B %Y')}",F["title"]); aq_ws.set_row(0,28)
        for c,(col,w) in enumerate(zip(["Year","Award Quarter","Deals","Gross (QAR)","Net (QAR)"],[8,14,8,20,20])):
            aq_ws.write(1,c,col,F["header"]); aq_ws.set_column(c,c,w)
        for i, row in q_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            aq_ws.write(2+i,0,row["Year"],ft); aq_ws.write(2+i,1,row["Award Quarter"],ft); aq_ws.write_number(2+i,2,row["Count"],fn); aq_ws.write_number(2+i,3,row["Gross"],fn); aq_ws.write_number(2+i,4,row["Net"],fn)
        tr=2+len(q_df)
        aq_ws.write(tr,0,"TOTAL",F["tot_lbl"]); aq_ws.write(tr,1,"",F["tot_lbl"]); aq_ws.write_number(tr,2,q_df["Count"].sum(),F["tot_num"]); aq_ws.write_number(tr,3,q_df["Gross"].sum(),F["tot_num"]); aq_ws.write_number(tr,4,q_df["Net"].sum(),F["tot_num"])
        nr_off=tr+2; aq_ws.merge_range(nr_off,0,nr_off,4,"New vs Renew Breakdown",F["title"])
        for c, col in enumerate(["Year","Type","Deals","Gross (QAR)","Net (QAR)"]): aq_ws.write(nr_off+1,c,col,F["header"])
        for i, row in nr_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            aq_ws.write(nr_off+2+i,0,row["Year"],ft); aq_ws.write(nr_off+2+i,1,row["Type"],ft); aq_ws.write_number(nr_off+2+i,2,row["Count"],fn); aq_ws.write_number(nr_off+2+i,3,row["Gross"],fn); aq_ws.write_number(nr_off+2+i,4,row["Net"],fn)

        # SHEET 5 — ALL AWARDED DEALS
        fd_ws = wb.add_worksheet("All Awarded Deals"); fd_ws.set_zoom(85); fd_ws.set_tab_color("#6495ED")
        writer.sheets["All Awarded Deals"] = fd_ws
        fd_ws.merge_range("A1:O1",f"All Awarded Deals — {TODAY.strftime('%d %B %Y')}",F["title"]); fd_ws.set_row(0,28); fd_ws.freeze_panes(2,0)
        fd_col_names = ["Year","SNo.","Account Name","Opportunity Name","Stage","Account Manager","Type","Total Gross","Total Net","Project Value","Award Quarter","Contracted","Contract Signed Quarter","ORF Number","Project Duration"]
        fd_widths2   = [7,5,32,38,36,24,8,18,18,20,14,12,24,14,16]
        for c,(col,w) in enumerate(zip(fd_col_names,fd_widths2)):
            fd_ws.write(1,c,col,F["header"]); fd_ws.set_column(c,c,w)
        for i, row in full_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            for c, col in enumerate(fd_col_names):
                val=row[col]
                if col in ("Total Gross","Total Net","Project Value","SNo.","Project Duration"):
                    fd_ws.write_number(2+i,c,val if pd.notna(val) else 0,fn)
                else:
                    fd_ws.write(2+i,c,str(val) if pd.notna(val) else "",ft)

        # SHEET 6 — AWARDED BREAKDOWN (styled, one row per DU per opportunity)
        _ACLR = {
            "title_bg":"1F3864","hdr_deal":"1F3864","hdr_du":"17375E",
            "hdr_finance":"1F4E79","hdr_other":"2E5FA3",
            "bu_fill":"EDF2F9","du_fill":"E4ECF7","num_fill":"EBF5FB",
            "tot_fill":"D5E8F5","alt_a":"F5F8FF","alt_b":"FFFFFF",
        }
        def _awf(bg,top,num_fmt=None):
            d={"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
            if num_fmt: d["num_format"]=num_fmt; d["align"]="right"
            return wb.add_format(d)
        aw_fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_ACLR["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        aw_fh_du      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_ACLR["hdr_du"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        aw_fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_ACLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        aw_fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_ACLR["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        aw_fmt_title  = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+_ACLR["title_bg"],"align":"center","valign":"vcenter"})
        aw_ft_a1=_awf(_ACLR["alt_a"],2);  aw_ft_an=_awf(_ACLR["alt_a"],1)
        aw_ft_b1=_awf(_ACLR["alt_b"],2);  aw_ft_bn=_awf(_ACLR["alt_b"],1)
        aw_fn_a1=_awf(_ACLR["alt_a"],2,"#,##0"); aw_fn_an=_awf(_ACLR["alt_a"],1,"#,##0")
        aw_fn_b1=_awf(_ACLR["alt_b"],2,"#,##0"); aw_fn_bn=_awf(_ACLR["alt_b"],1,"#,##0")
        aw_fbu1=_awf(_ACLR["bu_fill"],2); aw_fbun=_awf(_ACLR["bu_fill"],1)
        aw_fdu1=_awf(_ACLR["du_fill"],2); aw_fdun=_awf(_ACLR["du_fill"],1)
        aw_fxn1=_awf(_ACLR["num_fill"],2,"#,##0"); aw_fxnn=_awf(_ACLR["num_fill"],1,"#,##0")
        aw_ftot  =wb.add_format({"bg_color":"#"+_ACLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
        aw_ftotbl=wb.add_format({"bg_color":"#"+_ACLR["num_fill"],"border":1,"font_size":9})
        aw_oc=[
            ("SNo.",6,"deal"),("Account Name",24,"deal"),("Opportunity Name",36,"deal"),
            ("BU",36,"du"),("DU",34,"du"),
            ("Gross (breakdown)",16,"finance"),("Net (breakdown)",16,"finance"),
            ("Total Gross",15,"finance"),("Total Net",15,"finance"),
            ("Stage",28,"other"),("Account Manager",22,"other"),
            ("Award Quarter",10,"other"),("Contracted",10,"other"),("Year",8,"other"),
        ]
        aw_hfmt={"deal":aw_fh_deal,"du":aw_fh_du,"finance":aw_fh_finance,"other":aw_fh_other}
        aw_nc=len(aw_oc)
        aw_bw=wb.add_worksheet("Awarded Breakdown"); aw_bw.set_zoom(85); aw_bw.set_tab_color("#9370DB")
        writer.sheets["Awarded Breakdown"]=aw_bw
        aw_bw.merge_range(0,0,0,aw_nc-1,"Awarded Deals — Expanded by Delivery Unit",aw_fmt_title); aw_bw.set_row(0,28)
        for c,(cn,cw,ct) in enumerate(aw_oc): aw_bw.write(1,c,cn,aw_hfmt[ct]); aw_bw.set_column(c,c,cw)
        aw_bw.set_row(1,28); aw_bw.freeze_panes(2,0)
        aw_coa=load_coa()
        aw_exp=_expand_deals(df,mapping=aw_coa)
        aw_bw.autofilter(1,0,1+len(aw_exp),aw_nc-1)
        aw_cm={cn:idx for idx,(cn,_,__) in enumerate(aw_oc)}
        aw_prev=None; aw_alt=False
        for rp,(_, row) in enumerate(aw_exp.iterrows()):
            didx=row["_deal_idx"]; isf=bool(row["_is_first"])
            if didx!=aw_prev: aw_alt=not aw_alt; aw_prev=didx
            ft=(aw_ft_a1 if isf else aw_ft_an) if aw_alt else (aw_ft_b1 if isf else aw_ft_bn)
            fn=(aw_fn_a1 if isf else aw_fn_an) if aw_alt else (aw_fn_b1 if isf else aw_fn_bn)
            fbu=aw_fbu1 if isf else aw_fbun; fdu=aw_fdu1 if isf else aw_fdun
            fxn=aw_fxn1 if isf else aw_fxnn; xl_r=2+rp
            def _aws(ci,val,fmt):
                if val is None or (isinstance(val,float) and pd.isna(val)): aw_bw.write_blank(xl_r,ci,None,fmt)
                else: aw_bw.write(xl_r,ci,str(val),fmt)
            def _awn(ci,val,fmt):
                v=_parse_num(val) if not isinstance(val,(int,float)) else val
                if v is None or (isinstance(v,float) and pd.isna(v)): aw_bw.write_blank(xl_r,ci,None,fmt)
                else: aw_bw.write_number(xl_r,ci,v,fmt)
            _aws(aw_cm["SNo."],             row.get("SNo.")             if isf else None,ft)
            _aws(aw_cm["Account Name"],      row.get("Account Name")     if isf else None,ft)
            _aws(aw_cm["Opportunity Name"],  row.get("Opportunity Name") if isf else None,ft)
            _aws(aw_cm["BU"],row.get("BU_exp"),fbu); _aws(aw_cm["DU"],row.get("DU_exp"),fdu)
            _awn(aw_cm["Gross (breakdown)"], row.get("Gross_exp"),fxn)
            _awn(aw_cm["Net (breakdown)"],   row.get("Net_exp"),  fxn)
            if isf:
                _awn(aw_cm["Total Gross"],row.get("Total Gross"),aw_ftot)
                _awn(aw_cm["Total Net"],  row.get("Total Net"),  aw_ftot)
            else:
                aw_bw.write_blank(xl_r,aw_cm["Total Gross"],None,aw_ftotbl)
                aw_bw.write_blank(xl_r,aw_cm["Total Net"],  None,aw_ftotbl)
            _aws(aw_cm["Stage"],        row.get("Stage")         if isf else None,ft)
            _aws(aw_cm["Account Manager"],row.get("Account Manager") if isf else None,ft)
            _aws(aw_cm["Award Quarter"],row.get("Award Quarter") if isf else None,ft)
            _aws(aw_cm["Contracted"],   row.get("Contracted")    if isf else None,ft)
            _aws(aw_cm["Year"],         row.get("Year")          if isf else None,ft)

    output.seek(0)
    return output.read()

# ── SIDEBAR ────────────────────────────────────────────────────────────────────
st.sidebar.title("📁 Load Data")
uploaded      = st.sidebar.file_uploader("Pipeline Excel (weekly report)", type=["xlsx","xls"])
uploaded_aw   = st.sidebar.file_uploader("Awarded Deals 2026",             type=["xlsx","xls"])
uploaded_aw25 = st.sidebar.file_uploader("Awarded Deals 2025",             type=["xlsx","xls"])
uploaded_b3   = st.sidebar.file_uploader("Book3 (Resource Forecast)",      type=["xlsx","xls"])

have_awarded = uploaded_aw or uploaded_aw25

if not uploaded and not have_awarded:
    st.info("👆 Upload at least one Excel file to get started.")
    st.stop()

if uploaded:
    df_raw = load_data(uploaded)
    du_exp = build_du_breakdown(uploaded)
    st.sidebar.success("Pipeline file loaded ✓")

if have_awarded:
    parts_raw, parts_du = [], []
    if uploaded_aw:
        r = load_awarded(uploaded_aw); r["Year"] = "2026"; parts_raw.append(r)
        d = build_awarded_du(uploaded_aw); d["Year"] = "2026"; parts_du.append(d)
        st.sidebar.success("Awarded Deals 2026 loaded ✓")
    if uploaded_aw25:
        r = load_awarded(uploaded_aw25); r["Year"] = "2025"; parts_raw.append(r)
        d = build_awarded_du(uploaded_aw25); d["Year"] = "2025"; parts_du.append(d)
        st.sidebar.success("Awarded Deals 2025 loaded ✓")
    aw_raw = pd.concat(parts_raw, ignore_index=True)
    aw_du  = pd.concat(parts_du,  ignore_index=True)

# ── TABS ──────────────────────────────────────────────────────────────────────
st.title("📊 Sales Weekly Review Dashboard")
st.caption(f"Report date: {date.today().strftime('%d %B %Y')}")

tab_labels = []
if uploaded:     tab_labels.append("🔵 Pipeline")
if have_awarded: tab_labels.append("🟢 Awarded Deals")
if uploaded_b3:  tab_labels.append("🔗 Book3 Mapping")
tabs = st.tabs(tab_labels)
tab_idx = {name: i for i, name in enumerate(tab_labels)}

# ── PIPELINE FILTERS (sidebar) ────────────────────────────────────────────────
if uploaded:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔵 Pipeline Filters")
    sel_sector  = st.sidebar.multiselect("Sector",             sorted(df_raw["Sector"].dropna().unique()), default=[])
    sel_mgr     = st.sidebar.multiselect("Account Manager",    sorted(df_raw["Account Manager"].dropna().unique()), default=[])
    sel_quarter = st.sidebar.multiselect("Quarter",            sorted(df_raw["Closure Due Quarter"].dropna().unique()), default=[])
    sel_prob    = st.sidebar.multiselect("Winning Probability",sorted(df_raw["Winning Probability"].dropna().unique()), default=[])
    sel_bu      = st.sidebar.multiselect("BU",                 sorted(du_exp["BU"].dropna().unique()), default=[])
    show_overdue= st.sidebar.checkbox("Show only overdue deals", value=False)

    du_filtered = du_exp.copy()
    if sel_sector:  du_filtered = du_filtered[du_filtered["Sector"].isin(sel_sector)]
    if sel_mgr:     du_filtered = du_filtered[du_filtered["Account Manager"].isin(sel_mgr)]
    if sel_quarter: du_filtered = du_filtered[du_filtered["Closure Due Quarter"].isin(sel_quarter)]
    if sel_bu:      du_filtered = du_filtered[du_filtered["BU"].isin(sel_bu)]

    if sel_bu:
        bu_opps = du_filtered.set_index(["Account Name","Lead/Opp Name"]).index
        df_base = df_raw[df_raw.set_index(["Account Name","Lead/Opp Name"]).index.isin(bu_opps)].copy()
    else:
        df_base = df_raw.copy()

    df = df_base.copy()
    if sel_sector:  df = df[df["Sector"].isin(sel_sector)]
    if sel_mgr:     df = df[df["Account Manager"].isin(sel_mgr)]
    if sel_quarter: df = df[df["Closure Due Quarter"].isin(sel_quarter)]
    if sel_prob:    df = df[df["Winning Probability"].isin(sel_prob)]
    if show_overdue:df = df[df["Overdue"]]

# ── AWARDED FILTERS (sidebar) ─────────────────────────────────────────────────
if have_awarded:
    st.sidebar.markdown("---")
    st.sidebar.subheader("🟢 Awarded Filters")
    aw_sel_year = st.sidebar.multiselect("Year",              sorted(aw_raw["Year"].dropna().unique()), default=[])
    aw_sel_mgr  = st.sidebar.multiselect("AM (Awarded)",      sorted(aw_raw["Account Manager"].dropna().unique()), default=[])
    aw_sel_q    = st.sidebar.multiselect("Award Quarter",     sorted(aw_raw["Award Quarter"].dropna().unique()), default=[])
    aw_sel_stage= st.sidebar.multiselect("Stage (Awarded)",   sorted(aw_raw["Stage"].dropna().unique()), default=[])
    aw_sel_type = st.sidebar.multiselect("New / Renew",       ["New","Renew","Mixed"], default=[])
    aw_sel_bu   = st.sidebar.multiselect("BU (Awarded)",      sorted(aw_du["BU"].dropna().unique()), default=[])

    aw = aw_raw.copy()
    if aw_sel_year:  aw = aw[aw["Year"].isin(aw_sel_year)]
    if aw_sel_mgr:   aw = aw[aw["Account Manager"].isin(aw_sel_mgr)]
    if aw_sel_q:     aw = aw[aw["Award Quarter"].isin(aw_sel_q)]
    if aw_sel_stage: aw = aw[aw["Stage"].isin(aw_sel_stage)]
    if aw_sel_type:  aw = aw[aw["Type"].isin(aw_sel_type)]

    aw_du_f = aw_du.copy()
    if aw_sel_year: aw_du_f = aw_du_f[aw_du_f["Year"].isin(aw_sel_year)]
    if aw_sel_mgr:  aw_du_f = aw_du_f[aw_du_f["Account Manager"].isin(aw_sel_mgr)]
    if aw_sel_q:    aw_du_f = aw_du_f[aw_du_f["Award Quarter"].isin(aw_sel_q)]
    if aw_sel_bu:
        aw_du_f = aw_du_f[aw_du_f["BU"].isin(aw_sel_bu)]
        bu_opp_aw = aw_du_f.set_index(["Account Name","Opportunity"]).index
        aw = aw[aw.set_index(["Account Name","Opportunity Name"]).index.isin(bu_opp_aw)]

# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE TAB
# ══════════════════════════════════════════════════════════════════════════════
if uploaded:
  with tabs[tab_idx["🔵 Pipeline"]]:
    cap_col, btn_col = st.columns([6, 2])
    cap_col.caption(f"{len(df)} opportunities after filters")
    with btn_col:
        xl_bytes = export_pipeline_excel(uploaded, book3_file=uploaded_b3, awarded_file=uploaded_aw)
        st.download_button(
            label="⬇️ Export Excel Report",
            data=xl_bytes,
            file_name=f"Pipeline_Report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # KPIs
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Opportunities",  len(df))
    k2.metric("Total Gross Pipeline", fmt_m(df["Total Gross"].sum()))
    k3.metric("Total Net Pipeline",   fmt_m(df["Total Net"].sum()))
    k4.metric("Forecasted (Net)",     fmt_m(df[df["Forecasted"] == "Yes"]["Total Net"].sum()))
    k5.metric("Strategic Opps",       len(df[df["Strategic Opportunity"] == "Yes"]))
    k6.metric("⚠️ Overdue Deals",     int(df["Overdue"].sum()),
              delta=None if not int(df["Overdue"].sum()) else "Needs attention",
              delta_color="inverse")
    st.markdown("---")

    # Stage funnel + Sector
    col_left, col_right = st.columns(2)
    with col_left:
        st.subheader("Pipeline Funnel by Stage")
        stage_df = df.groupby("Stage_Short")["Total Net"].agg(["sum","count"]).reset_index()
        stage_df.columns = ["Stage_Short","Total Net","Opps"]
        order = [STAGE_SHORT[s] for s in STAGE_ORDER if STAGE_SHORT[s] in stage_df["Stage_Short"].values]
        stage_df["Stage_Short"] = pd.Categorical(stage_df["Stage_Short"], categories=order, ordered=True)
        stage_df = stage_df.sort_values("Stage_Short")
        stage_df["Color"] = stage_df["Stage_Short"].map(STAGE_COLORS)
        fig_stage = go.Figure(go.Funnel(
            y=stage_df["Stage_Short"], x=stage_df["Total Net"],
            textinfo="value+percent initial",
            text=[f"{r['Opps']} opps | {fmt_m(r['Total Net'])}" for _, r in stage_df.iterrows()],
            marker=dict(color=stage_df["Color"].tolist()),
        ))
        fig_stage.update_layout(height=380, margin=dict(l=10,r=10,t=10,b=10))
        st.plotly_chart(fig_stage, use_container_width=True)
    with col_right:
        st.subheader("Pipeline by Sector (Net)")
        sector_df = (df.groupby("Sector")["Total Net"].agg(["sum","count"]).reset_index()
                     .rename(columns={"sum":"Total Net","count":"Opps"})
                     .sort_values("Total Net", ascending=True).tail(12))
        fig_sector = px.bar(sector_df, x="Total Net", y="Sector", orientation="h",
                            text=sector_df["Total Net"].apply(fmt_m),
                            color="Total Net", color_continuous_scale="Blues", hover_data={"Opps":True})
        fig_sector.update_traces(textposition="outside")
        fig_sector.update_layout(height=380, coloraxis_showscale=False, margin=dict(l=10,r=80,t=10,b=10))
        st.plotly_chart(fig_sector, use_container_width=True)

    # Account Manager + Quarterly
    col_am, col_q = st.columns(2)
    with col_am:
        st.subheader("Account Manager Performance")
        am_df = (df.groupby("Account Manager")
                 .agg(Opps=("Lead/Opp Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
                 .reset_index().sort_values("Net", ascending=False))
        fig_am = go.Figure()
        fig_am.add_trace(go.Bar(name="Gross", x=am_df["Account Manager"], y=am_df["Gross"], marker_color="#6495ED", opacity=0.6))
        fig_am.add_trace(go.Bar(name="Net",   x=am_df["Account Manager"], y=am_df["Net"],   marker_color="#1a3a6b"))
        fig_am.update_layout(barmode="overlay", height=300, yaxis_title="QAR",
                             margin=dict(l=10,r=10,t=10,b=80), legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_am, use_container_width=True)
        st.dataframe(am_df.assign(**{"Gross (M)":am_df["Gross"]/1e6,"Net (M)":am_df["Net"]/1e6})
                     [["Account Manager","Opps","Gross (M)","Net (M)"]]
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)
    with col_q:
        st.subheader("Quarterly Close Plan")
        q_df = (df.groupby("Closure Due Quarter").agg(Net=("Total Net","sum"), Opps=("Lead/Opp Name","count"))
                .reset_index().sort_values("Closure Due Quarter"))
        fig_q = px.bar(q_df, x="Closure Due Quarter", y="Net", text=q_df["Net"].apply(fmt_m),
                       color="Closure Due Quarter", color_discrete_sequence=px.colors.qualitative.Set2,
                       hover_data={"Opps":True})
        fig_q.update_traces(textposition="outside")
        fig_q.update_layout(height=300, showlegend=False, yaxis_title="Net Value (QAR)", margin=dict(l=10,r=10,t=10,b=40))
        st.plotly_chart(fig_q, use_container_width=True)
        prob_df = df.groupby("Winning Probability")["Total Net"].sum().reset_index()
        fig_prob = px.pie(prob_df, names="Winning Probability", values="Total Net",
                          color_discrete_map={"High":"#228B22","Moderate":"#DAA520","Low":"#CD5C5C"},
                          title="Net Pipeline by Winning Probability")
        fig_prob.update_layout(height=260, margin=dict(l=10,r=10,t=40,b=10))
        st.plotly_chart(fig_prob, use_container_width=True)

    # Strategic + Source
    col_s, col_b = st.columns(2)
    with col_s:
        st.subheader("Strategic vs Regular")
        strat_df = df.groupby("Strategic Opportunity").agg(Count=("Lead/Opp Name","count"),Net=("Total Net","sum")).reset_index()
        fig_strat = px.pie(strat_df, names="Strategic Opportunity", values="Net",
                           color_discrete_map={"Yes":"#FF8C00","No":"#6495ED"}, hole=0.4)
        fig_strat.update_traces(textinfo="label+percent+value", texttemplate="%{label}<br>%{percent}<br>QAR %{value:,.0f}")
        fig_strat.update_layout(height=300, margin=dict(l=10,r=10,t=20,b=10), showlegend=False)
        st.plotly_chart(fig_strat, use_container_width=True)
    with col_b:
        st.subheader("Source of Opportunity")
        src_df = df.groupby("Source of Opportunity").agg(Count=("Lead/Opp Name","count"),Net=("Total Net","sum")).reset_index().sort_values("Net", ascending=False)
        fig_src = px.bar(src_df, x="Source of Opportunity", y="Count", color="Net", text="Count",
                         color_continuous_scale="Oranges", labels={"Net":"Net Value"})
        fig_src.update_traces(textposition="outside")
        fig_src.update_layout(height=300, coloraxis_showscale=True, margin=dict(l=10,r=10,t=10,b=60))
        st.plotly_chart(fig_src, use_container_width=True)

    # DU Breakdown
    st.markdown("---")
    st.subheader("🏢 Gross vs Net Breakdown by BU / Delivery Unit")
    du_totals = (du_filtered.groupby(["BU","DU"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","Net"], ascending=[True,False]))
    du_totals["DU_Label"] = du_totals["DU"].str.replace(r"^\d+\s*", "", regex=True)
    fore_by_du = (du_filtered[du_filtered["Forecasted"]=="Yes"].groupby("DU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"}))
    du_totals = du_totals.merge(fore_by_du, on="DU", how="left").fillna({"Forecasted Net":0})
    col_du1, col_du2 = st.columns([3, 2])
    with col_du1:
        fig_du = px.bar(du_totals, x="DU_Label", y=["Gross","Net"], barmode="group",
                        color_discrete_map={"Gross":"#6495ED","Net":"#1a3a6b"},
                        hover_data={"BU":True}, labels={"value":"QAR","variable":"","DU_Label":"DU"})
        fig_du.update_layout(height=400, yaxis_title="QAR", xaxis_tickangle=-35,
                             margin=dict(l=10,r=10,t=10,b=130), legend=dict(orientation="h",y=1.05))
        st.plotly_chart(fig_du, use_container_width=True)
    with col_du2:
        st.markdown("**Forecasted Net by DU**")
        du_fore = du_totals[du_totals["Forecasted Net"] > 0].sort_values("Forecasted Net", ascending=True)
        if not du_fore.empty:
            fig_du_fore = px.bar(du_fore, x="Forecasted Net", y="DU_Label", orientation="h",
                                 text=du_fore["Forecasted Net"].apply(fmt_m), color="BU")
            fig_du_fore.update_traces(textposition="outside")
            fig_du_fore.update_layout(height=400, margin=dict(l=10,r=90,t=10,b=10),
                                       yaxis_title="", xaxis_title="Net (QAR)",
                                       legend=dict(orientation="h", y=-0.25, font=dict(size=10)))
            st.plotly_chart(fig_du_fore, use_container_width=True)
        else:
            st.info("No forecasted deals match current filters.")
    bu_totals = (du_filtered.groupby("BU")[["Gross","Net"]].sum().reset_index().sort_values("Net", ascending=False))
    fore_bu   = (du_filtered[du_filtered["Forecasted"]=="Yes"].groupby("BU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"}))
    bu_totals = bu_totals.merge(fore_bu, on="BU", how="left").fillna({"Forecasted Net":0})
    bu_totals["Gross (M)"] = bu_totals["Gross"]/1e6
    bu_totals["Net (M)"]   = bu_totals["Net"]/1e6
    bu_totals["Forecasted Net (M)"] = bu_totals["Forecasted Net"]/1e6
    with st.expander("BU Summary Table", expanded=True):
        st.dataframe(bu_totals[["BU","Gross (M)","Net (M)","Forecasted Net (M)"]]
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}","Forecasted Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)
    du_table = du_totals.copy()
    du_table["Gross (M)"] = du_table["Gross"]/1e6
    du_table["Net (M)"]   = du_table["Net"]/1e6
    du_table["Forecasted Net (M)"] = du_table["Forecasted Net"]/1e6
    with st.expander("DU Detail Table"):
        st.dataframe(du_table[["BU","DU_Label","Gross (M)","Net (M)","Forecasted Net (M)"]]
                     .rename(columns={"DU_Label":"Delivery Unit"})
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}","Forecasted Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)

    # Forecast per DU
    st.markdown("---")
    st.subheader("📅 Forecast per DU")
    fore_du_exp = du_filtered[du_filtered["Forecasted"] == "Yes"].copy()
    fore_du_exp["DU_Label"] = fore_du_exp["DU"].str.replace(r"^\d+\s*", "", regex=True)
    fd1, fd2, fd3 = st.columns(3)
    fd1.metric("Forecasted Deals", len(fore_du_exp))
    fd2.metric("Forecasted Gross", fmt_m(fore_du_exp["Gross"].sum()))
    fd3.metric("Forecasted Net",   fmt_m(fore_du_exp["Net"].sum()))
    col_fd1, col_fd2 = st.columns(2)
    with col_fd1:
        st.markdown("**Forecasted Net by BU**")
        fore_bu_chart = fore_du_exp.groupby("BU")[["Gross","Net"]].sum().reset_index().sort_values("Net", ascending=True)
        fig_fbu = px.bar(fore_bu_chart, x="Net", y="BU", orientation="h",
                         text=fore_bu_chart["Net"].apply(fmt_m), color="Net", color_continuous_scale="Greens")
        fig_fbu.update_traces(textposition="outside")
        fig_fbu.update_layout(height=320, coloraxis_showscale=False, margin=dict(l=10,r=90,t=10,b=10), yaxis_title="")
        st.plotly_chart(fig_fbu, use_container_width=True)
    with col_fd2:
        st.markdown("**Forecasted Net by DU & Quarter**")
        fore_dq = fore_du_exp.groupby(["DU_Label","Closure Due Quarter"])["Net"].sum().reset_index()
        fig_fdq = px.bar(fore_dq, x="DU_Label", y="Net", color="Closure Due Quarter", barmode="stack",
                         color_discrete_sequence=px.colors.qualitative.Set2,
                         labels={"Net":"Net (QAR)","DU_Label":"DU"})
        fig_fdq.update_layout(height=320, xaxis_tickangle=-30, margin=dict(l=10,r=10,t=10,b=110),
                               legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_fdq, use_container_width=True)
    fore_summary = (fore_du_exp.groupby(["BU","DU_Label","Closure Due Quarter"])
                    .agg(Count=("Lead/Opp Name","count"), Gross=("Gross","sum"), Net=("Net","sum"))
                    .reset_index().sort_values(["BU","DU_Label","Closure Due Quarter"])
                    .rename(columns={"DU_Label":"Delivery Unit"}))
    with st.expander("Forecast Summary by BU / DU / Quarter", expanded=True):
        st.dataframe(fore_summary.style.format({"Gross":"{:,.0f}","Net":"{:,.0f}"}),
                     use_container_width=True, hide_index=True)
    with st.expander("Forecast Deal Detail"):
        fore_detail = (fore_du_exp[["BU","DU_Label","Account Name","Lead/Opp Name","Stage",
                                     "Account Manager","Sector","Closure Due Quarter","Gross","Net",
                                     "Winning Probability","Est. Close Date"]]
                       .sort_values(["BU","DU_Label","Net"], ascending=[True,True,False])
                       .rename(columns={"DU_Label":"Delivery Unit"}))
        st.dataframe(fore_detail.style.format(
            {"Gross":"{:,.0f}","Net":"{:,.0f}",
             "Est. Close Date": lambda d: d.strftime("%d-%b-%Y") if pd.notna(d) else "—"}),
            use_container_width=True, hide_index=True)

    # Forecast Analysis
    st.markdown("---")
    st.subheader("📈 Forecast Analysis")
    fore_deals = df[df["Forecasted"] == "Yes"].copy()
    not_fore   = df[df["Forecasted"] == "No"].copy()
    fc1, fc2, fc3, fc4 = st.columns(4)
    fc1.metric("Forecasted Deals",     len(fore_deals))
    fc2.metric("Forecasted Net",       fmt_m(fore_deals["Total Net"].sum()))
    fc3.metric("Forecasted Gross",     fmt_m(fore_deals["Total Gross"].sum()))
    fc4.metric("Not Forecasted (Net)", fmt_m(not_fore["Total Net"].sum()))
    col_fq, col_fs = st.columns(2)
    with col_fq:
        st.markdown("**Forecasted Net by Quarter**")
        fq = fore_deals.groupby("Closure Due Quarter")[["Total Net","Total Gross"]].sum().reset_index().sort_values("Closure Due Quarter")
        fig_fq = go.Figure()
        fig_fq.add_trace(go.Bar(name="Gross", x=fq["Closure Due Quarter"], y=fq["Total Gross"],
                                marker_color="#6495ED", opacity=0.6, text=fq["Total Gross"].apply(fmt_m), textposition="outside"))
        fig_fq.add_trace(go.Bar(name="Net",   x=fq["Closure Due Quarter"], y=fq["Total Net"],
                                marker_color="#228B22", text=fq["Total Net"].apply(fmt_m), textposition="outside"))
        fig_fq.update_layout(barmode="group", height=320, margin=dict(l=10,r=10,t=10,b=40), legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_fq, use_container_width=True)
    with col_fs:
        st.markdown("**Forecasted Net by Stage**")
        fs = fore_deals.groupby("Stage_Short")["Total Net"].sum().reset_index().sort_values("Total Net", ascending=True)
        fig_fs = px.bar(fs, x="Total Net", y="Stage_Short", orientation="h",
                        text=fs["Total Net"].apply(fmt_m), color="Total Net", color_continuous_scale="Greens")
        fig_fs.update_traces(textposition="outside")
        fig_fs.update_layout(height=320, coloraxis_showscale=False, margin=dict(l=10,r=80,t=10,b=10), yaxis_title="")
        st.plotly_chart(fig_fs, use_container_width=True)
    with st.expander("Forecasted Deals Detail"):
        st.dataframe(fore_deals[["Account Name","Lead/Opp Name","Stage_Short","Account Manager",
                                  "Sector","Total Gross","Total Net","Winning Probability",
                                  "Closure Due Quarter","Est. Close Date"]]
                     .sort_values("Total Net", ascending=False).rename(columns={"Stage_Short":"Stage"})
                     .style.format({"Total Gross":"{:,.0f}","Total Net":"{:,.0f}",
                                    "Est. Close Date": lambda d: d.strftime("%d-%b-%Y") if pd.notna(d) else "—"}),
                     use_container_width=True, hide_index=True)

    # Overdue
    overdue_df = df_raw[df_raw["Overdue"]].copy()
    if not overdue_df.empty:
        st.markdown("---")
        st.subheader(f"⚠️ Overdue Deals ({len(overdue_df)})")
        st.dataframe(overdue_df[["Account Name","Lead/Opp Name","Stage","Account Manager",
                                  "Total Net","Est. Close Date","Winning Probability"]]
                     .sort_values("Est. Close Date")
                     .style.format({"Total Net":"{:,.0f}",
                                    "Est. Close Date": lambda d: d.strftime("%d-%b-%Y") if pd.notna(d) else "—"})
                     .applymap(lambda _: "background-color: #fff3cd", subset=["Account Name"]),
                     use_container_width=True, hide_index=True)

    # Full table
    st.markdown("---")
    st.subheader("📋 Full Pipeline Table")
    search = st.text_input("Search by account / opportunity name", "")
    disp_df = df.copy()
    if search:
        mask = (disp_df["Account Name"].str.contains(search, case=False, na=False) |
                disp_df["Lead/Opp Name"].str.contains(search, case=False, na=False))
        disp_df = disp_df[mask]
    cols_show = ["SNo.","Account Name","Lead/Opp Name","Stage_Short","Account Manager","Sector",
                 "Total Gross","Total Net","Winning Probability","Forecasted",
                 "Closure Due Quarter","Est. Close Date","Strategic Opportunity","Overdue"]
    def highlight_overdue(row):
        color = "background-color: #ffe0e0" if row["Overdue"] else ""
        return [color] * len(row)
    st.dataframe(disp_df[cols_show].sort_values("Total Net", ascending=False)
                 .rename(columns={"Stage_Short":"Stage"})
                 .style.apply(highlight_overdue, axis=1)
                 .format({"Total Gross":"{:,.0f}","Total Net":"{:,.0f}",
                          "Est. Close Date": lambda d: d.strftime("%d-%b-%Y") if pd.notna(d) else "—"}),
                 use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# AWARDED DEALS TAB
# ══════════════════════════════════════════════════════════════════════════════
if have_awarded:
  with tabs[tab_idx["🟢 Awarded Deals"]]:
    years_loaded = sorted(aw_raw["Year"].unique())
    cap_col2, btn_col2 = st.columns([6, 2])
    cap_col2.caption(f"{len(aw)} awarded deals after filters · Years: {', '.join(years_loaded)}")
    with btn_col2:
        aw_xl_bytes = export_awarded_excel(uploaded_aw, uploaded_aw25)
        st.download_button(
            label="⬇️ Export Excel Report",
            data=aw_xl_bytes,
            file_name=f"Awarded_Report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # KPIs
    contracted = aw[aw["Contracted"]=="Yes"]
    loa        = aw[aw["Stage"]=="Stage 6: Letter Of Award"]
    a1,a2,a3,a4,a5,a6 = st.columns(6)
    a1.metric("Total Awarded Deals",   len(aw))
    a2.metric("Total Gross",           fmt_m(aw["Total Gross"].sum()))
    a3.metric("Total Net",             fmt_m(aw["Total Net"].sum()))
    a4.metric("Contract Value",        fmt_m(aw["Project Value"].sum()))
    a5.metric("Contracted (signed)",   len(contracted))
    a6.metric("LOA (not yet signed)",  len(loa))
    st.markdown("---")

    # Stage + New/Renew
    c1, c2, c3 = st.columns(3)
    with c1:
        st.subheader("By Stage")
        stg = aw.groupby("Stage").agg(Count=("Opportunity Name","count"), Net=("Total Net","sum")).reset_index()
        fig_stg = px.bar(stg, x="Stage", y="Net", text=stg["Net"].apply(fmt_m),
                         color="Stage", color_discrete_map={
                             "Stage 6: Letter Of Award":"#FF8C00",
                             "Stage 7: Contracting And Sign Off":"#228B22"})
        fig_stg.update_traces(textposition="outside")
        fig_stg.update_layout(height=320, showlegend=False, yaxis_title="Net (QAR)",
                               margin=dict(l=10,r=10,t=10,b=60), xaxis_tickangle=-10)
        st.plotly_chart(fig_stg, use_container_width=True)
    with c2:
        st.subheader("New vs Renew")
        nr = aw.groupby("Type").agg(Count=("Opportunity Name","count"), Net=("Total Net","sum")).reset_index()
        fig_nr = px.pie(nr, names="Type", values="Net",
                        color_discrete_map={"New":"#1a3a6b","Renew":"#6495ED","Mixed":"#DAA520"},
                        hole=0.4)
        fig_nr.update_traces(textinfo="label+percent+value",
                              texttemplate="%{label}<br>%{percent}<br>QAR %{value:,.0f}")
        fig_nr.update_layout(height=320, margin=dict(l=10,r=10,t=10,b=10), showlegend=False)
        st.plotly_chart(fig_nr, use_container_width=True)
    with c3:
        st.subheader("Contracted Status")
        ct = aw.groupby("Contracted").agg(Count=("Opportunity Name","count"), Net=("Total Net","sum")).reset_index()
        fig_ct = px.pie(ct, names="Contracted", values="Count",
                        color_discrete_map={"Yes":"#228B22","No":"#CD5C5C"},
                        hole=0.4)
        fig_ct.update_traces(textinfo="label+percent+value",
                              texttemplate="%{label}<br>%{percent}<br>%{value} deals")
        fig_ct.update_layout(height=320, margin=dict(l=10,r=10,t=10,b=10), showlegend=False)
        st.plotly_chart(fig_ct, use_container_width=True)

    # Account Manager + Award Quarter
    c4, c5 = st.columns(2)
    with c4:
        st.subheader("Account Manager — Gross vs Net")
        am_aw = (aw.groupby("Account Manager")
                 .agg(Deals=("Opportunity Name","count"), Gross=("Total Gross","sum"), Net=("Total Net","sum"))
                 .reset_index().sort_values("Net", ascending=False))
        fig_am_aw = go.Figure()
        fig_am_aw.add_trace(go.Bar(name="Gross", x=am_aw["Account Manager"], y=am_aw["Gross"], marker_color="#6495ED", opacity=0.6))
        fig_am_aw.add_trace(go.Bar(name="Net",   x=am_aw["Account Manager"], y=am_aw["Net"],   marker_color="#1a3a6b"))
        fig_am_aw.update_layout(barmode="overlay", height=320, yaxis_title="QAR",
                                 margin=dict(l=10,r=10,t=10,b=80), legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_am_aw, use_container_width=True)
        st.dataframe(am_aw.assign(**{"Gross (M)":am_aw["Gross"]/1e6,"Net (M)":am_aw["Net"]/1e6})
                     [["Account Manager","Deals","Gross (M)","Net (M)"]]
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)
    with c5:
        st.subheader("Net by Award Quarter")
        aq = (aw.groupby(["Year","Award Quarter"]).agg(Net=("Total Net","sum"), Deals=("Opportunity Name","count"))
              .reset_index().sort_values(["Year","Award Quarter"]))
        fig_aq = px.bar(aq, x="Award Quarter", y="Net", color="Year", barmode="group",
                        text=aq["Net"].apply(fmt_m),
                        color_discrete_map={"2025":"#6495ED","2026":"#228B22"},
                        hover_data={"Deals":True})
        fig_aq.update_traces(textposition="outside")
        fig_aq.update_layout(height=320, yaxis_title="Net (QAR)", margin=dict(l=10,r=10,t=10,b=40),
                              legend=dict(orientation="h",y=1.1))
        st.plotly_chart(fig_aq, use_container_width=True)

        st.markdown("**Contracted vs LOA by Quarter**")
        aq2 = aw.groupby(["Award Quarter","Stage"]).agg(Net=("Total Net","sum")).reset_index()
        fig_aq2 = px.bar(aq2, x="Award Quarter", y="Net", color="Stage", barmode="stack",
                         color_discrete_map={"Stage 6: Letter Of Award":"#FF8C00",
                                             "Stage 7: Contracting And Sign Off":"#228B22"})
        fig_aq2.update_layout(height=260, margin=dict(l=10,r=10,t=10,b=40),
                               legend=dict(orientation="h", y=-0.4, font=dict(size=10)))
        st.plotly_chart(fig_aq2, use_container_width=True)

    # DU Breakdown
    st.markdown("---")
    st.subheader("🏢 Gross vs Net by BU / Delivery Unit")
    du_aw = (aw_du_f.groupby(["BU","DU"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","Net"], ascending=[True,False]))
    du_aw["DU_Label"] = du_aw["DU"].str.replace(r"^\d+\s*", "", regex=True)
    c6, c7 = st.columns([3,2])
    with c6:
        fig_du_aw = px.bar(du_aw, x="DU_Label", y=["Gross","Net"], barmode="group",
                           color_discrete_map={"Gross":"#6495ED","Net":"#1a3a6b"},
                           hover_data={"BU":True}, labels={"value":"QAR","variable":"","DU_Label":"DU"})
        fig_du_aw.update_layout(height=380, xaxis_tickangle=-35, margin=dict(l=10,r=10,t=10,b=130),
                                 legend=dict(orientation="h",y=1.05))
        st.plotly_chart(fig_du_aw, use_container_width=True)
    with c7:
        st.markdown("**Net by BU**")
        bu_aw = du_aw.groupby("BU")[["Gross","Net"]].sum().reset_index().sort_values("Net", ascending=True)
        fig_bu_aw = px.bar(bu_aw, x="Net", y="BU", orientation="h",
                           text=bu_aw["Net"].apply(fmt_m), color="Net", color_continuous_scale="Blues")
        fig_bu_aw.update_traces(textposition="outside")
        fig_bu_aw.update_layout(height=380, coloraxis_showscale=False, margin=dict(l=10,r=90,t=10,b=10), yaxis_title="")
        st.plotly_chart(fig_bu_aw, use_container_width=True)

    bu_aw_tbl = du_aw.groupby("BU")[["Gross","Net"]].sum().reset_index().sort_values("Net", ascending=False)
    bu_aw_tbl["Gross (M)"] = bu_aw_tbl["Gross"]/1e6
    bu_aw_tbl["Net (M)"]   = bu_aw_tbl["Net"]/1e6
    with st.expander("BU Summary", expanded=True):
        st.dataframe(bu_aw_tbl[["BU","Gross (M)","Net (M)"]]
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)
    du_aw_tbl = du_aw.copy()
    du_aw_tbl["Gross (M)"] = du_aw_tbl["Gross"]/1e6
    du_aw_tbl["Net (M)"]   = du_aw_tbl["Net"]/1e6
    with st.expander("DU Detail"):
        st.dataframe(du_aw_tbl[["BU","DU_Label","Gross (M)","Net (M)"]]
                     .rename(columns={"DU_Label":"Delivery Unit"})
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)

    # Full awarded deals table
    st.markdown("---")
    st.subheader("📋 All Awarded Deals")
    aw_search = st.text_input("Search by account / opportunity", "", key="aw_search")
    disp_aw = aw.copy()
    if aw_search:
        disp_aw = disp_aw[
            disp_aw["Account Name"].str.contains(aw_search, case=False, na=False) |
            disp_aw["Opportunity Name"].str.contains(aw_search, case=False, na=False)]
    aw_cols = ["Year","SNo.","Account Name","Opportunity Name","Stage","Account Manager",
               "Type","Total Gross","Total Net","Project Value",
               "Award Quarter","Contracted","Contract Signed Quarter","ORF Number","Project Duration"]
    st.dataframe(disp_aw[aw_cols].sort_values("Total Net", ascending=False)
                 .style.format({"Total Gross":"{:,.0f}","Total Net":"{:,.0f}","Project Value":"{:,.0f}"}),
                 use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# BOOK3 MAPPING TAB
# ══════════════════════════════════════════════════════════════════════════════
if uploaded_b3 and "🔗 Book3 Mapping" in tab_idx:
  with tabs[tab_idx["🔗 Book3 Mapping"]]:
    with st.spinner("Running mapping..."):
        b3_df   = load_book3(uploaded_b3)
        pipe_df = df_raw if uploaded else None

        aw_parts2 = []
        if uploaded_aw:
            _r = load_awarded(uploaded_aw).copy(); _r["Year"] = "2026"; aw_parts2.append(_r)
        if uploaded_aw25:
            _r = load_awarded(uploaded_aw25).copy(); _r["Year"] = "2025"; aw_parts2.append(_r)
        aw_df2 = pd.concat(aw_parts2, ignore_index=True) if aw_parts2 else None

        map_df = build_book3_mapping(b3_df, pipe_df, aw_df2)

    total_b3  = len(map_df)
    strong_b3 = len(map_df[map_df[["Pipeline Score","Awarded Score"]].max(axis=1) >= 0.70])
    partial_b3= len(map_df[(map_df[["Pipeline Score","Awarded Score"]].max(axis=1) >= 0.55) &
                            (map_df[["Pipeline Score","Awarded Score"]].max(axis=1) <  0.70)])
    nomatch_b3= total_b3 - strong_b3 - partial_b3

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Total Book3 Projects", total_b3)
    k2.metric("🟢 Strong Match",  strong_b3)
    k3.metric("🟡 Partial Match", partial_b3)
    k4.metric("🔴 No Match",      nomatch_b3)

    st.markdown("---")

    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        bu_opts3 = sorted(map_df["Book3 BU"].dropna().unique().tolist())
        sel_bu3  = st.multiselect("Filter by BU", bu_opts3, default=[], key="b3_bu")
    with fc2:
        type_opts3 = sorted(map_df["Book3 Project Type"].dropna().unique().tolist())
        sel_type3  = st.multiselect("Filter by Project Type", type_opts3, default=[], key="b3_type")
    with fc3:
        match_map3 = {"All": None, "🟢 Strong (≥0.70)": "strong", "🟡 Partial (0.55–0.69)": "partial", "🔴 No Match (<0.55)": "nomatch"}
        sel_match3 = st.selectbox("Filter by Match Quality", list(match_map3.keys()), key="b3_match")

    view3 = map_df.copy()
    if sel_bu3:   view3 = view3[view3["Book3 BU"].isin(sel_bu3)]
    if sel_type3: view3 = view3[view3["Book3 Project Type"].isin(sel_type3)]
    sc3 = view3[["Pipeline Score","Awarded Score"]].max(axis=1)
    if match_map3[sel_match3] == "strong":   view3 = view3[sc3 >= 0.70]
    elif match_map3[sel_match3] == "partial":view3 = view3[(sc3 >= 0.55) & (sc3 < 0.70)]
    elif match_map3[sel_match3] == "nomatch":view3 = view3[sc3 < 0.55]

    st.caption(f"Showing {len(view3)} of {total_b3} projects")

    def _color_map_row(row):
        sc = max(row["Pipeline Score"], row["Awarded Score"])
        c = "#E2EFDA" if sc >= 0.70 else ("#FFF2CC" if sc >= 0.55 else "#FFE0E0")
        return [f"background-color: {c}"] * len(row)

    disp_cols3 = [
        "Book3 BU","Book3 Project Type","Book3 Project Name","Book3 Grand Total",
        "Pipeline Match","Pipeline Score","Pipeline Gross (QAR)","Pipeline Net (QAR)","Pipeline Stage","Pipeline AM","Pipeline Quarter",
        "Awarded Match","Awarded Score","Awarded Gross (QAR)","Awarded Net (QAR)","Awarded Stage","Awarded AM","Awarded Year",
    ]
    st.dataframe(
        view3[disp_cols3].style.apply(_color_map_row, axis=1).format({
            "Book3 Grand Total":    "{:,.0f}",
            "Pipeline Score":       "{:.2f}",
            "Pipeline Gross (QAR)": "{:,.0f}",
            "Pipeline Net (QAR)":   "{:,.0f}",
            "Awarded Score":        "{:.2f}",
            "Awarded Gross (QAR)":  "{:,.0f}",
            "Awarded Net (QAR)":    "{:,.0f}",
        }),
        use_container_width=True, height=500, hide_index=True,
    )

    st.markdown("---")
    xl_b3 = _export_mapping_excel(map_df, date.today())
    st.download_button(
        label="⬇️ Export Mapping to Excel",
        data=xl_b3,
        file_name=f"Book3_Mapping_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Sales Weekly Review Dashboard · Upload your Excel files each week to refresh all charts.")
