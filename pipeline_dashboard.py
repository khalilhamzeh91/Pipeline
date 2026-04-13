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
    fp_last_row = 2 + len(full_df)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        F = _xl_formats(wb)

        # SHEET 1 — SUMMARY
        ws = wb.add_worksheet("Summary"); ws.set_zoom(90); ws.set_tab_color("#1a3a6b")
        writer.sheets["Summary"] = ws
        ws.merge_range("A1:C1", f"Weekly Pipeline Summary — {TODAY.strftime('%d %B %Y')}", F["title"]); ws.set_row(0,28)
        _hdr(ws, 2, ["Metric","Value"], [34,22], F["header"])
        _fp = "'Full Pipeline'"; _n = fp_last_row
        kpi_formulas = [
            ("Total Opportunities",         f"=COUNTA({_fp}!C3:C{_n})",                                        F["kpi_val"]),
            ("Total Gross Pipeline (QAR)",  f"=SUM({_fp}!I3:I{_n})",                                           F["kpi_val"]),
            ("Total Net Pipeline (QAR)",    f"=SUM({_fp}!J3:J{_n})",                                           F["kpi_val"]),
            ("Forecasted Gross (QAR)",      f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!I3:I{_n})',                   F["kpi_val"]),
            ("Forecasted Net (QAR)",        f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!J3:J{_n})',                   F["kpi_val"]),
            ("Strategic Opportunities",     f'=COUNTIF({_fp}!M3:M{_n},"Yes")',                                  F["kpi_val"]),
            ("Overdue Deals",               f'=COUNTIF({_fp}!Q3:Q{_n},"YES")',                                  F["kpi_val"]),
        ]
        for i, (label, formula, fmt) in enumerate(kpi_formulas):
            ws.write(3+i, 0, label, F["kpi_lbl"]); ws.write_formula(3+i, 1, formula, fmt)
        ws.merge_range("D1:H1", "Pipeline by Stage", F["title"])
        _hdr(ws, 2, ["Stage","Count","Gross (QAR)","Net (QAR)"], None, F["header"])
        ws.set_column(3,3,22); ws.set_column(4,4,8); ws.set_column(5,5,18); ws.set_column(6,6,18)
        for i, row in stage_df.reset_index(drop=True).iterrows():
            xl1 = 4 + i
            alt = i%2==1
            ws.write(3+i,3,row["Stage"],F["alt"] if alt else F["text"])
            ws.write_formula(3+i,4,f'=COUNTIF({_fp}!D3:D{_n},D{xl1})',F["alt_num"] if alt else F["num"])
            ws.write_formula(3+i,5,f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!I3:I{_n})',F["alt_num"] if alt else F["num"])
            ws.write_formula(3+i,6,f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!J3:J{_n})',F["alt_num"] if alt else F["num"])
        _sl = 3 + len(stage_df)
        ws.write(3+len(stage_df),3,"TOTAL",F["tot_lbl"])
        ws.write_formula(3+len(stage_df),4,f"=SUM(E4:E{_sl})",F["tot_num"])
        ws.write_formula(3+len(stage_df),5,f"=SUM(F4:F{_sl})",F["tot_num"])
        ws.write_formula(3+len(stage_df),6,f"=SUM(G4:G{_sl})",F["tot_num"])

        # SHEET 2 — DU BREAKDOWN
        du_ws = wb.add_worksheet("DU Breakdown"); du_ws.set_zoom(90); du_ws.set_tab_color("#FF8C00")
        writer.sheets["DU Breakdown"] = du_ws
        du_ws.merge_range("A1:G1","Gross & Net Breakdown by BU / Delivery Unit",F["title"]); du_ws.set_row(0,28)
        _hdr(du_ws,1,["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[42,52,20,20,22],F["header"])
        # Two-pass: pre-compute layout then write with SUM formulas
        _du_layout = []
        _pos = 0
        for bu_name, bu_grp in du_totals.groupby("BU"):
            bu_r0 = 2 + _pos; _pos += 1
            du_list = []
            for _, drow in bu_grp.iterrows():
                du_r0 = 2 + _pos; _pos += 1
                deals = du_exp[du_exp["DU"] == drow["DU"]]
                opp_rows = []
                for _, deal in deals.iterrows():
                    opp_rows.append((2 + _pos, deal)); _pos += 1
                du_list.append((drow, du_r0, opp_rows))
            _du_layout.append((bu_name, bu_r0, du_list))
        _gt_r0 = 2 + _pos
        for bu_name, bu_r0, du_list in _du_layout:
            du_ws.write(bu_r0,0,bu_name,F["bu_lbl"]); du_ws.write(bu_r0,1,"",F["bu_lbl"])
            _du_rows = [dr0+1 for (_, dr0, _) in du_list]
            _buC = "+".join("C"+str(r) for r in _du_rows)
            _buD = "+".join("D"+str(r) for r in _du_rows)
            _buE = "+".join("E"+str(r) for r in _du_rows)
            du_ws.write_formula(bu_r0,2,"="+_buC,F["bu_num"]); du_ws.write_formula(bu_r0,3,"="+_buD,F["bu_num"]); du_ws.write_formula(bu_r0,4,"="+_buE,F["bu_num"])
            for drow, du_r0, opp_rows in du_list:
                alt=(du_r0-2)%2==1
                du_ws.write(du_r0,0,"",F["alt"] if alt else F["text"]); du_ws.write(du_r0,1,drow["DU"],F["alt"] if alt else F["text"])
                if opp_rows:
                    _r1=min(r for r,_ in opp_rows)+1; _rN=max(r for r,_ in opp_rows)+1
                    du_ws.write_formula(du_r0,2,f"=SUM(C{_r1}:C{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,3,f"=SUM(D{_r1}:D{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,4,f"=SUM(E{_r1}:E{_rN})",F["alt_num"] if alt else F["num"])
                else:
                    du_ws.write_number(du_r0,2,0,F["alt_num"] if alt else F["num"]); du_ws.write_number(du_r0,3,0,F["alt_num"] if alt else F["num"]); du_ws.write_number(du_r0,4,0,F["alt_num"] if alt else F["num"])
                for opp_r0, deal in opp_rows:
                    fore_net=deal["Net"] if str(deal.get("Forecasted","")).strip()=="Yes" else 0
                    du_ws.write(opp_r0,0,"",F["opp"]); du_ws.write(opp_r0,1,f"  ↳  {deal['Lead/Opp Name']}",F["opp"])
                    du_ws.write_number(opp_r0,2,deal["Gross"],F["opp_num"]); du_ws.write_number(opp_r0,3,deal["Net"],F["opp_num"]); du_ws.write_number(opp_r0,4,fore_net,F["opp_num"])
        t=_gt_r0
        du_ws.write(t,0,"GRAND TOTAL",F["tot_lbl"]); du_ws.write(t,1,"",F["tot_lbl"])
        _bu_rows = [br0+1 for (_, br0, _) in _du_layout]
        _gtC = "+".join("C"+str(r) for r in _bu_rows)
        _gtD = "+".join("D"+str(r) for r in _bu_rows)
        _gtE = "+".join("E"+str(r) for r in _bu_rows)
        du_ws.write_formula(t,2,"="+_gtC,F["tot_num"]); du_ws.write_formula(t,3,"="+_gtD,F["tot_num"]); du_ws.write_formula(t,4,"="+_gtE,F["tot_num"])
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
        # Two-pass: pre-compute layout for BU subtotal formulas
        _fds_layout = []
        _fds_pos = 0
        for bu_name, bu_grp in fore_du_summary.groupby("BU", sort=False):
            bu_r0 = 4 + _fds_pos; _fds_pos += 1
            detail_rows = []
            for _, drow in bu_grp.iterrows():
                detail_rows.append(4 + _fds_pos); _fds_pos += 1
            _fds_layout.append((bu_name, bu_grp, bu_r0, detail_rows))
        ts = 4 + _fds_pos
        for bu_name, bu_grp, bu_r0, detail_rows in _fds_layout:
            fd_ws.write(bu_r0,0,bu_name,F["bu_lbl"]); fd_ws.write(bu_r0,1,f"Total: {len(bu_grp)} rows",F["bu_lbl"]); fd_ws.write(bu_r0,2,"",F["bu_lbl"])
            if detail_rows:
                _r1=min(detail_rows)+1; _rN=max(detail_rows)+1
                fd_ws.write_formula(bu_r0,3,f"=SUM(D{_r1}:D{_rN})",F["bu_num"])
                fd_ws.write_formula(bu_r0,4,f"=SUM(E{_r1}:E{_rN})",F["bu_num"])
                fd_ws.write_formula(bu_r0,5,f"=SUM(F{_r1}:F{_rN})",F["bu_num"])
            for dr0, (_, row) in zip(detail_rows, bu_grp.iterrows()):
                alt=(dr0-4)%2==1
                fd_ws.write(dr0,0,"",F["alt"] if alt else F["text"]); fd_ws.write(dr0,1,row["DU"],F["alt"] if alt else F["text"]); fd_ws.write(dr0,2,row["Closure Due Quarter"],F["alt"] if alt else F["text"])
                fd_ws.write_number(dr0,3,row["Count"],F["alt_num"] if alt else F["num"]); fd_ws.write_number(dr0,4,row["Gross"],F["alt_num"] if alt else F["num"]); fd_ws.write_number(dr0,5,row["Net"],F["alt_num"] if alt else F["num"])
        fd_ws.write(ts,0,"GRAND TOTAL",F["tot_lbl"])
        for _c in range(1,3): fd_ws.write(ts,_c,"",F["tot_lbl"])
        _bu_r_nums = [br0+1 for (_, _, br0, _) in _fds_layout]
        _gtD = "+".join("D"+str(r) for r in _bu_r_nums)
        _gtE = "+".join("E"+str(r) for r in _bu_r_nums)
        _gtF = "+".join("F"+str(r) for r in _bu_r_nums)
        fd_ws.write_formula(ts,3,"="+_gtD,F["tot_num"]); fd_ws.write_formula(ts,4,"="+_gtE,F["tot_num"]); fd_ws.write_formula(ts,5,"="+_gtF,F["tot_num"])
        det_start=ts+2
        fd_ws.merge_range(det_start,0,det_start,9,"Detail: Forecasted Deals by BU / DU",F["sec_hdr"]); fd_ws.set_row(det_start,22)
        _hdr(fd_ws,det_start+1,["BU","Delivery Unit","Account Name","Opportunity","Stage","Account Manager","Quarter","Gross (QAR)","Net (QAR)","Win Prob"],[42,38,28,36,22,24,10,18,18,12],F["header"])
        # Two-pass for detail section
        _fdd_layout = []
        _fdd_pos = 0
        for bu_name, bu_grp in fore_du_detail.groupby("BU", sort=False):
            bu_r0 = det_start + 2 + _fdd_pos; _fdd_pos += 1
            detail_rows = []
            for _, drow in bu_grp.iterrows():
                detail_rows.append(det_start + 2 + _fdd_pos); _fdd_pos += 1
            _fdd_layout.append((bu_name, bu_grp, bu_r0, detail_rows))
        td = det_start + 2 + _fdd_pos
        for bu_name, bu_grp, bu_r0, detail_rows in _fdd_layout:
            fd_ws.write(bu_r0,0,bu_name,F["bu_lbl"])
            for _cc in range(1,10): fd_ws.write(bu_r0,_cc,"",F["bu_lbl"])
            if detail_rows:
                _r1=min(detail_rows)+1; _rN=max(detail_rows)+1
                fd_ws.write_formula(bu_r0,7,f"=SUM(H{_r1}:H{_rN})",F["bu_num"])
                fd_ws.write_formula(bu_r0,8,f"=SUM(I{_r1}:I{_rN})",F["bu_num"])
            for dr0, (_, row) in zip(detail_rows, bu_grp.iterrows()):
                alt=(dr0-det_start-2)%2==1; f_t=F["grn"] if not alt else F["alt"]; f_n=F["grn_num"] if not alt else F["alt_num"]
                fd_ws.write(dr0,0,"",f_t); fd_ws.write(dr0,1,str(row["DU"]) if pd.notna(row["DU"]) else "",f_t); fd_ws.write(dr0,2,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",f_t); fd_ws.write(dr0,3,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",f_t); fd_ws.write(dr0,4,str(row["Stage"]) if pd.notna(row["Stage"]) else "",f_t); fd_ws.write(dr0,5,str(row["Account Manager"]) if pd.notna(row["Account Manager"]) else "",f_t); fd_ws.write(dr0,6,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",f_t)
                fd_ws.write_number(dr0,7,row["Gross"],f_n); fd_ws.write_number(dr0,8,row["Net"],f_n); fd_ws.write(dr0,9,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",f_t)
        fd_ws.write(td,0,"GRAND TOTAL",F["tot_lbl"])
        for _c in range(1,7): fd_ws.write(td,_c,"",F["tot_lbl"])
        _bu_det_nums = [br0+1 for (_, _, br0, _) in _fdd_layout]
        _gtH = "+".join("H"+str(r) for r in _bu_det_nums)
        _gtI = "+".join("I"+str(r) for r in _bu_det_nums)
        fd_ws.write_formula(td,7,"="+_gtH,F["tot_num"]); fd_ws.write_formula(td,8,"="+_gtI,F["tot_num"]); fd_ws.write(td,9,"",F["tot_lbl"])

        # SHEET 4 — SECTOR & AM
        sa_ws = wb.add_worksheet("Sector & AM"); sa_ws.set_zoom(90); sa_ws.set_tab_color("#228B22")
        writer.sheets["Sector & AM"] = sa_ws
        sa_ws.merge_range("A1:E1","Pipeline by Sector",F["title"]); sa_ws.set_row(0,28)
        _hdr(sa_ws,1,["Sector","Count","Gross (QAR)","Net (QAR)"],[24,8,20,20],F["header"])
        for i, row in sector_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            sa_ws.write(2+i,0,row["Sector"],F["alt"] if alt else F["text"]); sa_ws.write_number(2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(2+i,3,row["Net"],F["alt_num"] if alt else F["num"])
        _sec_t = 2 + len(sector_df)
        sa_ws.write(_sec_t,0,"TOTAL",F["tot_lbl"])
        sa_ws.write_formula(_sec_t,1,f"=SUM(B3:B{_sec_t})",F["tot_num"])
        sa_ws.write_formula(_sec_t,2,f"=SUM(C3:C{_sec_t})",F["tot_num"])
        sa_ws.write_formula(_sec_t,3,f"=SUM(D3:D{_sec_t})",F["tot_num"])
        off=2+len(sector_df)+1+2
        sa_ws.merge_range(off,0,off,4,"Pipeline by Account Manager",F["title"])
        _hdr(sa_ws,off+1,["Account Manager","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[28,8,20,20,22],F["header"])
        for i, row in am_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            sa_ws.write(off+2+i,0,row["Account Manager"],F["alt"] if alt else F["text"]); sa_ws.write_number(off+2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,3,row["Net"],F["alt_num"] if alt else F["num"]); sa_ws.write_number(off+2+i,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])
        _am_t = off + 2 + len(am_df)
        _am_r1 = off + 3
        sa_ws.write(_am_t,0,"TOTAL",F["tot_lbl"])
        sa_ws.write_formula(_am_t,1,f"=SUM(B{_am_r1}:B{_am_t})",F["tot_num"])
        sa_ws.write_formula(_am_t,2,f"=SUM(C{_am_r1}:C{_am_t})",F["tot_num"])
        sa_ws.write_formula(_am_t,3,f"=SUM(D{_am_r1}:D{_am_t})",F["tot_num"])
        sa_ws.write_formula(_am_t,4,f"=SUM(E{_am_r1}:E{_am_t})",F["tot_num"])

        # SHEET 5 — QUARTERLY & PROBABILITY
        qp_ws = wb.add_worksheet("Quarterly & Probability"); qp_ws.set_zoom(90); qp_ws.set_tab_color("#DAA520")
        writer.sheets["Quarterly & Probability"] = qp_ws
        qp_ws.merge_range("A1:E1","Quarterly Close Plan",F["title"]); qp_ws.set_row(0,28)
        _hdr(qp_ws,1,["Quarter","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[12,8,20,20,22],F["header"])
        for i, row in q_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            qp_ws.write(2+i,0,row["Closure Due Quarter"],F["alt"] if alt else F["text"]); qp_ws.write_number(2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,3,row["Net"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(2+i,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])
        _q_t = 2 + len(q_df)
        qp_ws.write(_q_t,0,"TOTAL",F["tot_lbl"])
        qp_ws.write_formula(_q_t,1,f"=SUM(B3:B{_q_t})",F["tot_num"])
        qp_ws.write_formula(_q_t,2,f"=SUM(C3:C{_q_t})",F["tot_num"])
        qp_ws.write_formula(_q_t,3,f"=SUM(D3:D{_q_t})",F["tot_num"])
        qp_ws.write_formula(_q_t,4,f"=SUM(E3:E{_q_t})",F["tot_num"])
        off2=2+len(q_df)+1+2
        qp_ws.merge_range(off2,0,off2,3,"Pipeline by Winning Probability",F["title"])
        _hdr(qp_ws,off2+1,["Winning Probability","Count","Gross (QAR)","Net (QAR)"],[22,8,20,20],F["header"])
        for i, row in prob_df.reset_index(drop=True).iterrows():
            alt=i%2==1
            qp_ws.write(off2+2+i,0,row["Winning Probability"],F["alt"] if alt else F["text"]); qp_ws.write_number(off2+2+i,1,row["Count"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(off2+2+i,2,row["Gross"],F["alt_num"] if alt else F["num"]); qp_ws.write_number(off2+2+i,3,row["Net"],F["alt_num"] if alt else F["num"])
        _pb_t = off2 + 2 + len(prob_df)
        _pb_r1 = off2 + 3
        qp_ws.write(_pb_t,0,"TOTAL",F["tot_lbl"])
        qp_ws.write_formula(_pb_t,1,f"=SUM(B{_pb_r1}:B{_pb_t})",F["tot_num"])
        qp_ws.write_formula(_pb_t,2,f"=SUM(C{_pb_r1}:C{_pb_t})",F["tot_num"])
        qp_ws.write_formula(_pb_t,3,f"=SUM(D{_pb_r1}:D{_pb_t})",F["tot_num"])

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
        fw.write(t2,0,"TOTAL",F["tot_lbl"])
        fw.write_formula(t2,5,f"=SUM(F3:F{t2})",F["tot_num"]); fw.write_formula(t2,6,f"=SUM(G3:G{t2})",F["tot_num"])

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
        _ov_t = 2 + len(overdue_df)
        ow.write(_ov_t,0,"TOTAL",F["tot_lbl"])
        ow.write_formula(_ov_t,4,f"=SUM(E3:E{_ov_t})",F["tot_num"])

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
        pw.write(fp_last_row,0,"TOTAL",F["tot_lbl"])
        pw.write_formula(fp_last_row,8,f"=SUM(I3:I{fp_last_row})",F["tot_num"])
        pw.write_formula(fp_last_row,9,f"=SUM(J3:J{fp_last_row})",F["tot_num"])

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
        # Pre-compute Excel row range per deal for SUM formulas
        bd_deal_rows={}
        for rp,(_, row) in enumerate(bd_exp.iterrows()):
            _di=row["_deal_idx"]; _xr=2+rp
            if _di not in bd_deal_rows: bd_deal_rows[_di]=[_xr,_xr]
            else: bd_deal_rows[_di][1]=_xr
        bd_gcol=bd_cmap["Gross (breakdown)"]; bd_ncol=bd_cmap["Net (breakdown)"]
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
            _bws(bd_cmap["SNo."],              row.get("SNo.") if isf else None,ft)
            _bws(bd_cmap["Account Name"],       row.get("Account Name"),ft)
            _bws(bd_cmap["Lead/Opp Name"],      row.get("Lead/Opp Name"),ft)
            _bws(bd_cmap["BU"],row.get("BU_exp"),fbu); _bws(bd_cmap["DU"],row.get("DU_exp"),fdu)
            _bwn(bd_cmap["Gross (breakdown)"],  row.get("Gross_exp"),fxn)
            _bwn(bd_cmap["Net (breakdown)"],    row.get("Net_exp"),  fxn)
            if isf:
                r0,r1=bd_deal_rows[didx]; gc=chr(65+bd_gcol); nc=chr(65+bd_ncol)
                bw.write_formula(xl_r,bd_cmap["Total Gross"],f"=SUM({gc}{r0+1}:{gc}{r1+1})",bd_ftot)
                bw.write_formula(xl_r,bd_cmap["Total Net"],  f"=SUM({nc}{r0+1}:{nc}{r1+1})",bd_ftot)
            else:
                bw.write_blank(xl_r,bd_cmap["Total Gross"],None,bd_ftotbl)
                bw.write_blank(xl_r,bd_cmap["Total Net"],  None,bd_ftotbl)
            _bws(bd_cmap["Stage"],              row.get("Stage"),ft)
            _bws(bd_cmap["Account Manager"],    row.get("Account Manager"),ft)
            _bws(bd_cmap["Sector"],             row.get("Sector"),ft)
            _bws(bd_cmap["Closure Due Quarter"],row.get("Closure Due Quarter"),ft)
            _bws(bd_cmap["Winning Probability"],row.get("Winning Probability"),ft)
            _bws(bd_cmap["Forecasted"],         row.get("Forecasted"),ft)
            _bws(bd_cmap["Strategic Opportunity"],row.get("Strategic Opportunity"),ft)
            cd_i=bd_cmap["Est. Close Date"]
            cd_v=row.get("Est. Close Date")
            if pd.notna(cd_v):
                try: bw.write_datetime(xl_r,cd_i,pd.Timestamp(cd_v).to_pydatetime(),bd_fdate)
                except Exception: bw.write(xl_r,cd_i,str(cd_v),bd_fdate)
            else: bw.write_blank(xl_r,cd_i,None,bd_fdate if isf else bd_fdatebl)

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
        ws.write(tr,3,"TOTAL",F["tot_lbl"])
        ws.write_formula(tr,4,f"=SUM(E4:E{tr})",F["tot_num"]); ws.write_formula(tr,5,f"=SUM(F4:F{tr})",F["tot_num"]); ws.write_formula(tr,6,f"=SUM(G4:G{tr})",F["tot_num"]); ws.write_formula(tr,7,f"=SUM(H4:H{tr})",F["tot_num"])
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
        # Two-pass awarded DU breakdown
        _adu_layout = []
        _apos = 0
        for bu_name, bu_grp in du_totals.groupby("BU"):
            bu_r0 = 2 + _apos; _apos += 1
            du_list = []
            for _, drow in bu_grp.iterrows():
                du_r0 = 2 + _apos; _apos += 1
                deals = du_exp[du_exp["DU"] == drow["DU"]]
                opp_rows = []
                for _, deal in deals.iterrows():
                    opp_rows.append((2 + _apos, deal)); _apos += 1
                du_list.append((drow, du_r0, opp_rows))
            _adu_layout.append((bu_name, bu_r0, du_list))
        _agt_r0 = 2 + _apos
        for bu_name, bu_r0, du_list in _adu_layout:
            du_ws.write(bu_r0,0,bu_name,F["bu_lbl"]); du_ws.write(bu_r0,1,"",F["bu_lbl"])
            _du_rows = [dr0+1 for (_, dr0, _) in du_list]
            _buC = "+".join("C"+str(r) for r in _du_rows); _buD = "+".join("D"+str(r) for r in _du_rows)
            _buE = "+".join("E"+str(r) for r in _du_rows); _buF = "+".join("F"+str(r) for r in _du_rows)
            du_ws.write_formula(bu_r0,2,"="+_buC,F["bu_num"]); du_ws.write_formula(bu_r0,3,"="+_buD,F["bu_num"]); du_ws.write_formula(bu_r0,4,"="+_buE,F["bu_num"]); du_ws.write_formula(bu_r0,5,"="+_buF,F["bu_num"])
            for drow, du_r0, opp_rows in du_list:
                alt=(du_r0-2)%2==1
                du_ws.write(du_r0,0,"",F["alt"] if alt else F["text"]); du_ws.write(du_r0,1,drow["DU"],F["alt"] if alt else F["text"])
                if opp_rows:
                    _r1=min(r for r,_ in opp_rows)+1; _rN=max(r for r,_ in opp_rows)+1
                    du_ws.write_formula(du_r0,2,f"=SUM(C{_r1}:C{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,3,f"=SUM(D{_r1}:D{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,4,f"=SUM(E{_r1}:E{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,5,f"=SUM(F{_r1}:F{_rN})",F["alt_num"] if alt else F["num"])
                else:
                    for _c in range(2,6): du_ws.write_number(du_r0,_c,0,F["alt_num"] if alt else F["num"])
                for opp_r0, deal in opp_rows:
                    n25=deal["Net"] if deal["Year"]=="2025" else 0; n26=deal["Net"] if deal["Year"]=="2026" else 0
                    du_ws.write(opp_r0,0,"",F["opp"]); du_ws.write(opp_r0,1,f"  ↳  {deal['Opportunity']}",F["opp"])
                    du_ws.write_number(opp_r0,2,deal["Gross"],F["opp_num"]); du_ws.write_number(opp_r0,3,deal["Net"],F["opp_num"]); du_ws.write_number(opp_r0,4,n25,F["opp_num"]); du_ws.write_number(opp_r0,5,n26,F["opp_num"])
        t=_agt_r0
        du_ws.write(t,0,"GRAND TOTAL",F["tot_lbl"]); du_ws.write(t,1,"",F["tot_lbl"])
        _abu_rows = [br0+1 for (_, br0, _) in _adu_layout]
        _agtC = "+".join("C"+str(r) for r in _abu_rows); _agtD = "+".join("D"+str(r) for r in _abu_rows)
        _agtE = "+".join("E"+str(r) for r in _abu_rows); _agtF = "+".join("F"+str(r) for r in _abu_rows)
        du_ws.write_formula(t,2,"="+_agtC,F["tot_num"]); du_ws.write_formula(t,3,"="+_agtD,F["tot_num"]); du_ws.write_formula(t,4,"="+_agtE,F["tot_num"]); du_ws.write_formula(t,5,"="+_agtF,F["tot_num"])

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
        am_ws.write(tr,0,"GRAND TOTAL",F["tot_lbl"])
        am_ws.write_formula(tr,1,f"=SUM(B3:B{tr})",F["tot_num"]); am_ws.write_formula(tr,2,f"=SUM(C3:C{tr})",F["tot_num"]); am_ws.write_formula(tr,3,f"=SUM(D3:D{tr})",F["tot_num"]); am_ws.write_formula(tr,4,f"=SUM(E3:E{tr})",F["tot_num"]); am_ws.write_formula(tr,5,f"=SUM(F3:F{tr})",F["tot_num"])

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
        aq_ws.write(tr,0,"TOTAL",F["tot_lbl"]); aq_ws.write(tr,1,"",F["tot_lbl"])
        aq_ws.write_formula(tr,2,f"=SUM(C3:C{tr})",F["tot_num"]); aq_ws.write_formula(tr,3,f"=SUM(D3:D{tr})",F["tot_num"]); aq_ws.write_formula(tr,4,f"=SUM(E3:E{tr})",F["tot_num"])
        nr_off=tr+2; aq_ws.merge_range(nr_off,0,nr_off,4,"New vs Renew Breakdown",F["title"])
        for c, col in enumerate(["Year","Type","Deals","Gross (QAR)","Net (QAR)"]): aq_ws.write(nr_off+1,c,col,F["header"])
        for i, row in nr_df.reset_index(drop=True).iterrows():
            ft=F["y25"] if row["Year"]=="2025" else F["y26"]; fn=F["y25_num"] if row["Year"]=="2025" else F["y26_num"]
            aq_ws.write(nr_off+2+i,0,row["Year"],ft); aq_ws.write(nr_off+2+i,1,row["Type"],ft); aq_ws.write_number(nr_off+2+i,2,row["Count"],fn); aq_ws.write_number(nr_off+2+i,3,row["Gross"],fn); aq_ws.write_number(nr_off+2+i,4,row["Net"],fn)
        _nr_t = nr_off + 2 + len(nr_df); _nr_r1 = nr_off + 3
        aq_ws.write(_nr_t,0,"TOTAL",F["tot_lbl"]); aq_ws.write(_nr_t,1,"",F["tot_lbl"])
        aq_ws.write_formula(_nr_t,2,f"=SUM(C{_nr_r1}:C{_nr_t})",F["tot_num"]); aq_ws.write_formula(_nr_t,3,f"=SUM(D{_nr_r1}:D{_nr_t})",F["tot_num"]); aq_ws.write_formula(_nr_t,4,f"=SUM(E{_nr_r1}:E{_nr_t})",F["tot_num"])

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
        # Pre-compute Excel row range per deal for SUM formulas
        aw_deal_rows={}
        for rp,(_, row) in enumerate(aw_exp.iterrows()):
            _di=row["_deal_idx"]; _xr=2+rp
            if _di not in aw_deal_rows: aw_deal_rows[_di]=[_xr,_xr]
            else: aw_deal_rows[_di][1]=_xr
        aw_gcol=aw_cm["Gross (breakdown)"]; aw_ncol=aw_cm["Net (breakdown)"]
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
            _aws(aw_cm["SNo."],             row.get("SNo.") if isf else None,ft)
            _aws(aw_cm["Account Name"],      row.get("Account Name"),ft)
            _aws(aw_cm["Opportunity Name"],  row.get("Opportunity Name"),ft)
            _aws(aw_cm["BU"],row.get("BU_exp"),fbu); _aws(aw_cm["DU"],row.get("DU_exp"),fdu)
            _awn(aw_cm["Gross (breakdown)"], row.get("Gross_exp"),fxn)
            _awn(aw_cm["Net (breakdown)"],   row.get("Net_exp"),  fxn)
            if isf:
                r0,r1=aw_deal_rows[didx]; gc=chr(65+aw_gcol); nc=chr(65+aw_ncol)
                aw_bw.write_formula(xl_r,aw_cm["Total Gross"],f"=SUM({gc}{r0+1}:{gc}{r1+1})",aw_ftot)
                aw_bw.write_formula(xl_r,aw_cm["Total Net"],  f"=SUM({nc}{r0+1}:{nc}{r1+1})",aw_ftot)
            else:
                aw_bw.write_blank(xl_r,aw_cm["Total Gross"],None,aw_ftotbl)
                aw_bw.write_blank(xl_r,aw_cm["Total Net"],  None,aw_ftotbl)
            _aws(aw_cm["Stage"],         row.get("Stage"),ft)
            _aws(aw_cm["Account Manager"],row.get("Account Manager"),ft)
            _aws(aw_cm["Award Quarter"], row.get("Award Quarter"),ft)
            _aws(aw_cm["Contracted"],    row.get("Contracted"),ft)
            _aws(aw_cm["Year"],          row.get("Year"),ft)

    output.seek(0)
    return output.read()


# ── AM PIPELINE HELPERS ────────────────────────────────────────────────────────
MONTHS_AM = ["January","February","March","April","May","June",
             "July","August","September","October","November","December"]

def _clean_am_list(value):
    if pd.isna(value) or str(value).strip() in ("","nan"): return []
    seen = {}
    for name in str(value).split("\n"):
        name = name.strip()
        if not name: continue
        # Normalize known names
        if "khalil" in name.lower():  name = "Khalil Hamzeh"
        elif "yazan" in name.lower(): name = "Yazan Al Razem"
        seen[name] = None
    return list(seen.keys())

@st.cache_data
def load_am_pipeline(file):
    xl = pd.ExcelFile(file)
    sheet = "Export" if "Export" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(file, sheet_name=sheet)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    df["Total Gross"]     = pd.to_numeric(df.get("Total Gross",    0), errors="coerce").fillna(0)
    df["Total Net"]       = pd.to_numeric(df.get("Total Net",      0), errors="coerce").fillna(0)
    df["Est. Close Date"] = pd.to_datetime(df.get("Est. Close Date"), errors="coerce")
    df["Overdue"]         = df["Est. Close Date"] < pd.Timestamp(date.today())
    STAGE_SHORT_AM = {
        "Stage 1: Assessment & Qualification":                   "S1 - Assessment",
        "Stage 2: Discovery & Scoping":                          "S2 - Discovery",
        "Stage 3.1: RFP & BID Qualification":                   "S3.1 - RFP",
        "Stage 3.2: Solution Development & Proposal Submission": "S3.2 - Solution Dev",
        "Stage 4: Technical Evaluation By Customer":             "S4 - Tech Eval",
        "Stage 5: Resolution/Financial Negotiation":             "S5 - Negotiation",
    }
    df["Stage_Short"] = df["Stage"].map(STAGE_SHORT_AM).fillna(df["Stage"])
    for m in MONTHS_AM:
        df[m] = pd.to_numeric(df.get(m, 0), errors="coerce").fillna(0)
    # Normalize column names so _expand_deals can find them (AM file uses capital B)
    if "Gross (Breakdown)" in df.columns:
        df = df.rename(columns={"Gross (Breakdown)": "Gross (breakdown)"})
    if "Net (Breakdown)" in df.columns:
        df = df.rename(columns={"Net (Breakdown)": "Net (breakdown)"})
    # Normalize AM names in Capability Sales at load time so ALL downstream
    # code (Full Pipeline sheet, UI display, AM explosion) sees clean names
    if "Capability Sales" in df.columns:
        df["Capability Sales"] = df["Capability Sales"].apply(
            lambda v: "\n".join(_clean_am_list(v)) if pd.notna(v) else v
        )
    return df

@st.cache_data
def export_am_pipeline_excel(file):
    TODAY   = date.today()
    mapping = load_coa()
    df      = load_am_pipeline(file)

    # Explode by AM
    am_rows = []
    for deal_idx, row in df.iterrows():
        ams = _clean_am_list(row.get("Capability Sales",""))
        if not ams: ams = ["Unassigned"]
        for i, am in enumerate(ams):
            nr = {c: row[c] for c in df.columns}
            nr["AM_exp"] = am; nr["_is_first"] = (i == 0)
            nr["_am_count"] = len(ams); nr["_deal_idx"] = deal_idx
            am_rows.append(nr)
    am_exp = pd.DataFrame(am_rows)

    # Explode by DU
    du_rows_am = []
    for _, row in df.iterrows():
        dus   = str(row.get("DU","")).split("\n")   if pd.notna(row.get("DU")) else ["Unknown"]
        gross = str(row.get("Gross (breakdown)","")).replace(",","").split("\n") if pd.notna(row.get("Gross (breakdown)")) else ["0"]
        net   = str(row.get("Net (breakdown)","")).replace(",","").split("\n")   if pd.notna(row.get("Net (breakdown)"))   else ["0"]
        n = max(len(dus), len(gross), len(net))
        for i in range(n):
            du = dus[i].strip() if i < len(dus) else dus[-1].strip()
            try: g_val = float(gross[i].strip() if i < len(gross) else "0")
            except: g_val = 0.0
            try: n_val = float(net[i].strip()   if i < len(net)   else "0")
            except: n_val = 0.0
            du_rows_am.append({"BU": du_to_bu(du, mapping), "DU": du, "Gross": g_val, "Net": n_val,
                               "Forecasted": str(row.get("Forecasted","")).strip(),
                               "Stage": row.get("Stage",""), "Sector": row.get("Sector",""),
                               "Closure Due Quarter": row.get("Closure Due Quarter",""),
                               "Account Name": row.get("Account Name",""),
                               "Lead/Opp Name": row.get("Lead/Opp Name",""),
                               "Winning Probability": row.get("Winning Probability","")})
    du_exp_am = pd.DataFrame(du_rows_am)

    # Summary tables
    STAGE_SHORT_AM = {"Stage 1: Assessment & Qualification":"S1 - Assessment",
                      "Stage 2: Discovery & Scoping":"S2 - Discovery",
                      "Stage 3.1: RFP & BID Qualification":"S3.1 - RFP",
                      "Stage 3.2: Solution Development & Proposal Submission":"S3.2 - Solution Dev",
                      "Stage 4: Technical Evaluation By Customer":"S4 - Tech Eval",
                      "Stage 5: Resolution/Financial Negotiation":"S5 - Negotiation"}
    stage_df  = df.groupby("Stage_Short").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().rename(columns={"Stage_Short":"Stage"})
    stage_df["_ord"] = stage_df["Stage"].map({s:i for i,s in enumerate(STAGE_SHORT_AM.values())})
    stage_df = stage_df.sort_values("_ord").drop(columns="_ord")
    am_agg = (am_exp.groupby("AM_exp").agg(Count=("Lead/Opp Name","nunique"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().rename(columns={"AM_exp":"Capability Sales"}).sort_values("Net",ascending=False))
    fore_am_agg = am_exp[am_exp["Forecasted"]=="Yes"].groupby("AM_exp")["Total Net"].sum().reset_index().rename(columns={"AM_exp":"Capability Sales","Total Net":"Forecasted Net"})
    am_agg = am_agg.merge(fore_am_agg,on="Capability Sales",how="left").fillna({"Forecasted Net":0})
    q_df = df.groupby("Closure Due Quarter").agg(Count=("Lead/Opp Name","count"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Closure Due Quarter")
    fore_q = df[df["Forecasted"]=="Yes"].groupby("Closure Due Quarter")["Total Net"].sum().reset_index().rename(columns={"Total Net":"Forecasted Net"})
    q_df = q_df.merge(fore_q,on="Closure Due Quarter",how="left").fillna({"Forecasted Net":0})
    du_totals_am = du_exp_am.groupby(["BU","DU"])[["Gross","Net"]].sum().reset_index().sort_values(["BU","Net"],ascending=[True,False])
    fore_du_am = du_exp_am[du_exp_am["Forecasted"]=="Yes"].groupby("DU")["Net"].sum().reset_index().rename(columns={"Net":"Forecasted Net"})
    du_totals_am = du_totals_am.merge(fore_du_am,on="DU",how="left").fillna({"Forecasted Net":0})
    full_df_am = df[["SNo.","Account Name","Lead/Opp Name","Stage_Short","Capability Sales","Sector","BU","DU","Total Gross","Total Net","Winning Probability","Forecasted","Strategic Opportunity","Closure Due Quarter","Est. Close Date","Source of Opportunity","Overdue"]].sort_values("Total Net",ascending=False).rename(columns={"Stage_Short":"Stage"})
    fp_last_row_am = 2 + len(full_df_am)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        F = _xl_formats(wb)
        # Extra formats for monthly columns
        fmt_month     = wb.add_format({"num_format":"#,##0","border":1,"align":"right","bg_color":"#EFF7F0"})
        fmt_month_hdr = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#2E7D32","border":1,"align":"center","text_wrap":True})
        fmt_month_tot = wb.add_format({"bold":True,"bg_color":"#1B5E20","font_color":"#FFFFFF","num_format":"#,##0","border":1,"align":"right"})
        fmt_bu_hdr2   = wb.add_format({"bold":True,"bg_color":"#2E5FA3","font_color":"#FFFFFF","border":1,"align":"left","font_size":11})

        # SHEET 1 — SUMMARY
        ws = wb.add_worksheet("Summary"); ws.set_zoom(90); ws.set_tab_color("#1a3a6b")
        writer.sheets["Summary"] = ws
        ws.merge_range("A1:C1",f"AM Pipeline Summary — {TODAY.strftime('%d %B %Y')}",F["title"]); ws.set_row(0,28)
        _hdr(ws,2,["Metric","Value"],[34,22],F["header"])
        _fp = "'Full Pipeline'"; _n = fp_last_row_am
        kpi_f = [("Total Opportunities",f"=COUNTA({_fp}!C3:C{_n})",F["kpi_val"]),
                 ("Total Gross Pipeline (QAR)",f"=SUM({_fp}!I3:I{_n})",F["kpi_val"]),
                 ("Total Net Pipeline (QAR)",f"=SUM({_fp}!J3:J{_n})",F["kpi_val"]),
                 ("Forecasted Gross (QAR)",f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!I3:I{_n})',F["kpi_val"]),
                 ("Forecasted Net (QAR)",f'=SUMIF({_fp}!L3:L{_n},"Yes",{_fp}!J3:J{_n})',F["kpi_val"]),
                 ("Strategic Opportunities",f'=COUNTIF({_fp}!M3:M{_n},"Yes")',F["kpi_val"]),
                 ("Overdue Deals",f'=COUNTIF({_fp}!Q3:Q{_n},"YES")',F["kpi_val"])]
        for i,(lbl,frm,fmt) in enumerate(kpi_f):
            ws.write(3+i,0,lbl,F["kpi_lbl"]); ws.write_formula(3+i,1,frm,fmt)
        ws.merge_range("D1:H1","Pipeline by Stage",F["title"])
        _hdr(ws,2,["Stage","Count","Gross (QAR)","Net (QAR)"],None,F["header"])
        ws.set_column(3,3,22); ws.set_column(4,4,8); ws.set_column(5,5,18); ws.set_column(6,6,18)
        for i,row in stage_df.reset_index(drop=True).iterrows():
            xl1=4+i; alt=i%2==1
            ws.write(3+i,3,row["Stage"],F["alt"] if alt else F["text"])
            ws.write_formula(3+i,4,f'=COUNTIF({_fp}!D3:D{_n},D{xl1})',F["alt_num"] if alt else F["num"])
            ws.write_formula(3+i,5,f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!I3:I{_n})',F["alt_num"] if alt else F["num"])
            ws.write_formula(3+i,6,f'=SUMIF({_fp}!D3:D{_n},D{xl1},{_fp}!J3:J{_n})',F["alt_num"] if alt else F["num"])
        _sl=3+len(stage_df)
        ws.write(3+len(stage_df),3,"TOTAL",F["tot_lbl"])
        ws.write_formula(3+len(stage_df),4,f"=SUM(E4:E{_sl})",F["tot_num"])
        ws.write_formula(3+len(stage_df),5,f"=SUM(F4:F{_sl})",F["tot_num"])
        ws.write_formula(3+len(stage_df),6,f"=SUM(G4:G{_sl})",F["tot_num"])

        # SHEET 2 — BY ACCOUNT MANAGER
        am_ws = wb.add_worksheet("By Account Manager"); am_ws.set_zoom(90); am_ws.set_tab_color("#FF8C00")
        writer.sheets["By Account Manager"] = am_ws
        am_ws.merge_range("A1:F1",f"Pipeline by Capability Sales — {TODAY.strftime('%d %B %Y')}",F["title"]); am_ws.set_row(0,28)
        _hdr(am_ws,1,["Capability Sales","Deals","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[34,8,20,20,22],F["header"])
        for r,row in am_agg.reset_index(drop=True).iterrows():
            alt=r%2==1
            am_ws.write(2+r,0,row["Capability Sales"],F["alt"] if alt else F["text"])
            am_ws.write_number(2+r,1,row["Count"],F["alt_num"] if alt else F["num"])
            am_ws.write_number(2+r,2,row["Gross"],F["alt_num"] if alt else F["num"])
            am_ws.write_number(2+r,3,row["Net"],F["alt_num"] if alt else F["num"])
            am_ws.write_number(2+r,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])
        _am_rN=2+len(am_agg)
        am_ws.write(_am_rN,0,"TOTAL",F["tot_lbl"])
        am_ws.write_formula(_am_rN,1,f"=SUM(B3:B{_am_rN})",F["tot_num"]); am_ws.write_formula(_am_rN,2,f"=SUM(C3:C{_am_rN})",F["tot_num"])
        am_ws.write_formula(_am_rN,3,f"=SUM(D3:D{_am_rN})",F["tot_num"]); am_ws.write_formula(_am_rN,4,f"=SUM(E3:E{_am_rN})",F["tot_num"])
        # Deal detail by AM
        det_off=_am_rN+2
        am_ws.merge_range(det_off,0,det_off,6,"Deal Detail by Capability Sales",fmt_bu_hdr2); am_ws.set_row(det_off,22)
        _hdr(am_ws,det_off+1,["Capability Sales","Account Name","Opportunity","Stage","Quarter","Gross (QAR)","Net (QAR)"],[30,28,40,22,12,18,18],F["header"])
        _det_layout=[]
        _det_pos=det_off+2
        for am_name,am_grp in am_exp.groupby("AM_exp",sort=False):
            bu_xl=_det_pos; _det_pos+=1
            deal_rows_xl=[]
            for _ in am_grp.itertuples():
                deal_rows_xl.append(_det_pos); _det_pos+=1
            _det_layout.append((am_name,bu_xl,am_grp,deal_rows_xl))
        _det_bu_rows=[]
        for am_name,bu_xl,am_grp,deal_rows_xl in _det_layout:
            _dF="+".join("F"+str(r+1) for r in deal_rows_xl); _dG="+".join("G"+str(r+1) for r in deal_rows_xl)
            am_ws.write(bu_xl,0,am_name,F["bu_lbl"])
            for cc in range(1,7): am_ws.write(bu_xl,cc,"",F["bu_lbl"])
            am_ws.write_formula(bu_xl,5,"="+_dF,F["bu_num"]); am_ws.write_formula(bu_xl,6,"="+_dG,F["bu_num"])
            _det_bu_rows.append(bu_xl+1)
            for r_pos,(_, row) in zip(deal_rows_xl,am_grp.iterrows()):
                alt=r_pos%2==1
                am_ws.write(r_pos,0,am_name,F["alt"] if alt else F["text"])
                am_ws.write(r_pos,1,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",F["alt"] if alt else F["text"])
                am_ws.write(r_pos,2,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",F["alt"] if alt else F["text"])
                am_ws.write(r_pos,3,str(row["Stage_Short"]) if pd.notna(row["Stage_Short"]) else "",F["alt"] if alt else F["text"])
                am_ws.write(r_pos,4,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",F["alt"] if alt else F["text"])
                am_ws.write_number(r_pos,5,row["Total Gross"],F["alt_num"] if alt else F["num"])
                am_ws.write_number(r_pos,6,row["Total Net"],F["alt_num"] if alt else F["num"])
        _gt_F="+".join("F"+str(r) for r in _det_bu_rows); _gt_G="+".join("G"+str(r) for r in _det_bu_rows)
        _gt_row=_det_pos
        am_ws.write(_gt_row,0,"GRAND TOTAL",F["tot_lbl"])
        for cc in range(1,5): am_ws.write(_gt_row,cc,"",F["tot_lbl"])
        am_ws.write_formula(_gt_row,5,"="+_gt_F,F["tot_num"]); am_ws.write_formula(_gt_row,6,"="+_gt_G,F["tot_num"])

        # SHEET 3 — MONTHLY PIPELINE
        mw = wb.add_worksheet("Monthly Pipeline"); mw.set_zoom(90); mw.set_tab_color("#228B22")
        writer.sheets["Monthly Pipeline"] = mw
        mw.merge_range(0,0,0,14,f"Monthly Pipeline Breakdown — {TODAY.strftime('%d %B %Y')}",F["title"]); mw.set_row(0,28)
        mw.write(1,0,"Opportunity / Account",F["header"]); mw.set_column(0,0,40)
        mw.write(1,1,"Account Name",F["header"]); mw.set_column(1,1,28)
        for ci,m in enumerate(MONTHS_AM):
            mw.write(1,2+ci,m,fmt_month_hdr); mw.set_column(2+ci,2+ci,14)
        mw.write(1,14,"Total Net",fmt_month_hdr); mw.set_column(14,14,16)
        mw.freeze_panes(2,0)
        for r,row in df.reset_index(drop=True).iterrows():
            alt=r%2==1
            mw.write(2+r,0,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",F["alt"] if alt else F["text"])
            mw.write(2+r,1,str(row["Account Name"])  if pd.notna(row["Account Name"])  else "",F["alt"] if alt else F["text"])
            for ci,m in enumerate(MONTHS_AM):
                v=row[m]
                if v and v!=0: mw.write_number(2+r,2+ci,v,fmt_month if alt else F["num"])
                else: mw.write_blank(2+r,2+ci,None,F["alt"] if alt else F["text"])
            mw.write_number(2+r,14,row["Total Net"],F["alt_num"] if alt else F["num"])
        _m_rN=2+len(df)
        mw.write(_m_rN,0,"TOTAL",F["tot_lbl"]); mw.write(_m_rN,1,"",F["tot_lbl"])
        for ci,m in enumerate(MONTHS_AM):
            col_letter=chr(67+ci)
            mw.write_formula(_m_rN,2+ci,f"=SUM({col_letter}3:{col_letter}{_m_rN})",fmt_month_tot)
        mw.write_formula(_m_rN,14,f"=SUM(O3:O{_m_rN})",fmt_month_tot)
        mw.merge_range(_m_rN+2,0,_m_rN+2,14,"Monthly Totals Summary",fmt_bu_hdr2); mw.set_row(_m_rN+2,22)
        _hdr(mw,_m_rN+3,["Month"]+MONTHS_AM+["Total"],[14]+[14]*12+[16],F["header"])
        mw.write(_m_rN+4,0,"Net (QAR)",F["bu_lbl"])
        for ci,m in enumerate(MONTHS_AM):
            col_letter=chr(67+ci)
            mw.write_formula(_m_rN+4,1+ci,f"={col_letter}{_m_rN+1}",F["bu_num"])
        mw.write_formula(_m_rN+4,13,f"=O{_m_rN+1}",F["bu_num"])

        # SHEET 4 — DU BREAKDOWN
        du_ws = wb.add_worksheet("DU Breakdown"); du_ws.set_zoom(90); du_ws.set_tab_color("#FF8C00")
        writer.sheets["DU Breakdown"] = du_ws
        du_ws.merge_range("A1:E1","Gross & Net Breakdown by BU / Delivery Unit",F["title"]); du_ws.set_row(0,28)
        _hdr(du_ws,1,["BU","Delivery Unit / Opportunity","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[42,52,20,20,22],F["header"])
        _adu_layout2=[]; _apos2=0
        for bu_name,bu_grp in du_totals_am.groupby("BU"):
            bu_r0=2+_apos2; _apos2+=1
            du_list2=[]
            for _,drow in bu_grp.iterrows():
                du_r0=2+_apos2; _apos2+=1
                deals2=du_exp_am[du_exp_am["DU"]==drow["DU"]]
                opp_rows2=[]
                for _,deal in deals2.iterrows():
                    opp_rows2.append((2+_apos2,deal)); _apos2+=1
                du_list2.append((drow,du_r0,opp_rows2))
            _adu_layout2.append((bu_name,bu_r0,du_list2))
        _agt2=2+_apos2
        for bu_name,bu_r0,du_list2 in _adu_layout2:
            du_ws.write(bu_r0,0,bu_name,F["bu_lbl"]); du_ws.write(bu_r0,1,"",F["bu_lbl"])
            _dur2=[dr0+1 for (_,dr0,_) in du_list2]
            _bC2="+".join("C"+str(r) for r in _dur2); _bD2="+".join("D"+str(r) for r in _dur2); _bE2="+".join("E"+str(r) for r in _dur2)
            du_ws.write_formula(bu_r0,2,"="+_bC2,F["bu_num"]); du_ws.write_formula(bu_r0,3,"="+_bD2,F["bu_num"]); du_ws.write_formula(bu_r0,4,"="+_bE2,F["bu_num"])
            for drow,du_r0,opp_rows2 in du_list2:
                alt=(du_r0-2)%2==1
                du_ws.write(du_r0,0,"",F["alt"] if alt else F["text"]); du_ws.write(du_r0,1,drow["DU"],F["alt"] if alt else F["text"])
                if opp_rows2:
                    _r1=min(r for r,_ in opp_rows2)+1; _rN=max(r for r,_ in opp_rows2)+1
                    du_ws.write_formula(du_r0,2,f"=SUM(C{_r1}:C{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,3,f"=SUM(D{_r1}:D{_rN})",F["alt_num"] if alt else F["num"])
                    du_ws.write_formula(du_r0,4,f"=SUM(E{_r1}:E{_rN})",F["alt_num"] if alt else F["num"])
                else:
                    for _c in range(2,5): du_ws.write_number(du_r0,_c,0,F["alt_num"] if alt else F["num"])
                for opp_r0,deal in opp_rows2:
                    fore_net=deal["Net"] if str(deal.get("Forecasted","")).strip()=="Yes" else 0
                    du_ws.write(opp_r0,0,"",F["opp"]); du_ws.write(opp_r0,1,f"  ↳  {deal['Lead/Opp Name']}",F["opp"])
                    du_ws.write_number(opp_r0,2,deal["Gross"],F["opp_num"]); du_ws.write_number(opp_r0,3,deal["Net"],F["opp_num"]); du_ws.write_number(opp_r0,4,fore_net,F["opp_num"])
        du_ws.write(_agt2,0,"GRAND TOTAL",F["tot_lbl"]); du_ws.write(_agt2,1,"",F["tot_lbl"])
        _gbC2="+".join("C"+str(br0+1) for (_,br0,_) in _adu_layout2)
        _gbD2="+".join("D"+str(br0+1) for (_,br0,_) in _adu_layout2)
        _gbE2="+".join("E"+str(br0+1) for (_,br0,_) in _adu_layout2)
        du_ws.write_formula(_agt2,2,"="+_gbC2,F["tot_num"]); du_ws.write_formula(_agt2,3,"="+_gbD2,F["tot_num"]); du_ws.write_formula(_agt2,4,"="+_gbE2,F["tot_num"])

        # SHEET 5 — QUARTERLY PLAN
        qw = wb.add_worksheet("Quarterly Plan"); qw.set_zoom(90); qw.set_tab_color("#DAA520")
        writer.sheets["Quarterly Plan"] = qw
        qw.merge_range("A1:E1",f"Quarterly Close Plan — {TODAY.strftime('%d %B %Y')}",F["title"]); qw.set_row(0,28)
        _hdr(qw,1,["Quarter","Count","Gross (QAR)","Net (QAR)","Forecasted Net (QAR)"],[14,8,20,20,22],F["header"])
        for r,row in q_df.reset_index(drop=True).iterrows():
            alt=r%2==1
            qw.write(2+r,0,row["Closure Due Quarter"],F["alt"] if alt else F["text"])
            qw.write_number(2+r,1,row["Count"],F["alt_num"] if alt else F["num"])
            qw.write_number(2+r,2,row["Gross"],F["alt_num"] if alt else F["num"])
            qw.write_number(2+r,3,row["Net"],F["alt_num"] if alt else F["num"])
            qw.write_number(2+r,4,row["Forecasted Net"],F["alt_num"] if alt else F["num"])
        _q2_rN=2+len(q_df)
        qw.write(_q2_rN,0,"TOTAL",F["tot_lbl"])
        qw.write_formula(_q2_rN,1,f"=SUM(B3:B{_q2_rN})",F["tot_num"]); qw.write_formula(_q2_rN,2,f"=SUM(C3:C{_q2_rN})",F["tot_num"])
        qw.write_formula(_q2_rN,3,f"=SUM(D3:D{_q2_rN})",F["tot_num"]); qw.write_formula(_q2_rN,4,f"=SUM(E3:E{_q2_rN})",F["tot_num"])

        # SHEET 6 — AM BREAKDOWN (styled)
        _AMCLR={"title_bg":"1F3864","hdr_deal":"1F3864","hdr_am":"17375E","hdr_finance":"1F4E79","hdr_other":"2E5FA3",
                "am_fill":"EDF2F9","num_fill":"EBF5FB","tot_fill":"D5E8F5","date_fill":"FFF2CC","alt_a":"F5F8FF","alt_b":"FFFFFF"}
        def _amf(bg,top,num_fmt=None):
            d={"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
            if num_fmt: d["num_format"]=num_fmt; d["align"]="right"
            return wb.add_format(d)
        am_fh_deal=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_AMCLR["hdr_deal"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        am_fh_am=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_AMCLR["hdr_am"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        am_fh_finance=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_AMCLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        am_fh_other=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_AMCLR["hdr_other"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        am_fmt_title=wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+_AMCLR["title_bg"],"align":"center","valign":"vcenter"})
        am_ft_a1=_amf(_AMCLR["alt_a"],2); am_ft_an=_amf(_AMCLR["alt_a"],1)
        am_ft_b1=_amf(_AMCLR["alt_b"],2); am_ft_bn=_amf(_AMCLR["alt_b"],1)
        am_fn_a1=_amf(_AMCLR["alt_a"],2,"#,##0"); am_fn_an=_amf(_AMCLR["alt_a"],1,"#,##0")
        am_fn_b1=_amf(_AMCLR["alt_b"],2,"#,##0"); am_fn_bn=_amf(_AMCLR["alt_b"],1,"#,##0")
        am_fam_f=_amf(_AMCLR["am_fill"],2); am_fam_n=_amf(_AMCLR["am_fill"],1)
        am_fxn_f=_amf(_AMCLR["num_fill"],2,"#,##0"); am_fxn_n=_amf(_AMCLR["num_fill"],1,"#,##0")
        am_bd_ftot=wb.add_format({"bg_color":"#"+_AMCLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
        am_bd_ftotbl=wb.add_format({"bg_color":"#"+_AMCLR["num_fill"],"border":1,"font_size":9})
        am_bd_fdate=wb.add_format({"bg_color":"#"+_AMCLR["date_fill"],"num_format":"dd-mmm-yyyy","border":1,"align":"center","font_size":9})
        am_bd_fdatebl=wb.add_format({"bg_color":"#"+_AMCLR["alt_a"],"border":1,"font_size":9})
        am_bw=wb.add_worksheet("AM Breakdown"); am_bw.set_zoom(85); am_bw.set_tab_color("#9370DB")
        writer.sheets["AM Breakdown"]=am_bw
        am_bd_ocols=[("SNo.",6,"deal"),("Account Name",24,"deal"),("Lead/Opp Name",36,"deal"),
                     ("Capability Sales",30,"am"),("Stage",22,"other"),("All Capability Sales",24,"other"),
                     ("Sector",16,"other"),("Quarter",10,"other"),("Win Prob",12,"other"),
                     ("Forecasted",10,"other"),("Strategic",10,"other"),
                     ("Gross (breakdown)",16,"finance"),("Net (breakdown)",16,"finance"),
                     ("Total Gross",15,"finance"),("Total Net",15,"finance"),("Est. Close Date",14,"other")]
        am_hfmt={"deal":am_fh_deal,"am":am_fh_am,"finance":am_fh_finance,"other":am_fh_other}
        am_nc=len(am_bd_ocols)
        am_bw.merge_range(0,0,0,am_nc-1,"AM Pipeline — Expanded by Capability Sales",am_fmt_title); am_bw.set_row(0,28)
        for c,(cn,cw,ct) in enumerate(am_bd_ocols): am_bw.write(1,c,cn,am_hfmt[ct]); am_bw.set_column(c,c,cw)
        am_bw.set_row(1,28); am_bw.freeze_panes(2,0)
        am_bd_deal_rows={}
        for rp,(_, row) in enumerate(am_exp.iterrows()):
            didx=row["_deal_idx"]; xl_r=2+rp
            if didx not in am_bd_deal_rows: am_bd_deal_rows[didx]=[xl_r,xl_r]
            else: am_bd_deal_rows[didx][1]=xl_r
        am_bw.autofilter(1,0,1+len(am_exp),am_nc-1)
        am_prev=None; am_alt=False
        for rp,(_, row) in enumerate(am_exp.iterrows()):
            didx=row["_deal_idx"]; isf=bool(row["_is_first"])
            if didx!=am_prev: am_alt=not am_alt; am_prev=didx
            ft=(am_ft_a1 if isf else am_ft_an) if am_alt else (am_ft_b1 if isf else am_ft_bn)
            fn=(am_fn_a1 if isf else am_fn_an) if am_alt else (am_fn_b1 if isf else am_fn_bn)
            fam=am_fam_f if isf else am_fam_n; fxn=am_fxn_f if isf else am_fxn_n; xl_r=2+rp
            def _amws(ci,val,fmt):
                if val is None or (isinstance(val,float) and pd.isna(val)): am_bw.write_blank(xl_r,ci,None,fmt)
                else: am_bw.write(xl_r,ci,str(val),fmt)
            def _amwn(ci,val,fmt):
                v=_parse_num(val) if not isinstance(val,(int,float)) else val
                if v is None or (isinstance(v,float) and pd.isna(v)): am_bw.write_blank(xl_r,ci,None,fmt)
                else: am_bw.write_number(xl_r,ci,v,fmt)
            _amws(0,row.get("SNo.") if isf else None,ft)
            _amws(1,row.get("Account Name"),ft); _amws(2,row.get("Lead/Opp Name"),ft)
            _amws(3,row.get("AM_exp"),fam); _amws(4,row.get("Stage_Short"),ft)
            _amws(5,row.get("Capability Sales"),ft); _amws(6,row.get("Sector"),ft)
            _amws(7,row.get("Closure Due Quarter"),ft); _amws(8,row.get("Winning Probability"),ft)
            _amws(9,row.get("Forecasted"),ft); _amws(10,row.get("Strategic Opportunity"),ft)
            g_col_idx=11; n_col_idx=12
            _amwn(g_col_idx,_parse_num(str(row.get("Gross (Breakdown)","")).split("\n")[0]) if isf else None,fxn)
            _amwn(n_col_idx,_parse_num(str(row.get("Net (breakdown)","")).split("\n")[0])   if isf else None,fxn)
            if isf:
                r0,r1=am_bd_deal_rows[didx]; gc=chr(65+g_col_idx); nc=chr(65+n_col_idx)
                am_bw.write_formula(xl_r,13,f"=SUM({gc}{r0+1}:{gc}{r1+1})",am_bd_ftot)
                am_bw.write_formula(xl_r,14,f"=SUM({nc}{r0+1}:{nc}{r1+1})",am_bd_ftot)
            else:
                am_bw.write_blank(xl_r,13,None,am_bd_ftotbl); am_bw.write_blank(xl_r,14,None,am_bd_ftotbl)
            cd_v=row.get("Est. Close Date")
            if isf and pd.notna(cd_v):
                try: am_bw.write_datetime(xl_r,15,pd.Timestamp(cd_v).to_pydatetime(),am_bd_fdate)
                except: am_bw.write(xl_r,15,str(cd_v),am_bd_fdate)
            else: am_bw.write_blank(xl_r,15,None,am_bd_fdate if isf else am_bd_fdatebl)

        # SHEET 7 — FULL PIPELINE
        pw = wb.add_worksheet("Full Pipeline"); pw.set_zoom(85); pw.set_tab_color("#6495ED")
        writer.sheets["Full Pipeline"]=pw
        pw.merge_range("A1:Q1","Full AM Pipeline — All Opportunities",F["title"]); pw.set_row(0,28); pw.freeze_panes(2,0)
        _hdr(pw,1,["#","Account Name","Opportunity","Stage","Capability Sales","Sector","BU","DU","Gross (QAR)","Net (QAR)","Win Prob","Forecasted","Strategic","Quarter","Est. Close Date","Source","Overdue"],[5,28,36,22,28,16,30,36,18,18,12,12,10,10,16,16,8],F["header"])
        pw.autofilter(1,0,1+len(full_df_am),16)
        for r,row in full_df_am.reset_index(drop=True).iterrows():
            alt=r%2==1; is_ov=bool(row["Overdue"])
            ft=F["red"] if is_ov else (F["alt"] if alt else F["text"]); fn=F["red_num"] if is_ov else (F["alt_num"] if alt else F["num"]); fd_=F["red_dt"] if is_ov else F["date"]
            pw.write(2+r,0,str(row["SNo."]) if pd.notna(row["SNo."]) else "",ft)
            pw.write(2+r,1,str(row["Account Name"]) if pd.notna(row["Account Name"]) else "",ft)
            pw.write(2+r,2,str(row["Lead/Opp Name"]) if pd.notna(row["Lead/Opp Name"]) else "",ft)
            pw.write(2+r,3,str(row["Stage"]) if pd.notna(row["Stage"]) else "",ft)
            pw.write(2+r,4,str(row["Capability Sales"]) if pd.notna(row["Capability Sales"]) else "",ft)
            pw.write(2+r,5,str(row["Sector"]) if pd.notna(row["Sector"]) else "",ft)
            pw.write(2+r,6,str(row["BU"]) if pd.notna(row["BU"]) else "",ft)
            pw.write(2+r,7,str(row["DU"]) if pd.notna(row["DU"]) else "",ft)
            pw.write_number(2+r,8,row["Total Gross"],fn); pw.write_number(2+r,9,row["Total Net"],fn)
            pw.write(2+r,10,str(row["Winning Probability"]) if pd.notna(row["Winning Probability"]) else "",ft)
            pw.write(2+r,11,str(row["Forecasted"]) if pd.notna(row["Forecasted"]) else "",ft)
            pw.write(2+r,12,str(row["Strategic Opportunity"]) if pd.notna(row["Strategic Opportunity"]) else "",ft)
            pw.write(2+r,13,str(row["Closure Due Quarter"]) if pd.notna(row["Closure Due Quarter"]) else "",ft)
            if pd.notna(row["Est. Close Date"]):
                pw.write_datetime(2+r,14,row["Est. Close Date"].to_pydatetime(),fd_)
            else: pw.write_blank(2+r,14,None,fd_)
            pw.write(2+r,15,str(row["Source of Opportunity"]) if pd.notna(row["Source of Opportunity"]) else "",ft)
            pw.write(2+r,16,"YES" if is_ov else "",ft)
        pw.write(fp_last_row_am,0,"TOTAL",F["tot_lbl"])
        pw.write_formula(fp_last_row_am,8,f"=SUM(I3:I{fp_last_row_am})",F["tot_num"])
        pw.write_formula(fp_last_row_am,9,f"=SUM(J3:J{fp_last_row_am})",F["tot_num"])

        # SHEET 8 — PIPELINE BREAKDOWN (styled, one row per DU per opportunity)
        _PBCLR = {"title_bg":"1F3864","hdr_deal":"1F3864","hdr_du":"17375E",
                  "hdr_finance":"1F4E79","hdr_other":"2E5FA3",
                  "bu_fill":"EDF2F9","du_fill":"E4ECF7","num_fill":"EBF5FB",
                  "tot_fill":"D5E8F5","date_fill":"FFF2CC","alt_a":"F5F8FF","alt_b":"FFFFFF"}
        def _pbf(bg,top,num_fmt=None):
            d={"bg_color":"#"+bg,"top":top,"bottom":1,"left":1,"right":1,"font_size":9}
            if num_fmt: d["num_format"]=num_fmt; d["align"]="right"
            return wb.add_format(d)
        pb_fh_deal    = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_PBCLR["hdr_deal"],   "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        pb_fh_du      = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_PBCLR["hdr_du"],     "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        pb_fh_finance = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_PBCLR["hdr_finance"],"border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        pb_fh_other   = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#"+_PBCLR["hdr_other"],  "border":1,"align":"center","valign":"vcenter","text_wrap":True,"font_size":9})
        pb_fmt_title  = wb.add_format({"bold":True,"font_size":13,"font_color":"#FFFFFF","bg_color":"#"+_PBCLR["title_bg"],"align":"center","valign":"vcenter"})
        pb_ft_a1=_pbf(_PBCLR["alt_a"],2); pb_ft_an=_pbf(_PBCLR["alt_a"],1)
        pb_ft_b1=_pbf(_PBCLR["alt_b"],2); pb_ft_bn=_pbf(_PBCLR["alt_b"],1)
        pb_fn_a1=_pbf(_PBCLR["alt_a"],2,"#,##0"); pb_fn_an=_pbf(_PBCLR["alt_a"],1,"#,##0")
        pb_fn_b1=_pbf(_PBCLR["alt_b"],2,"#,##0"); pb_fn_bn=_pbf(_PBCLR["alt_b"],1,"#,##0")
        pb_fbu1=_pbf(_PBCLR["bu_fill"],2); pb_fbun=_pbf(_PBCLR["bu_fill"],1)
        pb_fdu1=_pbf(_PBCLR["du_fill"],2); pb_fdun=_pbf(_PBCLR["du_fill"],1)
        pb_fxn1=_pbf(_PBCLR["num_fill"],2,"#,##0"); pb_fxnn=_pbf(_PBCLR["num_fill"],1,"#,##0")
        pb_ftot   = wb.add_format({"bg_color":"#"+_PBCLR["tot_fill"],"num_format":"#,##0","border":1,"align":"right","bold":True,"font_size":9})
        pb_ftotbl = wb.add_format({"bg_color":"#"+_PBCLR["num_fill"],"border":1,"font_size":9})
        pb_fdate  = wb.add_format({"bg_color":"#"+_PBCLR["date_fill"],"border":1,"align":"center","num_format":"DD-MMM-YYYY","font_size":9})
        pb_fdatebl= wb.add_format({"bg_color":"#"+_PBCLR["alt_a"],"border":1,"font_size":9})

        pbw = wb.add_worksheet("Pipeline Breakdown"); pbw.set_zoom(85); pbw.set_tab_color("#9370DB")
        writer.sheets["Pipeline Breakdown"] = pbw
        pb_ocols = [
            ("SNo.",         6, "deal"),("Account Name",   24,"deal"),("Lead/Opp Name",   36,"deal"),
            ("BU",          36, "du"),  ("DU",             34,"du"),
            ("Gross (breakdown)",16,"finance"),("Net (breakdown)",16,"finance"),
            ("Total Gross", 15, "finance"),("Total Net",   15,"finance"),
            ("Stage",       36, "other"),("Capability Sales",28,"other"),("Sector",16,"other"),
            ("Quarter",      9, "other"),("Win Prob",      10,"other"),
            ("Forecasted",  10, "other"),("Strategic",     10,"other"),("Est. Close Date",14,"other"),
        ]
        pb_hfmt = {"deal":pb_fh_deal,"du":pb_fh_du,"finance":pb_fh_finance,"other":pb_fh_other}
        pb_nc = len(pb_ocols)
        pbw.merge_range(0,0,0,pb_nc-1,"AM Pipeline — Expanded by Delivery Unit",pb_fmt_title); pbw.set_row(0,28)
        for c,(cn,cw,ct) in enumerate(pb_ocols): pbw.write(1,c,cn,pb_hfmt[ct]); pbw.set_column(c,c,cw)
        pbw.set_row(1,28); pbw.freeze_panes(2,0)

        # Expand deals by DU using mapping
        pb_exp = _expand_deals(df, mapping=mapping)

        pbw.autofilter(1,0,1+len(pb_exp),pb_nc-1)
        pb_cmap = {cn:idx for idx,(cn,_,__) in enumerate(pb_ocols)}
        pb_gcol = pb_cmap["Gross (breakdown)"]; pb_ncol = pb_cmap["Net (breakdown)"]

        # Pre-compute deal row ranges
        pb_deal_rows = {}
        for rp,(_, row) in enumerate(pb_exp.iterrows()):
            _di=row["_deal_idx"]; _xr=2+rp
            if _di not in pb_deal_rows: pb_deal_rows[_di]=[_xr,_xr]
            else: pb_deal_rows[_di][1]=_xr

        pb_prev=None; pb_alt=False
        for rp,(_, row) in enumerate(pb_exp.iterrows()):
            didx=row["_deal_idx"]; isf=bool(row["_is_first"])
            if didx!=pb_prev: pb_alt=not pb_alt; pb_prev=didx
            ft=(pb_ft_a1 if isf else pb_ft_an) if pb_alt else (pb_ft_b1 if isf else pb_ft_bn)
            fn=(pb_fn_a1 if isf else pb_fn_an) if pb_alt else (pb_fn_b1 if isf else pb_fn_bn)
            fbu=pb_fbu1 if isf else pb_fbun; fdu=pb_fdu1 if isf else pb_fdun
            fxn=pb_fxn1 if isf else pb_fxnn; xl_r=2+rp
            def _pbs(ci,val,fmt):
                if val is None or (isinstance(val,float) and pd.isna(val)): pbw.write_blank(xl_r,ci,None,fmt)
                else: pbw.write(xl_r,ci,str(val),fmt)
            def _pbn(ci,val,fmt):
                v=_parse_num(val) if not isinstance(val,(int,float)) else val
                if v is None or (isinstance(v,float) and pd.isna(v)): pbw.write_blank(xl_r,ci,None,fmt)
                else: pbw.write_number(xl_r,ci,v,fmt)
            _pbs(pb_cmap["SNo."],           row.get("SNo.") if isf else None, ft)
            _pbs(pb_cmap["Account Name"],    row.get("Account Name"), ft)
            _pbs(pb_cmap["Lead/Opp Name"],   row.get("Lead/Opp Name"), ft)
            _pbs(pb_cmap["BU"],              row.get("BU_exp"), fbu)
            _pbs(pb_cmap["DU"],              row.get("DU_exp"), fdu)
            _pbn(pb_cmap["Gross (breakdown)"], row.get("Gross_exp"), fxn)
            _pbn(pb_cmap["Net (breakdown)"],   row.get("Net_exp"),   fxn)
            if isf:
                r0,r1=pb_deal_rows[didx]; gc=chr(65+pb_gcol); nc=chr(65+pb_ncol)
                pbw.write_formula(xl_r,pb_cmap["Total Gross"],f"=SUM({gc}{r0+1}:{gc}{r1+1})",pb_ftot)
                pbw.write_formula(xl_r,pb_cmap["Total Net"],  f"=SUM({nc}{r0+1}:{nc}{r1+1})",pb_ftot)
            else:
                pbw.write_blank(xl_r,pb_cmap["Total Gross"],None,pb_ftotbl)
                pbw.write_blank(xl_r,pb_cmap["Total Net"],  None,pb_ftotbl)
            _pbs(pb_cmap["Stage"],           row.get("Stage_Short"), ft)
            _pbs(pb_cmap["Capability Sales"],row.get("Capability Sales"), ft)
            _pbs(pb_cmap["Sector"],          row.get("Sector"), ft)
            _pbs(pb_cmap["Quarter"],         row.get("Closure Due Quarter"), ft)
            _pbs(pb_cmap["Win Prob"],        row.get("Winning Probability"), ft)
            _pbs(pb_cmap["Forecasted"],      row.get("Forecasted"), ft)
            _pbs(pb_cmap["Strategic"],       row.get("Strategic Opportunity"), ft)
            cd_v=row.get("Est. Close Date")
            if isf and pd.notna(cd_v):
                try: pbw.write_datetime(xl_r,pb_cmap["Est. Close Date"],pd.Timestamp(cd_v).to_pydatetime(),pb_fdate)
                except: pbw.write(xl_r,pb_cmap["Est. Close Date"],str(cd_v),pb_fdate)
            else: pbw.write_blank(xl_r,pb_cmap["Est. Close Date"],None,pb_fdate if isf else pb_fdatebl)

    output.seek(0)
    return output.read()


# ── SIDEBAR ────────────────────────────────────────────────────────────────────
st.sidebar.title("📁 Load Data")
uploaded      = st.sidebar.file_uploader("Pipeline Excel (weekly report)", type=["xlsx","xls"])
uploaded_aw   = st.sidebar.file_uploader("Awarded Deals 2026",             type=["xlsx","xls"])
uploaded_aw25 = st.sidebar.file_uploader("Awarded Deals 2025",             type=["xlsx","xls"])
uploaded_b3   = st.sidebar.file_uploader("Book3 (Resource Forecast)",      type=["xlsx","xls"])
uploaded_am   = st.sidebar.file_uploader("AM Pipeline (Capability Sales)", type=["xlsx","xls"])

have_awarded = uploaded_aw or uploaded_aw25

if not uploaded and not have_awarded and not uploaded_am:
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

if uploaded_am:
    st.sidebar.success("AM Pipeline file loaded ✓")

# ── TABS ──────────────────────────────────────────────────────────────────────
st.title("📊 Sales Weekly Review Dashboard")
st.caption(f"Report date: {date.today().strftime('%d %B %Y')}")

tab_labels = []
if uploaded:     tab_labels.append("🔵 Pipeline")
if have_awarded: tab_labels.append("🟢 Awarded Deals")
if uploaded_am:  tab_labels.append("🟠 AM Pipeline")
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
# AM PIPELINE TAB
# ══════════════════════════════════════════════════════════════════════════════
if uploaded_am and "🟠 AM Pipeline" in tab_idx:
  with tabs[tab_idx["🟠 AM Pipeline"]]:
    df_am = load_am_pipeline(uploaded_am)
    cap_col_am, btn_col_am = st.columns([6, 2])
    cap_col_am.caption(f"{len(df_am)} opportunities — AM Pipeline (Capability Sales view)")
    with btn_col_am:
        am_xl_bytes = export_am_pipeline_excel(uploaded_am)
        st.download_button(
            label="⬇️ Export AM Pipeline Report",
            data=am_xl_bytes,
            file_name=f"AM_Pipeline_Report_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # KPIs
    k1,k2,k3,k4,k5 = st.columns(5)
    k1.metric("Total Opportunities", len(df_am))
    k2.metric("Total Gross Pipeline", f"{df_am['Total Gross'].sum()/1e6:.1f}M")
    k3.metric("Total Net Pipeline",   f"{df_am['Total Net'].sum()/1e6:.1f}M")
    k4.metric("Forecasted (Net)",     f"{df_am[df_am['Forecasted']=='Yes']['Total Net'].sum()/1e6:.1f}M")
    k5.metric("Strategic Opps",       int((df_am["Strategic Opportunity"]=="Yes").sum()))
    st.markdown("---")

    # AM Summary table
    _am_exp_ui = []
    for deal_idx, row in df_am.iterrows():
        ams = _clean_am_list(row.get("Capability Sales",""))
        if not ams: ams = ["Unassigned"]
        for am in ams:
            _am_exp_ui.append({"Account Manager": am, "Total Gross": row["Total Gross"], "Total Net": row["Total Net"],
                               "Forecasted": row.get("Forecasted",""), "Lead/Opp Name": row.get("Lead/Opp Name","")})
    am_ui_df = pd.DataFrame(_am_exp_ui)
    am_ui_agg = am_ui_df.groupby("Account Manager").agg(Opps=("Lead/Opp Name","nunique"),Gross=("Total Gross","sum"),Net=("Total Net","sum")).reset_index().sort_values("Net",ascending=False)

    col_am1, col_am2 = st.columns(2)
    with col_am1:
        st.subheader("Net Pipeline by Account Manager")
        fig_am2 = px.bar(am_ui_agg, x="Account Manager", y="Net", text=am_ui_agg["Net"].apply(lambda v: f"{v/1e6:.1f}M"),
                         color="Net", color_continuous_scale="Blues")
        fig_am2.update_traces(textposition="outside")
        fig_am2.update_layout(height=350, showlegend=False, coloraxis_showscale=False,
                              margin=dict(l=10,r=10,t=10,b=80), xaxis_tickangle=-30)
        st.plotly_chart(fig_am2, use_container_width=True)
    with col_am2:
        st.subheader("AM Summary")
        st.dataframe(am_ui_agg.assign(**{"Gross (M)":am_ui_agg["Gross"]/1e6,"Net (M)":am_ui_agg["Net"]/1e6})
                     [["Account Manager","Opps","Gross (M)","Net (M)"]]
                     .style.format({"Gross (M)":"{:.1f}","Net (M)":"{:.1f}"}),
                     use_container_width=True, hide_index=True)

    # Monthly pipeline chart
    st.markdown("---")
    st.subheader("Monthly Pipeline Breakdown")
    monthly_vals = {m: float(df_am[m].sum()) for m in MONTHS_AM}
    monthly_df = pd.DataFrame({"Month": list(monthly_vals.keys()), "Net (QAR)": list(monthly_vals.values())})
    fig_month = px.bar(monthly_df, x="Month", y="Net (QAR)", text=monthly_df["Net (QAR)"].apply(lambda v: f"{v/1e6:.1f}M"),
                       color="Net (QAR)", color_continuous_scale="Greens")
    fig_month.update_traces(textposition="outside")
    fig_month.update_layout(height=350, showlegend=False, coloraxis_showscale=False, margin=dict(l=10,r=10,t=10,b=40))
    st.plotly_chart(fig_month, use_container_width=True)

    # Full data table
    st.markdown("---")
    st.subheader("Full AM Pipeline")
    st.dataframe(df_am[["SNo.","Account Name","Lead/Opp Name","Stage_Short","Capability Sales","Sector","Total Gross","Total Net","Winning Probability","Forecasted","Closure Due Quarter"]].rename(columns={"Stage_Short":"Stage"}),
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
