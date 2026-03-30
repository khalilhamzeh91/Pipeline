"""
Book3 Mapping Tool
Upload Pipeline, Awarded Deals, and Book3 (Resource Forecast) to get a
cross-reference mapping showing which Book3 projects match pipeline/awarded deals.
Run: streamlit run book3_mapping.py
"""

import re
import io
import streamlit as st
import pandas as pd
from datetime import date
from difflib import SequenceMatcher
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="Book3 Mapping", page_icon="🔗", layout="wide")

MONTHS     = ["Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
SKIP_KEYS  = ["Total","Grand Total","Existing Renewal Total",
              "Opportunity (ORF) Total","Opportunity Pipeline Total"]

# ── LOADERS ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_pipeline(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    df = df.dropna(subset=["Stage"])
    df["Total Gross"] = pd.to_numeric(df["Total Gross"], errors="coerce").fillna(0)
    df["Total Net"]   = pd.to_numeric(df["Total Net"],   errors="coerce").fillna(0)
    return df

@st.cache_data
def load_awarded(file, year):
    df = pd.read_excel(file, sheet_name="Export")
    df.columns = df.columns.str.strip()
    df["Total Gross"] = pd.to_numeric(df["Total Gross"], errors="coerce").fillna(0)
    df["Total Net"]   = pd.to_numeric(df["Total Net"],   errors="coerce").fillna(0)
    df["Year"] = year
    return df

@st.cache_data
def load_book3(file):
    raw = pd.read_excel(file, header=None)
    cols = ["_","BU","Project Type","Project Name"] + MONTHS + ["Grand Total"]
    raw.columns = cols[:len(raw.columns)]
    raw = raw.iloc[2:].reset_index(drop=True)
    current_bu, current_type = None, None
    rows = []
    for _, r in raw.iterrows():
        bu    = str(r["BU"]).strip()           if pd.notna(r["BU"])           else ""
        ptype = str(r["Project Type"]).strip() if pd.notna(r["Project Type"]) else ""
        name  = str(r["Project Name"]).strip() if pd.notna(r["Project Name"]) else ""
        if bu and bu != "nan" and not any(k in bu for k in SKIP_KEYS):
            current_bu = bu
        if ptype and ptype != "nan" and not any(k in ptype for k in SKIP_KEYS):
            current_type = ptype
        if not name or name == "nan" or any(k in name for k in SKIP_KEYS):
            continue
        if any(k in (ptype or "") for k in SKIP_KEYS):
            continue
        row_data = {"BU": current_bu, "Project Type": current_type, "Project Name": name}
        for m in MONTHS:
            val = r.get(m, None)
            try:    row_data[m] = float(val) if pd.notna(val) else 0.0
            except: row_data[m] = 0.0
        try:    row_data["Grand Total"] = float(r.get("Grand Total", 0)) if pd.notna(r.get("Grand Total")) else 0.0
        except: row_data["Grand Total"] = 0.0
        rows.append(row_data)
    return pd.DataFrame(rows)

# ── MATCHING ─────────────────────────────────────────────────────────────────
def _clean(s):
    return set(re.sub(r"[^a-z0-9]", " ", str(s).lower()).split())

def best_match(name, candidates, threshold=0.55):
    tokens = _clean(name)
    best, best_score = None, 0.0
    for cand in candidates:
        overlap = len(tokens & _clean(cand)) / max(len(tokens | _clean(cand)), 1)
        seq     = SequenceMatcher(None, name.lower(), cand.lower()).ratio()
        score   = max(overlap, seq)
        if score > best_score:
            best, best_score = cand, score
    return (best, round(best_score, 2)) if best_score >= threshold else (None, 0.0)

def build_mapping(book3_df, pipeline_df, awarded_df):
    pipe_names  = pipeline_df["Lead/Opp Name"].dropna().tolist()  if pipeline_df  is not None else []
    award_names = awarded_df["Opportunity Name"].dropna().tolist() if awarded_df   is not None else []
    rows = []
    for _, b3 in book3_df.iterrows():
        pipe_match,  pipe_score  = best_match(b3["Project Name"], pipe_names)
        award_match, award_score = best_match(b3["Project Name"], award_names)
        p = pipeline_df[pipeline_df["Lead/Opp Name"]==pipe_match].iloc[0]    if pipe_match  and pipeline_df  is not None else None
        a = awarded_df[awarded_df["Opportunity Name"]==award_match].iloc[0]  if award_match and awarded_df   is not None else None
        row = {
            "Book3 BU":           b3["BU"],
            "Book3 Project Type": b3["Project Type"],
            "Book3 Project Name": b3["Project Name"],
            "Book3 Grand Total":  b3["Grand Total"],
        }
        for m in MONTHS: row[f"Book3 {m}"] = b3[m]
        row.update({
            "Pipeline Match":       pipe_match  or "",
            "Pipeline Score":       pipe_score,
            "Pipeline Account":     str(p["Account Name"])    if p is not None else "",
            "Pipeline Gross (QAR)": float(p["Total Gross"])   if p is not None else 0.0,
            "Pipeline Net (QAR)":   float(p["Total Net"])     if p is not None else 0.0,
            "Pipeline Stage":       str(p["Stage"])           if p is not None else "",
            "Pipeline AM":          str(p["Account Manager"]) if p is not None else "",
            "Pipeline Quarter":     str(p["Closure Due Quarter"]) if p is not None else "",
            "Awarded Match":        award_match or "",
            "Awarded Score":        award_score,
            "Awarded Account":      str(a["Account Name"])    if a is not None else "",
            "Awarded Gross (QAR)":  float(a["Total Gross"])   if a is not None else 0.0,
            "Awarded Net (QAR)":    float(a["Total Net"])     if a is not None else 0.0,
            "Awarded Stage":        str(a["Stage"])           if a is not None else "",
            "Awarded AM":           str(a["Account Manager"]) if a is not None else "",
            "Awarded Year":         str(a["Year"])            if a is not None else "",
        })
        rows.append(row)
    return pd.DataFrame(rows)

# ── EXCEL EXPORT ─────────────────────────────────────────────────────────────
def export_mapping_excel(map_df, today):
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

        total_cols = 4 + len(MONTHS) + 8 + 8
        ws.merge_range(0,0,0,total_cols-1, f"Book3 ↔ Pipeline & Awarded Mapping — {today.strftime('%d %B %Y')}", fmt_title)
        ws.set_row(0, 28)

        ws.write(1,0,"🟢 Strong match (score ≥ 0.70)", fmt_matched)
        ws.write(1,1,"🟡 Partial match (score 0.55–0.69)", fmt_partial)
        ws.write(1,2,"🔴 No match (score < 0.55)", fmt_nomatch)
        ws.write(1,3,"", fmt_text); ws.set_row(1,18)

        headers = (
            ["Book3 BU","Book3 Project Type","Book3 Project Name","Grand Total (QAR)"] +
            MONTHS +
            ["Pipeline Match","Score","Account","Gross (QAR)","Net (QAR)","Stage","AM","Quarter"] +
            ["Awarded Match","Score","Account","Gross (QAR)","Net (QAR)","Stage","AM","Year"]
        )
        widths = [36,22,42,18] + [10]*len(MONTHS) + [38,8,28,16,16,20,22,10] + [38,8,28,16,16,20,22,8]

        for c,(h,w) in enumerate(zip(headers, widths)):
            ws.write(2, c, h, fmt_mhdr if h in MONTHS else fmt_header)
            ws.set_column(c, c, w)
        ws.set_row(2, 20)
        ws.autofilter(2, 0, 2+len(map_df), len(headers)-1)

        for i, row in map_df.reset_index(drop=True).iterrows():
            sc = max(row["Pipeline Score"], row["Awarded Score"])
            if sc >= 0.70:
                ft, fn, fs, fn_neg = fmt_matched, fmt_matched_n, fmt_matched_s, fmt_neg_grn
            elif sc >= 0.55:
                ft, fn, fs, fn_neg = fmt_partial, fmt_partial_n, fmt_partial_s, fmt_neg_yel
            else:
                ft, fn, fs, fn_neg = fmt_nomatch, fmt_nomatch_n, fmt_nomatch_s, fmt_neg_red

            c = 0
            ws.write(3+i, c, str(row["Book3 BU"]) or "",           ft); c+=1
            ws.write(3+i, c, str(row["Book3 Project Type"]) or "", ft); c+=1
            ws.write(3+i, c, str(row["Book3 Project Name"]) or "", ft); c+=1
            ws.write_number(3+i, c, row["Book3 Grand Total"],   fn_neg); c+=1
            for m in MONTHS:
                ws.write_number(3+i, c, row.get(f"Book3 {m}", 0), fn_neg); c+=1
            # Pipeline
            ws.write(3+i, c, str(row["Pipeline Match"]) or "",    ft); c+=1
            ws.write_number(3+i, c, row["Pipeline Score"],        fs); c+=1
            ws.write(3+i, c, str(row["Pipeline Account"]) or "",  ft); c+=1
            ws.write_number(3+i, c, row["Pipeline Gross (QAR)"],  fn); c+=1
            ws.write_number(3+i, c, row["Pipeline Net (QAR)"],    fn); c+=1
            ws.write(3+i, c, str(row["Pipeline Stage"]) or "",    ft); c+=1
            ws.write(3+i, c, str(row["Pipeline AM"]) or "",       ft); c+=1
            ws.write(3+i, c, str(row["Pipeline Quarter"]) or "",  ft); c+=1
            # Awarded
            ws.write(3+i, c, str(row["Awarded Match"]) or "",     ft); c+=1
            ws.write_number(3+i, c, row["Awarded Score"],         fs); c+=1
            ws.write(3+i, c, str(row["Awarded Account"]) or "",   ft); c+=1
            ws.write_number(3+i, c, row["Awarded Gross (QAR)"],   fn); c+=1
            ws.write_number(3+i, c, row["Awarded Net (QAR)"],     fn); c+=1
            ws.write(3+i, c, str(row["Awarded Stage"]) or "",     ft); c+=1
            ws.write(3+i, c, str(row["Awarded AM"]) or "",        ft); c+=1
            ws.write(3+i, c, str(row["Awarded Year"]) or "",      ft); c+=1

    output.seek(0)
    return output.read()

# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════
st.title("🔗 Book3 Mapping Tool")
st.caption("Cross-reference your resource forecast against pipeline and awarded deals.")

with st.sidebar:
    st.header("📁 Upload Files")
    f_pipeline = st.file_uploader("Pipeline Excel",       type=["xlsx","xls"])
    f_awarded26= st.file_uploader("Awarded Deals 2026",   type=["xlsx","xls"])
    f_awarded25= st.file_uploader("Awarded Deals 2025",   type=["xlsx","xls"])
    f_book3    = st.file_uploader("Book3 (Resource Forecast)", type=["xlsx","xls"])

if not f_book3:
    st.info("👆 Upload at least the Book3 file to get started. Pipeline and Awarded are optional.")
    st.stop()

# Load data
with st.spinner("Loading files..."):
    book3_df   = load_book3(f_book3)
    pipeline_df = load_pipeline(f_pipeline) if f_pipeline else None

    aw_parts = []
    if f_awarded26: aw_parts.append(load_awarded(f_awarded26, "2026"))
    if f_awarded25: aw_parts.append(load_awarded(f_awarded25, "2025"))
    awarded_df = pd.concat(aw_parts, ignore_index=True) if aw_parts else None

    map_df = build_mapping(book3_df, pipeline_df, awarded_df)

TODAY = date.today()

# ── KPIs ─────────────────────────────────────────────────────────────────────
total   = len(map_df)
strong  = len(map_df[map_df[["Pipeline Score","Awarded Score"]].max(axis=1) >= 0.70])
partial = len(map_df[(map_df[["Pipeline Score","Awarded Score"]].max(axis=1) >= 0.55) &
                     (map_df[["Pipeline Score","Awarded Score"]].max(axis=1) <  0.70)])
nomatch = total - strong - partial

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Book3 Projects", total)
c2.metric("🟢 Strong Match",  strong)
c3.metric("🟡 Partial Match", partial)
c4.metric("🔴 No Match",      nomatch)

st.markdown("---")

# ── FILTERS ──────────────────────────────────────────────────────────────────
col_f1, col_f2, col_f3 = st.columns(3)
with col_f1:
    bu_opts = sorted(map_df["Book3 BU"].dropna().unique().tolist())
    sel_bu  = st.multiselect("Filter by BU", bu_opts, default=[])
with col_f2:
    type_opts = sorted(map_df["Book3 Project Type"].dropna().unique().tolist())
    sel_type  = st.multiselect("Filter by Project Type", type_opts, default=[])
with col_f3:
    match_opts = {"All": None, "🟢 Strong (≥0.70)": "strong", "🟡 Partial (0.55–0.69)": "partial", "🔴 No Match (<0.55)": "nomatch"}
    sel_match  = st.selectbox("Filter by Match Quality", list(match_opts.keys()))

view_df = map_df.copy()
if sel_bu:   view_df = view_df[view_df["Book3 BU"].isin(sel_bu)]
if sel_type: view_df = view_df[view_df["Book3 Project Type"].isin(sel_type)]
if match_opts[sel_match] == "strong":
    view_df = view_df[view_df[["Pipeline Score","Awarded Score"]].max(axis=1) >= 0.70]
elif match_opts[sel_match] == "partial":
    sc = view_df[["Pipeline Score","Awarded Score"]].max(axis=1)
    view_df = view_df[(sc >= 0.55) & (sc < 0.70)]
elif match_opts[sel_match] == "nomatch":
    view_df = view_df[view_df[["Pipeline Score","Awarded Score"]].max(axis=1) < 0.55]

st.caption(f"Showing {len(view_df)} of {total} projects")

# ── TABLE ────────────────────────────────────────────────────────────────────
def color_row(row):
    sc = max(row["Pipeline Score"], row["Awarded Score"])
    if sc >= 0.70:   color = "#E2EFDA"
    elif sc >= 0.55: color = "#FFF2CC"
    else:            color = "#FFE0E0"
    return [f"background-color: {color}"] * len(row)

display_cols = [
    "Book3 BU","Book3 Project Type","Book3 Project Name","Book3 Grand Total",
    "Pipeline Match","Pipeline Score","Pipeline Gross (QAR)","Pipeline Net (QAR)","Pipeline Stage","Pipeline AM","Pipeline Quarter",
    "Awarded Match","Awarded Score","Awarded Gross (QAR)","Awarded Net (QAR)","Awarded Stage","Awarded AM","Awarded Year",
]
styled = (
    view_df[display_cols]
    .style
    .apply(color_row, axis=1)
    .format({
        "Book3 Grand Total":    "{:,.0f}",
        "Pipeline Score":       "{:.2f}",
        "Pipeline Gross (QAR)": "{:,.0f}",
        "Pipeline Net (QAR)":   "{:,.0f}",
        "Awarded Score":        "{:.2f}",
        "Awarded Gross (QAR)":  "{:,.0f}",
        "Awarded Net (QAR)":    "{:,.0f}",
    })
)
st.dataframe(styled, use_container_width=True, height=500)

# ── EXPORT ────────────────────────────────────────────────────────────────────
st.markdown("---")
xl = export_mapping_excel(map_df, TODAY)
st.download_button(
    label="⬇️ Export Full Mapping to Excel",
    data=xl,
    file_name=f"Book3_Mapping_{TODAY}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
