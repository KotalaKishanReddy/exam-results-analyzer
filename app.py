import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Exam Results Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://api.fontshare.com/v2/css?f[]=satoshi@400,500,700,900&display=swap');

html, body, [class*="css"] {
    font-family: 'Satoshi', 'Inter', sans-serif !important;
}

/* ── Page background ── */
.stApp { background: #f4f3ef; }
section[data-testid="stSidebar"] { background: #1c1b19 !important; }
section[data-testid="stSidebar"] * { color: #cdccca !important; }
section[data-testid="stSidebar"] .stMarkdown h2 {
    color: #4f98a3 !important;
    font-size: 0.7rem !important;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-top: 1.25rem !important;
}
section[data-testid="stSidebar"] hr {
    border-color: #2d2c2a !important;
}
section[data-testid="stSidebar"] .stButton button {
    background: #01696f !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
}
section[data-testid="stSidebar"] .stButton button:hover {
    background: #0c4e54 !important;
}

/* ── Metrics ── */
div[data-testid="metric-container"] {
    background: white;
    border: 1px solid #dcd9d5;
    border-radius: 12px;
    padding: 1rem 1.25rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
}
div[data-testid="metric-container"] label {
    font-size: 0.7rem !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #7a7974 !important;
    font-weight: 600 !important;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 1.9rem !important;
    font-weight: 800 !important;
    color: #28251d !important;
    letter-spacing: -0.03em;
}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 0;
    background: white;
    border-radius: 10px;
    padding: 4px;
    border: 1px solid #dcd9d5;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 7px;
    padding: 0.4rem 1.1rem;
    font-weight: 600;
    font-size: 0.82rem;
    color: #7a7974;
    background: transparent;
    border: none;
}
.stTabs [aria-selected="true"] {
    background: #01696f !important;
    color: white !important;
}

/* ── Branch card ── */
.branch-card {
    background: white;
    border: 1px solid #dcd9d5;
    border-radius: 14px;
    padding: 1.25rem 1.4rem 0.9rem;
    margin-bottom: 1rem;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.branch-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 0.85rem;
}
.branch-name {
    font-size: 1rem;
    font-weight: 800;
    color: #28251d;
    letter-spacing: -0.01em;
}
.branch-pill {
    background: #f0f9f9;
    color: #01696f;
    border: 1px solid #b8d8d9;
    border-radius: 999px;
    padding: 0.15rem 0.65rem;
    font-size: 0.72rem;
    font-weight: 700;
}
.stat-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 0.5rem;
    margin-bottom: 1rem;
}
.stat-box {
    background: #f7f6f2;
    border-radius: 8px;
    padding: 0.5rem 0.6rem;
    display: flex;
    flex-direction: column;
}
.stat-label {
    font-size: 0.62rem;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    color: #7a7974;
    font-weight: 600;
    margin-bottom: 0.1rem;
}
.stat-value {
    font-size: 1.15rem;
    font-weight: 800;
    color: #28251d;
    letter-spacing: -0.02em;
    line-height: 1.2;
}
.stat-sub { font-size: 0.68rem; color: #bab9b4; }

/* ── Range bars ── */
.range-row {
    display: flex;
    align-items: center;
    gap: 0.65rem;
    padding: 0.45rem 0;
    border-bottom: 1px solid #f4f3ef;
}
.range-row:last-child { border-bottom: none; }
.range-label {
    font-size: 0.78rem;
    font-weight: 700;
    min-width: 90px;
    color: #28251d;
}
.range-bar-bg {
    flex: 1;
    height: 7px;
    background: #ede9e4;
    border-radius: 999px;
    overflow: hidden;
}
.range-bar-fill {
    height: 100%;
    border-radius: 999px;
}
.range-count {
    min-width: 38px;
    text-align: right;
    font-weight: 800;
    font-size: 0.82rem;
    color: #28251d;
}
.range-pct {
    min-width: 42px;
    text-align: right;
    font-size: 0.72rem;
    color: #7a7974;
}
.section-header {
    font-size: 1.1rem;
    font-weight: 800;
    color: #28251d;
    letter-spacing: -0.02em;
    margin: 0.5rem 0 1rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.section-header::after {
    content: '';
    flex: 1;
    height: 1px;
    background: #dcd9d5;
    margin-left: 0.5rem;
}
</style>
""", unsafe_allow_html=True)


# ── Constants ───────────────────────────────────────────────────────────────────
DEFAULT_RANGES = [(0, 20), (21, 40), (41, 50), (51, 60), (61, 70), (71, 80), (81, 100)]
RANGE_COLORS   = ["#a12c7b", "#964219", "#da7101", "#d19900", "#437a22", "#01696f", "#006494"]
TEAL = "#01696f"


# ── Data helpers ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    data = raw.iloc[1:, :3].copy()
    data.columns = ["branch", "roll_no", "marks"]
    data = data.dropna(subset=["branch", "roll_no"])
    data["branch"] = data["branch"].astype(str).str.strip()
    bad = {"branch", "brahch", "nan", "none", ""}
    data = data[~data["branch"].str.lower().isin(bad)]
    return data

def parse_ranges(text):
    ranges = []
    for part in text.replace(" ", "").split(","):
        if "-" in part:
            try:
                lo, hi = part.split("-")
                ranges.append((int(lo), int(hi)))
            except:
                pass
    return ranges or DEFAULT_RANGES

def compute_stats(data, ranges, pass_mark):
    results = {}
    for b in sorted(data["branch"].unique()):
        bdf = data[data["branch"] == b]
        total   = len(bdf)
        absent  = int((bdf["marks"] == "AB").sum())
        present = total - absent
        numeric = pd.to_numeric(bdf["marks"], errors="coerce").dropna()
        passed  = int((numeric >= pass_mark).sum())

        range_counts = []
        for lo, hi in ranges:
            cnt = int(((numeric >= lo) & (numeric <= hi)).sum())
            range_counts.append({
                "label": f"{lo}\u2013{hi}",
                "lo": lo, "hi": hi,
                "count": cnt,
                "pct": round(cnt / present * 100, 1) if present else 0
            })

        results[b] = {
            "total":       total,
            "absent":      absent,
            "absent_pct":  round(absent / total * 100, 1) if total else 0,
            "present":     present,
            "passed":      passed,
            "pass_pct":    round(passed / present * 100, 1) if present else 0,
            "avg":         round(float(numeric.mean()), 1) if len(numeric) else 0,
            "max":         int(numeric.max()) if len(numeric) else 0,
            "min":         int(numeric.min()) if len(numeric) else 0,
            "range_counts": range_counts,
            "numeric":     numeric,
        }
    return results


# ── Excel export ───────────────────────────────────────────────────────────────────
def build_excel(stats, pass_mark):
    wb = Workbook()
    ws = wb.active
    ws.title = "Branch Summary"

    teal_fill   = PatternFill("solid", fgColor="01696F")
    grey_fill   = PatternFill("solid", fgColor="F4F3EF")
    alt_fill    = PatternFill("solid", fgColor="FFFFFF")
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    body_font   = Font(name="Calibri", size=10)
    bold_font   = Font(name="Calibri", bold=True, size=10)
    center      = Alignment(horizontal="center", vertical="center")
    left        = Alignment(horizontal="left",   vertical="center")
    thin        = Side(style="thin", color="DCD9D5")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    range_labels = [rc["label"] for rc in list(stats.values())[0]["range_counts"]]
    base_headers = ["Branch", "Total", "Present", "Absent", "Absent %",
                    f"Pass (\u2265{pass_mark})", "Pass %", "Avg", "Highest", "Lowest"]
    all_headers  = base_headers + [f"{r} marks" for r in range_labels]

    ws.row_dimensions[1].height = 28
    for col_idx, h in enumerate(all_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = teal_fill; cell.font = header_font
        cell.alignment = center; cell.border = border

    for row_idx, (branch, s) in enumerate(stats.items(), 2):
        fill = grey_fill if row_idx % 2 == 0 else alt_fill
        ws.row_dimensions[row_idx].height = 20
        row_data = [
            branch, s["total"], s["present"], s["absent"],
            f"{s['absent_pct']}%", s["passed"], f"{s['pass_pct']}%",
            s["avg"], s["max"], s["min"],
        ]
        for rc in s["range_counts"]:
            row_data.append(f"{rc['count']}  ({rc['pct']}%)")
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = fill
            cell.font = bold_font if col_idx == 1 else body_font
            cell.alignment = left if col_idx == 1 else center
            cell.border = border

    col_widths = [14,9,9,9,10,12,9,7,9,9] + [15]*len(range_labels)
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"

    for branch, s in stats.items():
        ws2 = wb.create_sheet(title=branch[:30])
        ws2.append(["Range", "Count", "% of Present"])
        for cell in ws2[1]: cell.font = header_font; cell.fill = teal_fill
        for rc in s["range_counts"]:
            ws2.append([rc["label"], rc["count"], f"{rc['pct']}%"])
        ws2.column_dimensions["A"].width = 14
        ws2.column_dimensions["B"].width = 10
        ws2.column_dimensions["C"].width = 16

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ── Card renderer ──────────────────────────────────────────────────────────────────
def stat_box(label, value, color="#28251d", sub=None):
    sub_html = f"<span class='stat-sub'>{sub}</span>" if sub else ""
    return f"""
    <div class="stat-box">
      <span class="stat-label">{label}</span>
      <span class="stat-value" style="color:{color}">{value}</span>
      {sub_html}
    </div>"""

def render_card(branch, s, colors):
    stat_grid = "".join([
        stat_box("Total",   s["total"]),
        stat_box("Present", s["present"]),
        stat_box("Absent",  s["absent"],  "#964219", f"{s['absent_pct']}%"),
        stat_box("Average", s["avg"]),
        stat_box("Highest", s["max"],     "#437a22"),
    ])
    max_cnt = max((r["count"] for r in s["range_counts"]), default=1) or 1
    bars = ""
    for i, rc in enumerate(s["range_counts"]):
        w = rc["count"] / max_cnt * 100
        clr = colors[i % len(colors)]
        bars += f"""
        <div class="range-row">
          <span class="range-label">{rc['label']}</span>
          <div class="range-bar-bg">
            <div class="range-bar-fill" style="width:{w}%;background:{clr};"></div>
          </div>
          <span class="range-count">{rc['count']}</span>
          <span class="range-pct">{rc['pct']}%</span>
        </div>"""
    st.markdown(f"""
    <div class="branch-card">
      <div class="branch-header">
        <span class="branch-name">{branch}</span>
        <span class="branch-pill">Pass {s['pass_pct']}%</span>
      </div>
      <div class="stat-grid">{stat_grid}</div>
      <div>{bars}</div>
    </div>
    """, unsafe_allow_html=True)


# ── Sidebar ─────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:0.6rem;padding:0.5rem 0 1rem;">
      <svg width="28" height="28" viewBox="0 0 32 32" fill="none">
        <rect width="32" height="32" rx="8" fill="#01696f"/>
        <path d="M7 22 L13 14 L17 18 L21 10 L25 14" stroke="white" stroke-width="2.2"
              stroke-linecap="round" stroke-linejoin="round" fill="none"/>
        <circle cx="7" cy="22" r="2" fill="white"/>
        <circle cx="25" cy="14" r="2" fill="white"/>
      </svg>
      <div>
        <div style="font-size:0.9rem;font-weight:800;color:#cdccca;letter-spacing:-0.01em;">Results Analyzer</div>
        <div style="font-size:0.65rem;color:#5a5957;">Exam Analytics</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("## Upload")
    uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"],
        help="Branch | Roll No | Marks (number or AB)")

    st.markdown("## Pass Mark")
    pass_mark = st.number_input("Pass mark threshold", min_value=0, max_value=100,
        value=40, step=1, label_visibility="collapsed")
    st.caption(f"Students scoring \u2265 {pass_mark} are counted as passed")

    st.markdown("## Mark Ranges")
    ranges_input = st.text_input("Ranges",
        value="0-20,21-40,41-50,51-60,61-70,71-80,81-100",
        label_visibility="collapsed")
    ranges = parse_ranges(ranges_input)
    st.caption(f"\u2705 {len(ranges)} ranges active")

    st.markdown("## Filter")
    if uploaded:
        raw_data = load_data(uploaded)
        all_branches = sorted(raw_data["branch"].unique())
        selected = st.multiselect("Branches", all_branches, default=all_branches,
                                  label_visibility="collapsed")
    else:
        selected = []

    st.markdown("## Export")
    dl_xlsx = st.button("\u2b07  Download Excel Report", use_container_width=True)


# ── Header ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-bottom:1.25rem;">
  <div style="font-size:1.6rem;font-weight:900;color:#28251d;
              letter-spacing:-0.04em;line-height:1.1;">
    Exam Results Analyzer
  </div>
  <div style="font-size:0.82rem;color:#7a7974;margin-top:0.2rem;">
    Branch-wise breakdown · Absentee tracking · Custom mark ranges
  </div>
</div>
""", unsafe_allow_html=True)

if not uploaded:
    st.markdown("""
    <div style="background:white;border:1px solid #dcd9d5;border-radius:14px;
                padding:2.5rem 2rem;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,0.05);">
      <div style="font-size:2.5rem;margin-bottom:0.75rem;">📂</div>
      <div style="font-size:1.1rem;font-weight:800;color:#28251d;margin-bottom:0.4rem;">
        Upload an Excel results file to begin
      </div>
      <div style="color:#7a7974;font-size:0.85rem;max-width:420px;margin:0 auto;">
        File must have 3 columns: <strong>Branch</strong> · <strong>Roll No</strong> ·
        <strong>Marks</strong> (number or <code>AB</code> for absent).
        Header row is auto-skipped.
      </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

raw_data = load_data(uploaded)
if not selected:
    st.warning("Select at least one branch from the sidebar.")
    st.stop()

data  = raw_data[raw_data["branch"].isin(selected)]
stats = compute_stats(data, ranges, pass_mark)

# ── KPI strip ──────────────────────────────────────────────────────────────────────
all_numeric = pd.concat([s["numeric"] for s in stats.values()])
total_s  = sum(s["total"]  for s in stats.values())
total_ab = sum(s["absent"] for s in stats.values())
total_pr = sum(s["present"] for s in stats.values())
total_pa = sum(s["passed"] for s in stats.values())
ovg_avg  = round(float(all_numeric.mean()), 1) if len(all_numeric) else 0

c1,c2,c3,c4,c5,c6 = st.columns(6)
c1.metric("Total Students",    f"{total_s:,}")
c2.metric("Branches",          len(stats))
c3.metric("Total Absentees",   f"{total_ab}", f"{round(total_ab/total_s*100,1)}%")
c4.metric("Pass Mark",         f"\u2265 {pass_mark}")
c5.metric("Overall Pass Rate", f"{round(total_pa/total_pr*100,1) if total_pr else 0}%")
c6.metric("Overall Avg",       f"{ovg_avg}")

st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)

if dl_xlsx:
    xlsx_bytes = build_excel(stats, pass_mark)
    st.download_button(
        label="📥 Click to save Excel file",
        data=xlsx_bytes,
        file_name="exam_results_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="xlsx_dl"
    )

# ── Tabs ────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["  📋  Branch Detail  ", "  📊  Charts  ", "  📑  Summary Table  "])

with tab1:
    st.markdown("<div class='section-header'>Branch-wise Results</div>", unsafe_allow_html=True)
    branch_list = list(stats.items())
    for i in range(0, len(branch_list), 2):
        row = branch_list[i:i+2]
        cols = st.columns(len(row))
        for col, (branch, bstats) in zip(cols, row):
            with col:
                render_card(branch, bstats, RANGE_COLORS)

with tab2:
    branches_ch = list(stats.keys())
    totals_ch   = [stats[b]["total"]    for b in branches_ch]
    absent_ch   = [stats[b]["absent"]   for b in branches_ch]
    avg_ch      = [stats[b]["avg"]      for b in branches_ch]
    pass_pct_ch = [stats[b]["pass_pct"] for b in branches_ch]

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        fig1 = go.Figure()
        fig1.add_trace(go.Bar(name="Present", x=branches_ch,
            y=[t-a for t,a in zip(totals_ch,absent_ch)],
            marker_color=TEAL, marker_line_width=0))
        fig1.add_trace(go.Bar(name="Absent", x=branches_ch, y=absent_ch,
            marker_color="#964219", marker_line_width=0))
        fig1.update_layout(barmode="stack", title="Attendance by Branch",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white",
            legend=dict(orientation="h",y=1.12), height=340,
            margin=dict(t=55,b=35,l=35,r=15),
            xaxis=dict(gridcolor="#f0efeb"), yaxis=dict(gridcolor="#f0efeb"))
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        fig2 = go.Figure(go.Bar(
            x=branches_ch, y=pass_pct_ch,
            marker_color=[TEAL if p>=50 else "#964219" for p in pass_pct_ch],
            marker_line_width=0,
            text=[f"{v}%" for v in pass_pct_ch], textposition="outside"
        ))
        fig2.add_hline(y=50, line_dash="dot", line_color="#7a7974", line_width=1)
        fig2.update_layout(title=f"Pass Rate per Branch (\u2265{pass_mark} marks)",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white", height=340,
            margin=dict(t=55,b=35,l=35,r=15),
            yaxis=dict(ticksuffix="%",range=[0,115],gridcolor="#f0efeb"),
            xaxis=dict(gridcolor="#f0efeb"))
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        fig3 = go.Figure(go.Bar(
            x=branches_ch, y=avg_ch,
            marker_color=RANGE_COLORS[:len(branches_ch)],
            marker_line_width=0, text=avg_ch, textposition="outside"
        ))
        fig3.update_layout(title="Average Score per Branch",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white", height=300,
            margin=dict(t=50,b=35,l=35,r=15),
            yaxis=dict(range=[0,115],gridcolor="#f0efeb"),
            xaxis=dict(gridcolor="#f0efeb"))
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        ab_pcts = [stats[b]["absent_pct"] for b in branches_ch]
        fig4 = go.Figure(go.Bar(
            x=branches_ch, y=ab_pcts, marker_color="#da7101", marker_line_width=0,
            text=[f"{v}%" for v in ab_pcts], textposition="outside"
        ))
        fig4.update_layout(title="Absentee % per Branch",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white", height=300,
            margin=dict(t=50,b=35,l=35,r=15),
            yaxis=dict(ticksuffix="%",gridcolor="#f0efeb"),
            xaxis=dict(gridcolor="#f0efeb"))
        st.plotly_chart(fig4, use_container_width=True)

    st.markdown("<div class='section-header' style='margin-top:0.5rem'>Marks Distribution Heatmap</div>",
                unsafe_allow_html=True)
    range_labels = [r["label"] for r in list(stats.values())[0]["range_counts"]]
    hz, ht = [], []
    for b in branches_ch:
        rz, rt = [], []
        for rc in stats[b]["range_counts"]:
            rz.append(rc["count"]); rt.append(f"{rc['count']} ({rc['pct']}%)")
        hz.append(rz); ht.append(rt)
    fig5 = go.Figure(go.Heatmap(
        z=hz, x=range_labels, y=branches_ch, text=ht, texttemplate="%{text}",
        colorscale=[[0,"#f7f6f2"],[0.4,"#cedcd8"],[1,"#01696f"]],
        showscale=True, hoverongaps=False
    ))
    fig5.update_layout(font_family="Satoshi, Inter, sans-serif",
        paper_bgcolor="white", height=260+len(branches_ch)*22,
        margin=dict(t=15,b=40,l=70,r=20))
    st.plotly_chart(fig5, use_container_width=True)

    if len(all_numeric) > 20:
        st.markdown("<div class='section-header'>Score Distribution (Violin)</div>",
                    unsafe_allow_html=True)
        vrows = [{"branch":b,"score":v} for b in branches_ch for v in stats[b]["numeric"]]
        vdf = pd.DataFrame(vrows)
        fig6 = px.violin(vdf, x="branch", y="score", box=True, points="outliers",
                         color="branch", color_discrete_sequence=RANGE_COLORS)
        fig6.update_layout(showlegend=False, font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white", height=400,
            margin=dict(t=15,b=40,l=40,r=20),
            xaxis=dict(gridcolor="#f0efeb"), yaxis=dict(gridcolor="#f0efeb"))
        st.plotly_chart(fig6, use_container_width=True)

with tab3:
    rows = []
    for b, s in stats.items():
        row = {
            "Branch": b, "Total": s["total"], "Present": s["present"],
            "Absent": s["absent"], "Absent %": f"{s['absent_pct']}%",
            f"Pass (\u2265{pass_mark})": s["passed"], "Pass %": f"{s['pass_pct']}%",
            "Avg": s["avg"], "High": s["max"], "Low": s["min"],
        }
        for rc in s["range_counts"]:
            row[rc["label"]] = f"{rc['count']} ({rc['pct']}%)"
        rows.append(row)
    summary_df = pd.DataFrame(rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True, height=420)

    col_dl1, col_dl2, _ = st.columns([1,1,3])
    with col_dl1:
        st.download_button("\u2b07 CSV",
            data=summary_df.to_csv(index=False).encode(),
            file_name="exam_results_summary.csv", mime="text/csv")
    with col_dl2:
        st.download_button("\u2b07 Excel",
            data=build_excel(stats, pass_mark),
            file_name="exam_results_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="xlsx_tab3")

st.markdown("""
<div style="text-align:center;color:#bab9b4;font-size:0.72rem;
            margin-top:2rem;padding-top:1rem;border-top:1px solid #dcd9d5;">
  Exam Results Analyzer — supports any exam cycle in the same .xlsx format
</div>
""", unsafe_allow_html=True)
