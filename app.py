import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io

st.set_page_config(
    page_title="Exam Results Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://api.fontshare.com/v2/css?f[]=satoshi@400,500,700&display=swap');

:root {
    --primary: #01696f;
    --surface: #f9f8f5;
    --border: #d4d1ca;
    --text-muted: #7a7974;
}

html, body, [class*="css"] {
    font-family: 'Satoshi', 'Inter', sans-serif;
}

.main > div { padding-top: 1.5rem; }

/* Branch card */
.branch-card {
    background: #ffffff;
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
}
.branch-name {
    font-size: 1.1rem;
    font-weight: 700;
    color: #28251d;
    margin-bottom: 0.5rem;
    letter-spacing: -0.01em;
}
.stat-row {
    display: flex;
    gap: 2rem;
    flex-wrap: wrap;
}
.stat-item {
    display: flex;
    flex-direction: column;
}
.stat-label {
    font-size: 0.72rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    color: var(--text-muted);
    font-weight: 500;
}
.stat-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #28251d;
    line-height: 1.2;
}
.stat-sub {
    font-size: 0.82rem;
    color: var(--text-muted);
}
.section-header {
    font-size: 1.3rem;
    font-weight: 700;
    color: #28251d;
    border-bottom: 2px solid var(--primary);
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem 0;
    letter-spacing: -0.02em;
}
.range-row {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.55rem 0;
    border-bottom: 1px solid #f0efeb;
}
.range-label {
    font-size: 0.85rem;
    font-weight: 600;
    min-width: 110px;
    color: #28251d;
}
.range-bar-bg {
    flex: 1;
    height: 10px;
    background: #edeae5;
    border-radius: 999px;
    overflow: hidden;
}
.range-bar-fill {
    height: 100%;
    border-radius: 999px;
    background: var(--primary);
    transition: width 0.4s ease;
}
.range-count {
    min-width: 50px;
    text-align: right;
    font-weight: 700;
    font-size: 0.9rem;
    color: #28251d;
}
.range-pct {
    min-width: 48px;
    text-align: right;
    font-size: 0.78rem;
    color: var(--text-muted);
}
div[data-testid="metric-container"] {
    background: white;
    border: 1px solid #d4d1ca;
    border-radius: 10px;
    padding: 12px 16px;
}
</style>
""", unsafe_allow_html=True)


# ── Helpers ─────────────────────────────────────────────────────────────────

DEFAULT_RANGES = [(0, 20), (21, 40), (41, 50), (51, 60), (61, 70), (71, 80), (81, 100)]
RANGE_COLORS   = ["#a12c7b","#964219","#da7101","#d19900","#437a22","#01696f","#006494"]

@st.cache_data
def load_data(uploaded_file):
    raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    data = raw.iloc[1:, :3].copy()
    data.columns = ["branch", "roll_no", "marks"]
    data = data.dropna(subset=["branch", "roll_no"])
    data["branch"] = data["branch"].astype(str).str.strip()
    data = data[data["branch"].str.lower() != "branch"]
    data = data[data["branch"].str.lower() != "brahch"]
    data = data[data["branch"] != "nan"]
    return data

def parse_ranges(ranges_text):
    """Parse ranges like '0-20,21-40,...' or use defaults."""
    ranges = []
    for part in ranges_text.replace(" ", "").split(","):
        if "-" in part:
            try:
                lo, hi = part.split("-")
                ranges.append((int(lo), int(hi)))
            except:
                pass
    return ranges if ranges else DEFAULT_RANGES

def compute_branch_stats(data, ranges):
    branches = sorted(data["branch"].unique())
    results = {}
    for b in branches:
        bdf = data[data["branch"] == b].copy()
        total = len(bdf)
        absent = (bdf["marks"] == "AB").sum()
        present = total - absent
        numeric = pd.to_numeric(bdf["marks"], errors="coerce").dropna()

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
            "total": total,
            "absent": int(absent),
            "absent_pct": round(absent / total * 100, 1) if total else 0,
            "present": present,
            "passed": int((numeric >= 40).sum()),
            "avg": round(float(numeric.mean()), 1) if len(numeric) else 0,
            "max": int(numeric.max()) if len(numeric) else 0,
            "min": int(numeric.min()) if len(numeric) else 0,
            "range_counts": range_counts,
            "numeric": numeric,
        }
    return results

def bar_html(pct, color="#01696f"):
    return f"""
    <div class="range-bar-bg">
      <div class="range-bar-fill" style="width:{pct}%;background:{color};"></div>
    </div>"""

def render_branch_card(branch, stats, ranges, colors):
    st.markdown(f"""
    <div class="branch-card">
      <div class="branch-name">{branch}</div>
      <div class="stat-row">
        <div class="stat-item">
          <span class="stat-label">Total Students</span>
          <span class="stat-value">{stats['total']}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Present</span>
          <span class="stat-value">{stats['present']}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Absentees</span>
          <span class="stat-value" style="color:#964219">{stats['absent']}</span>
          <span class="stat-sub">{stats['absent_pct']}% of total</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Avg Score</span>
          <span class="stat-value">{stats['avg']}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Highest</span>
          <span class="stat-value" style="color:#437a22">{stats['max']}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Lowest</span>
          <span class="stat-value" style="color:#a12c7b">{stats['min']}</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    max_count = max((r["count"] for r in stats["range_counts"]), default=1) or 1
    st.markdown("<div style='margin-top:0.25rem;'>", unsafe_allow_html=True)
    for i, rc in enumerate(stats["range_counts"]):
        pct_bar = rc["count"] / max_count * 100
        color = colors[i % len(colors)]
        st.markdown(f"""
        <div class="range-row">
          <span class="range-label">{rc['label']} marks</span>
          {bar_html(pct_bar, color)}
          <span class="range-count">{rc['count']}</span>
          <span class="range-pct">{rc['pct']}%</span>
        </div>
        """, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## \U0001f4c2 Upload File")
    uploaded = st.file_uploader(
        "Upload .xlsx result file",
        type=["xlsx"],
        help="File must have columns: Branch | Roll No | Marks (AB for absent)"
    )

    st.markdown("---")
    st.markdown("## \U0001f3af Mark Ranges")
    st.markdown("Comma-separated ranges (e.g. `0-20,21-40`)")
    ranges_input = st.text_input(
        "Define ranges",
        value="0-20,21-40,41-50,51-60,61-70,71-80,81-100",
        label_visibility="collapsed"
    )
    ranges = parse_ranges(ranges_input)
    st.caption(f"\u2705 {len(ranges)} ranges parsed")

    st.markdown("---")
    st.markdown("## \U0001f50d Filter")
    if uploaded:
        raw_data = load_data(uploaded)
        all_branches = sorted(raw_data["branch"].unique())
        selected_branches = st.multiselect(
            "Select branches",
            all_branches,
            default=all_branches
        )
    else:
        selected_branches = []

    st.markdown("---")
    st.markdown("## \U0001f4e5 Export")
    export_btn = st.button("Download CSV Report", use_container_width=True)


# ── Main Content ─────────────────────────────────────────────────────────────
st.markdown("""
<div style='display:flex;align-items:center;gap:0.75rem;margin-bottom:0.25rem;'>
  <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
    <rect width="32" height="32" rx="8" fill="#01696f"/>
    <path d="M8 22 L14 14 L18 18 L22 10 L26 14" stroke="white" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" fill="none"/>
    <circle cx="8" cy="22" r="2" fill="white"/>
    <circle cx="26" cy="14" r="2" fill="white"/>
  </svg>
  <div>
    <div style='font-size:1.5rem;font-weight:800;color:#28251d;letter-spacing:-0.03em;line-height:1'>Exam Results Analyzer</div>
    <div style='font-size:0.8rem;color:#7a7974;'>Branch-wise analysis with custom mark ranges</div>
  </div>
</div>
<hr style='margin:0.75rem 0 1.25rem 0;border:none;border-top:1px solid #dcd9d5;'>
""", unsafe_allow_html=True)

if not uploaded:
    col1, col2 = st.columns([2, 1])
    with col1:
        st.info("\U0001f448 Upload an .xlsx result file from the sidebar to get started.")
        st.markdown("""
        **Expected file format:**
        | Column 1 | Column 2 | Column 3 |
        |----------|----------|----------|
        | Branch (e.g. CSE) | Roll Number | Marks (number or AB) |

        The app automatically:
        - Detects all branches
        - Counts absentees (marked `AB`)
        - Groups marks by your defined ranges
        - Generates visual charts
        """)
    st.stop()

raw_data = load_data(uploaded)
if not selected_branches:
    st.warning("No branches selected. Select at least one branch from the sidebar.")
    st.stop()

data = raw_data[raw_data["branch"].isin(selected_branches)]
stats = compute_branch_stats(data, ranges)

# ── Summary Metrics ──────────────────────────────────────────────────────────
total_students  = sum(s["total"]  for s in stats.values())
total_absent    = sum(s["absent"] for s in stats.values())
total_present   = sum(s["present"] for s in stats.values())
all_numeric     = pd.concat([s["numeric"] for s in stats.values()])
overall_avg     = round(float(all_numeric.mean()), 1) if len(all_numeric) else 0
total_pass      = sum(s["passed"] for s in stats.values())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Students",  f"{total_students:,}")
c2.metric("Total Branches",  len(stats))
c3.metric("Total Absentees", f"{total_absent}",  f"{round(total_absent/total_students*100,1)}%")
c4.metric("Overall Avg",     f"{overall_avg}")
c5.metric("Pass Rate (\u226540)", f"{round(total_pass/total_present*100,1) if total_present else 0}%")

st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

# ── Tabs ─────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["\U0001f4cb Branch-wise Detail", "\U0001f4ca Charts", "\U0001f4d1 Summary Table"])

with tab1:
    st.markdown("<div class='section-header'>Branch-wise Results</div>", unsafe_allow_html=True)
    cols_per_row = 2
    branch_list = list(stats.items())
    for i in range(0, len(branch_list), cols_per_row):
        row = branch_list[i:i+cols_per_row]
        cols = st.columns(len(row))
        for col, (branch, bstats) in zip(cols, row):
            with col:
                render_branch_card(branch, bstats, ranges, RANGE_COLORS)

with tab2:
    chart_data = {b: s for b, s in stats.items()}
    branches_ch = list(chart_data.keys())
    totals_ch   = [chart_data[b]["total"]  for b in branches_ch]
    absent_ch   = [chart_data[b]["absent"] for b in branches_ch]
    avg_ch      = [chart_data[b]["avg"]    for b in branches_ch]

    fig1 = go.Figure()
    fig1.add_trace(go.Bar(name="Present", x=branches_ch,
        y=[t - a for t,a in zip(totals_ch, absent_ch)],
        marker_color="#01696f", marker_line_width=0))
    fig1.add_trace(go.Bar(name="Absent", x=branches_ch, y=absent_ch,
        marker_color="#964219", marker_line_width=0))
    fig1.update_layout(
        barmode="stack", title="Students per Branch (Present vs Absent)",
        font_family="Satoshi, Inter, sans-serif",
        paper_bgcolor="white", plot_bgcolor="white",
        legend=dict(orientation="h", y=1.12),
        height=380,
        margin=dict(t=60, b=40, l=40, r=20),
        xaxis=dict(gridcolor="#f0efeb"), yaxis=dict(gridcolor="#f0efeb")
    )
    st.plotly_chart(fig1, use_container_width=True)

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        fig2 = go.Figure(go.Bar(
            x=branches_ch, y=avg_ch,
            marker_color=RANGE_COLORS[:len(branches_ch)],
            marker_line_width=0,
            text=[f"{v}" for v in avg_ch],
            textposition="outside"
        ))
        fig2.update_layout(
            title="Average Score per Branch",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white",
            height=320, margin=dict(t=50, b=40, l=40, r=20),
            yaxis_range=[0, 110],
            xaxis=dict(gridcolor="#f0efeb"), yaxis=dict(gridcolor="#f0efeb")
        )
        st.plotly_chart(fig2, use_container_width=True)

    with col_c2:
        ab_pcts = [chart_data[b]["absent_pct"] for b in branches_ch]
        fig3 = go.Figure(go.Bar(
            x=branches_ch, y=ab_pcts,
            marker_color="#da7101", marker_line_width=0,
            text=[f"{v}%" for v in ab_pcts],
            textposition="outside"
        ))
        fig3.update_layout(
            title="Absentee % per Branch",
            font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white",
            height=320, margin=dict(t=50, b=40, l=40, r=20),
            yaxis=dict(ticksuffix="%", gridcolor="#f0efeb"),
            xaxis=dict(gridcolor="#f0efeb")
        )
        st.plotly_chart(fig3, use_container_width=True)

    st.markdown("#### Marks Distribution Heatmap (all branches)")
    range_labels = [r["label"] for r in list(stats.values())[0]["range_counts"]]
    heatmap_z, heatmap_text = [], []
    for b in branches_ch:
        row_z, row_t = [], []
        for rc in stats[b]["range_counts"]:
            row_z.append(rc["count"])
            row_t.append(f"{rc['count']} ({rc['pct']}%)")
        heatmap_z.append(row_z)
        heatmap_text.append(row_t)

    fig4 = go.Figure(go.Heatmap(
        z=heatmap_z, x=range_labels, y=branches_ch,
        text=heatmap_text, texttemplate="%{text}",
        colorscale=[[0,"#f7f6f2"],[0.5,"#cedcd8"],[1,"#01696f"]],
        showscale=True,
        hoverongaps=False
    ))
    fig4.update_layout(
        font_family="Satoshi, Inter, sans-serif",
        paper_bgcolor="white", height=280 + len(branches_ch)*20,
        margin=dict(t=20, b=40, l=60, r=20)
    )
    st.plotly_chart(fig4, use_container_width=True)

    if len(all_numeric) > 10:
        st.markdown("#### Score Distribution by Branch")
        violin_rows = []
        for b in branches_ch:
            for v in stats[b]["numeric"]:
                violin_rows.append({"branch": b, "score": v})
        vdf = pd.DataFrame(violin_rows)
        fig5 = px.violin(
            vdf, x="branch", y="score", box=True, points="outliers",
            color="branch",
            color_discrete_sequence=RANGE_COLORS
        )
        fig5.update_layout(
            showlegend=False, font_family="Satoshi, Inter, sans-serif",
            paper_bgcolor="white", plot_bgcolor="white",
            height=400, margin=dict(t=20, b=40, l=40, r=20),
            xaxis=dict(gridcolor="#f0efeb"), yaxis=dict(gridcolor="#f0efeb")
        )
        st.plotly_chart(fig5, use_container_width=True)

with tab3:
    range_labels_tab = [r["label"] for r in list(stats.values())[0]["range_counts"]]
    rows = []
    for b, s in stats.items():
        row = {
            "Branch": b,
            "Total": s["total"],
            "Present": s["present"],
            "Absent": s["absent"],
            "Absent %": f"{s['absent_pct']}%",
            "Avg Score": s["avg"],
            "Highest": s["max"],
            "Lowest": s["min"],
        }
        for rc in s["range_counts"]:
            row[f"{rc['label']}"] = f"{rc['count']} ({rc['pct']}%)"
        rows.append(row)

    summary_df = pd.DataFrame(rows)
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    if export_btn:
        csv = summary_df.to_csv(index=False)
        st.download_button(
            "\u2b07\ufe0f Download CSV",
            data=csv,
            file_name="exam_results_summary.csv",
            mime="text/csv"
        )

st.markdown("""
<div style='text-align:center;color:#bab9b4;font-size:0.75rem;margin-top:2rem;padding-top:1rem;border-top:1px solid #dcd9d5;'>
  Exam Results Analyzer — Upload any .xlsx file in the same format
</div>
""", unsafe_allow_html=True)
