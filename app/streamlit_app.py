"""
streamlit_app.py — Interactive UI for the Data Quality Profiler.

Upload any .xlsx or .csv file and get an instant, per-column quality
report: inferred type, completeness, validity, uniqueness, outliers, and
a 0-100 quality score — with zero configuration and zero assumptions
about what the file contains.

Run with:
    streamlit run app/streamlit_app.py
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.profiler import profile_dataset

st.set_page_config(page_title="Data Quality Profiler", page_icon="🔍", layout="wide")

# ---------------------------------------------------------------- palette --
# Same accent family as the in-browser demo: electric violet + teal on a
# near-black ground, so the repo and the live demo read as one product.
BG = "#06070f"
SURFACE = "#0f1222"
BORDER = "#292f52"
TEXT = "#edeffb"
TEXT_MUTED = "#8d94ba"
ACCENT = "#8d6bff"
ACCENT_2 = "#22f0c6"
GOOD = "#3fe089"
WARN = "#ffb648"
BAD = "#ff5c80"

DIMENSION_COLORS = {"Completeness": "#5fc9ff", "Validity": "#ff7ad4", "Uniqueness": WARN, "None": GOOD}

TYPE_COLORS = {
    "email": "#5fc9ff", "url": "#5fc9ff", "date": "#b491ff",
    "integer": GOOD, "float": GOOD, "currency": GOOD, "percent": GOOD,
    "categorical": WARN, "boolean": WARN,
    "saudi_id": "#ff7ad4", "phone": "#ff7ad4", "free_text": TEXT_MUTED,
}

# ---------------------------------------------------------------- styling --
st.markdown(
    f"""
    <style>
    .stApp {{ background-color: {BG}; }}
    .metric-card {{
        background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 12px;
        padding: 1rem 1.2rem; text-align: center;
    }}
    .metric-card h2 {{ margin: 0; font-size: 1.9rem; }}
    .metric-card p {{ margin: 0; color: {TEXT_MUTED}; font-size: 0.85rem; }}
    .metric-card.small h2 {{ font-size: 1.4rem; }}
    .kpi-row-label {{
        color: {TEXT_MUTED}; font-size: 0.78rem; text-transform: uppercase;
        letter-spacing: 0.06em; margin: 1.1rem 0 0.4rem 0;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def score_color(score: float) -> str:
    if score >= 90:
        return GOOD
    if score >= 75:
        return WARN
    return BAD


def type_badge(t: str) -> str:
    color = TYPE_COLORS.get(t, TEXT_MUTED)
    return f'<span style="background:{color}22;color:{color};padding:2px 8px;border-radius:12px;font-size:0.78rem;font-weight:600;">{t}</span>'


def dimension_badge(dimension: str, pct: float) -> str:
    if dimension == "None":
        return f'<span style="background:{GOOD}22;color:{GOOD};padding:2px 8px;border-radius:12px;font-size:0.78rem;font-weight:600;">✓ no weak dimension</span>'
    color = DIMENSION_COLORS.get(dimension, TEXT_MUTED)
    return (
        f'<span style="background:{color}22;color:{color};padding:2px 8px;border-radius:12px;'
        f'font-size:0.78rem;font-weight:600;">⚠ weakest: {dimension} ({pct:.0f}%)</span>'
    )


def metric_card(label: str, value: str, color: str = TEXT, small: bool = False) -> str:
    cls = "metric-card small" if small else "metric-card"
    return f'<div class="{cls}"><h2 style="color:{color}">{value}</h2><p>{label}</p></div>'


def get_sheet_names(file_or_path) -> list:
    """Sheet names for an .xlsx/.xls source, [] for CSV (which has none)."""
    try:
        return pd.ExcelFile(file_or_path).sheet_names
    except Exception:
        return []


def load_dataframe(file_or_path, sheet_name=None) -> pd.DataFrame:
    name = getattr(file_or_path, "name", str(file_or_path))
    if name.lower().endswith(".csv"):
        return pd.read_csv(file_or_path, dtype=str, keep_default_na=False)
    return pd.read_excel(
        file_or_path, sheet_name=sheet_name or 0, dtype=str, keep_default_na=False
    )


# ------------------------------------------------------------------ header --
st.title("🔍 Data Quality Profiler")
st.caption(
    "Upload **any** spreadsheet — employee records, project trackers, sales data, "
    "anything with columns and rows. No schema, no config. The profiler infers what "
    "each column *should* look like and tells you exactly where it doesn't."
)
st.caption(
    "🌐 Prefer zero install? Try the [in-browser demo](https://claude.ai/code/artifact/dff15977-cbaf-4bca-9da4-d3f45b7cec77) — "
    "same engine, runs entirely client-side, nothing to set up."
)

with st.sidebar:
    st.header("Try it instantly")
    st.write("No file handy? Use a sample dataset:")
    sample = st.radio(
        "Sample dataset",
        ["— upload my own —", "Sample: Employees", "Sample: Projects"],
        label_visibility="collapsed",
    )
    st.divider()
    st.markdown(
        "**How it works**\n\n"
        "1. Every column is scored against ~10 pattern detectors "
        "(email, date, phone, currency, categorical, ...)\n"
        "2. The best-matching type is kept as the column's *inferred schema*\n"
        "3. Completeness, validity, uniqueness and outliers are measured "
        "against that inferred schema\n"
        "4. Nothing is hardcoded — the same engine works on any table."
    )
    st.divider()
    st.caption("Built by Omar Althunaiyan · Python, pandas, Streamlit, Plotly")

uploaded_file = None
if sample == "Sample: Employees":
    uploaded_file = Path(__file__).resolve().parents[1] / "examples" / "employees.xlsx"
elif sample == "Sample: Projects":
    uploaded_file = Path(__file__).resolve().parents[1] / "examples" / "projects.xlsx"
else:
    uploaded_file = st.file_uploader("Upload a .xlsx or .csv file", type=["xlsx", "xls", "csv"])

if uploaded_file is None:
    st.info("Upload a spreadsheet on the left, or pick a sample dataset to see the profiler in action.")
    st.stop()

base_file_name = uploaded_file.name  # pathlib.Path and Streamlit's UploadedFile both expose .name

# ------------------------------------------------------- multi-sheet picker --
sheet_names = get_sheet_names(uploaded_file) if not base_file_name.lower().endswith(".csv") else []
if not isinstance(uploaded_file, Path) and hasattr(uploaded_file, "seek"):
    uploaded_file.seek(0)  # ExcelFile() above consumed the buffer — rewind before the real read

selected_sheet = None
if len(sheet_names) > 1:
    with st.sidebar:
        st.divider()
        st.markdown("**Sheet**")
        selected_sheet = st.selectbox(
            "This workbook has multiple sheets — pick one to profile:",
            sheet_names,
            label_visibility="collapsed",
        )
    if not isinstance(uploaded_file, Path) and hasattr(uploaded_file, "seek"):
        uploaded_file.seek(0)

# ------------------------------------------------------------ run profiler --
df = load_dataframe(uploaded_file, sheet_name=selected_sheet)
file_label = f"{base_file_name} — {selected_sheet}" if selected_sheet else base_file_name

with st.spinner("Profiling..."):
    profile = profile_dataset(df)

st.subheader(f"Report for `{file_label}`")

# ------------------------------------------------------------- top metrics --
c1, c2, c3, c4 = st.columns(4)
for col, label, value, color in [
    (c1, "Overall Quality Score", f"{profile.overall_quality_score}", score_color(profile.overall_quality_score)),
    (c2, "Rows", f"{profile.row_count:,}", TEXT),
    (c3, "Columns Profiled", f"{profile.column_count}", TEXT),
    (c4, "Duplicate Rows", f"{profile.duplicate_rows}", BAD if profile.duplicate_rows else GOOD),
]:
    with col:
        st.markdown(metric_card(label, value, color), unsafe_allow_html=True)

# --------------------------------------------------------- dimension KPIs --
st.markdown('<p class="kpi-row-label">Quality by dimension</p>', unsafe_allow_html=True)
d1, d2, d3, d4, d5 = st.columns(5)
for col, label, value, color in [
    (d1, "Avg Completeness", f"{profile.avg_completeness}%", score_color(profile.avg_completeness)),
    (d2, "Avg Validity", f"{profile.avg_validity}%", score_color(profile.avg_validity)),
    (d3, "Avg Uniqueness", f"{profile.avg_uniqueness}%", score_color(profile.avg_uniqueness)),
    (d4, "Columns With Issues", f"{profile.columns_with_issues}/{profile.column_count}",
     BAD if profile.columns_with_issues else GOOD),
    (d5, "Missing Values", f"{profile.total_missing_values:,}",
     BAD if profile.total_missing_values else GOOD),
]:
    with col:
        st.markdown(metric_card(label, value, color, small=True), unsafe_allow_html=True)

if profile.columns:
    b1, b2 = st.columns(2)
    worst = next(c for c in profile.columns if c.column == profile.worst_column)
    best = next(c for c in profile.columns if c.column == profile.best_column)
    with b1:
        st.markdown(
            f'<div class="metric-card small" style="text-align:left;padding:0.8rem 1.2rem;">'
            f'<p style="margin-bottom:4px;">🔻 Weakest column</p>'
            f'<h2 style="font-size:1.2rem;color:{score_color(worst.quality_score)}">{worst.column} — {worst.quality_score:.0f}/100</h2>'
            f"</div>",
            unsafe_allow_html=True,
        )
    with b2:
        st.markdown(
            f'<div class="metric-card small" style="text-align:left;padding:0.8rem 1.2rem;">'
            f'<p style="margin-bottom:4px;">🏆 Strongest column</p>'
            f'<h2 style="font-size:1.2rem;color:{score_color(best.quality_score)}">{best.column} — {best.quality_score:.0f}/100</h2>'
            f"</div>",
            unsafe_allow_html=True,
        )

st.write("")

# ------------------------------------------------------------- score chart --
col_names = [c.column for c in profile.columns]
col_scores = [c.quality_score for c in profile.columns]
bar_colors = [score_color(s) for s in col_scores]
hover_text = [
    f"{c.column}<br>Score: {c.quality_score:.0f}<br>Weakest: {c.weakest_dimension}"
    + (f" ({c.weakest_dimension_pct:.0f}%)" if c.weakest_dimension != "None" else "")
    for c in profile.columns
]

fig = go.Figure(
    go.Bar(
        x=col_scores, y=col_names, orientation="h",
        marker_color=bar_colors,
        text=[f"{s:.0f}" for s in col_scores], textposition="outside",
        hovertext=hover_text, hoverinfo="text",
    )
)
fig.update_layout(
    height=max(280, 40 * len(col_names)),
    xaxis=dict(title="Quality score (0-100)", range=[0, 108], gridcolor=BORDER),
    yaxis=dict(autorange="reversed"),
    plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
    font_color=TEXT,
    margin=dict(l=10, r=10, t=10, b=10),
)
st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------ deeper analysis --
st.subheader("Deeper analysis")

a1, a2 = st.columns(2)

with a1:
    st.markdown("**Where is the dataset weakest, dimension by dimension?**")
    dims = ["Completeness", "Validity", "Uniqueness"]
    dim_values = [profile.avg_completeness, profile.avg_validity, profile.avg_uniqueness]
    dim_colors = [DIMENSION_COLORS[d] for d in dims]
    fig_dims = go.Figure(
        go.Bar(x=dims, y=dim_values, marker_color=dim_colors, text=[f"{v:.0f}%" for v in dim_values], textposition="outside")
    )
    fig_dims.update_layout(
        height=320,
        yaxis=dict(title="Average % across all columns", range=[0, 108], gridcolor=BORDER),
        plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
        font_color=TEXT,
        margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig_dims, use_container_width=True)

with a2:
    st.markdown("**Which dimension causes the most column-level issues?**")
    if profile.dimension_issue_counts:
        labels = list(profile.dimension_issue_counts.keys())
        values = list(profile.dimension_issue_counts.values())
        colors = [DIMENSION_COLORS.get(l, TEXT_MUTED) for l in labels]
        fig_pie = go.Figure(go.Pie(labels=labels, values=values, marker_colors=colors, hole=0.5))
        fig_pie.update_layout(
            height=320,
            paper_bgcolor=SURFACE, font_color=TEXT,
            margin=dict(l=10, r=10, t=10, b=10),
            legend=dict(orientation="h", y=-0.1),
        )
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.success("No column has a weak dimension — every column clears the 95% bar on completeness, validity, and (where relevant) uniqueness.")

st.markdown("**Inferred type mix across columns**")
type_counts = pd.Series([c.inferred_type for c in profile.columns]).value_counts()
fig_types = go.Figure(
    go.Bar(
        x=type_counts.values, y=type_counts.index, orientation="h",
        marker_color=ACCENT,
        text=type_counts.values, textposition="outside",
    )
)
fig_types.update_layout(
    height=max(220, 32 * len(type_counts)),
    xaxis=dict(title="Number of columns", gridcolor=BORDER),
    yaxis=dict(autorange="reversed"),
    plot_bgcolor=SURFACE, paper_bgcolor=SURFACE,
    font_color=TEXT,
    margin=dict(l=10, r=10, t=10, b=10),
)
st.plotly_chart(fig_types, use_container_width=True)

# --------------------------------------------------------------- per-column --
st.subheader("Column-by-column breakdown")

problem_columns = sorted(
    [c for c in profile.columns if c.weakest_dimension != "None"], key=lambda c: c.quality_score
)
if problem_columns:
    st.markdown("**Top problem columns**, ranked worst first:")
    st.dataframe(
        pd.DataFrame(
            [
                {
                    "Column": c.column,
                    "Type": c.inferred_type,
                    "Score": c.quality_score,
                    "Weakest dimension": c.weakest_dimension,
                    "Weakest %": c.weakest_dimension_pct,
                    "Notes": c.notes or "—",
                }
                for c in problem_columns
            ]
        ),
        use_container_width=True,
        hide_index=True,
    )

for c in profile.columns:
    header = f"**{c.column}**  ·  score {c.quality_score:.0f}/100"
    with st.expander(header, expanded=c.quality_score < 85):
        st.markdown(dimension_badge(c.weakest_dimension, c.weakest_dimension_pct), unsafe_allow_html=True)
        st.write("")
        top1, top2, top3, top4 = st.columns(4)
        top1.markdown(f"**Inferred type**<br>{type_badge(c.inferred_type)}", unsafe_allow_html=True)
        top2.metric("Completeness", f"{c.completeness_pct}%")
        top3.metric("Validity", f"{c.validity_pct}%")
        top4.metric("Uniqueness", f"{c.uniqueness_pct}%")

        if c.notes:
            st.warning(c.notes)
        if c.sample_invalid:
            st.markdown(f"**Sample problem values** ({len(c.sample_invalid)} shown):")
            st.code("\n".join(c.sample_invalid) or "—")
        if c.issue_count == 0 and not c.notes:
            st.success("No issues detected in this column.")

# ------------------------------------------------------------------ export --
st.divider()
report_rows = [
    {
        "column": c.column, "inferred_type": c.inferred_type, "confidence": c.confidence,
        "completeness_pct": c.completeness_pct, "validity_pct": c.validity_pct,
        "uniqueness_pct": c.uniqueness_pct, "quality_score": c.quality_score,
        "issue_count": c.issue_count, "missing_count": c.missing_count,
        "weakest_dimension": c.weakest_dimension, "weakest_dimension_pct": c.weakest_dimension_pct,
        "notes": c.notes,
    }
    for c in profile.columns
]
report_df = pd.DataFrame(report_rows)


def safe_filename_slug(name: str, sheet: str | None) -> str:
    """'employees.xlsx' + 'Q1 Sales' -> 'employees-Q1_Sales' (filesystem-safe)."""
    base = re.sub(r"\.[^./\\]+$", "", name) or "report"
    slug = f"{base}-{sheet}" if sheet else base
    slug = re.sub(r"[^a-zA-Z0-9._-]+", "_", slug).strip("_")
    return slug or "report"


st.download_button(
    "⬇ Download full report as CSV",
    report_df.to_csv(index=False).encode("utf-8"),
    file_name=f"quality_report_{safe_filename_slug(base_file_name, selected_sheet)}.csv",
    mime="text/csv",
)
