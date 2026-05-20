import streamlit as st
import pandas as pd

from analytics import (
    calculate_funding_efficiency,
    calculate_kpis,
    calculate_trends,
    clean_data,
    export_excel_dashboard,
    generate_executive_summary,
    generate_visualizations,
    identify_bottom_programs,
    identify_top_programs,
    load_data,
    prepare_display_tables,
    validate_data,
)
from config import ASSET_DIR, ERROR_LOG_FILE, REPORT_FILE, SUMMARY_FILE


st.set_page_config(
    page_title="PANYNJ Port Workforce Analytics",
    page_icon="PA",
    layout="wide",
)


def inject_styles():
    st.markdown(
        """
        <style>
        :root {
            --heading: #244a73;
            --blue: #2f6f9f;
            --line: #d9e2ec;
            --surface: #f6f8fb;
            --text: #1f2933;
        }
        .stApp {
            background: #f7f9fb;
            color: var(--text);
        }
        [data-testid="stSidebar"] {
            background: #ffffff;
            border-right: 1px solid #d9e2ec;
        }
        [data-testid="stSidebar"] * {
            color: #243b53;
        }
        .block-container {
            padding-top: 1.2rem;
            padding-bottom: 2rem;
        }
        .pa-header {
            background: #ffffff;
            color: #243b53;
            padding: 26px 30px;
            border-radius: 8px;
            border: 1px solid #d9e2ec;
            margin-bottom: 18px;
            box-shadow: 0 2px 10px rgba(36, 59, 83, 0.06);
        }
        .pa-eyebrow {
            font-size: 0.78rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: #5f7d95;
            margin-bottom: 8px;
        }
        .pa-title {
            font-size: 2rem;
            font-weight: 760;
            margin: 0;
            line-height: 1.1;
        }
        .pa-subtitle {
            max-width: 1100px;
            color: #486581;
            margin-top: 8px;
            font-size: 1rem;
        }
        [data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-left: 5px solid #2f6f9f;
            padding: 16px 16px 12px;
            border-radius: 8px;
            min-height: 116px;
            box-shadow: 0 2px 8px rgba(11, 37, 69, 0.06);
        }
        [data-testid="stMetricLabel"] p {
            color: #52616f;
            font-size: 0.82rem;
        }
        [data-testid="stMetricValue"] {
            color: #244a73;
            font-weight: 760;
        }
        .section-title {
            color: #244a73;
            font-weight: 760;
            font-size: 1.1rem;
            margin: 14px 0 8px;
        }
        .insight-box {
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-radius: 8px;
            padding: 16px 18px;
            white-space: pre-wrap;
            max-height: 420px;
            overflow-y: auto;
        }
        .status-note {
            background: #eef5ff;
            border: 1px solid #c8dcf5;
            color: #244a73;
            padding: 10px 12px;
            border-radius: 8px;
            font-size: 0.9rem;
        }
        .chart-note {
            background: #ffffff;
            border-left: 4px solid #2f6f9f;
            color: #334e68;
            padding: 10px 12px;
            margin: -6px 0 14px;
            font-size: 0.88rem;
            line-height: 1.35;
        }
        .context-strip {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 10px;
            margin: 8px 0 16px;
        }
        .context-item {
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-radius: 8px;
            padding: 12px 14px;
            color: #334e68;
            font-size: 0.9rem;
        }
        .context-item strong {
            color: #244a73;
            display: block;
            margin-bottom: 4px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def filter_data(df):
    st.sidebar.image(str(ASSET_DIR / "panynj_logo_placeholder.png"), use_container_width=True)
    st.sidebar.markdown("### Port Workforce Filters")

    quarters = sorted(df["quarter"].unique(), key=lambda q: int(q.split()[1]) * 10 + int(q[1]))
    programs = sorted(df["program_name"].unique())
    counties = sorted(df["county"].unique())
    specializations = sorted(df["logistics_specialization"].unique())

    selected_quarters = st.sidebar.multiselect("Quarter", quarters, default=quarters)
    selected_programs = st.sidebar.multiselect("Workforce Program", programs, default=programs)
    selected_counties = st.sidebar.multiselect("County", counties, default=counties)
    selected_specializations = st.sidebar.multiselect(
        "Logistics Specialization",
        specializations,
        default=specializations,
    )
    min_participants = st.sidebar.slider(
        "Minimum Participants",
        0,
        int(max(df["participants"].max(), 1)),
        0,
    )
    min_readiness = st.sidebar.slider("Minimum Workforce Readiness Score", 0, 100, 0)

    filtered = df[
        df["quarter"].isin(selected_quarters)
        & df["program_name"].isin(selected_programs)
        & df["county"].isin(selected_counties)
        & df["logistics_specialization"].isin(selected_specializations)
        & (df["participants"] >= min_participants)
        & (df["workforce_readiness_score"] >= min_readiness)
    ].copy()
    return filtered


def render_header():
    st.markdown(
        """
        <div class="pa-header">
            <div class="pa-eyebrow">Port Authority of New York & New Jersey | Port Department</div>
            <h1 class="pa-title">Enterprise Workforce Development Analytics</h1>
            <div class="pa-subtitle">
                Internal Port Department reporting for companywide maritime, supply chain, logistics,
                freight, terminal, and regional Transportation, Logistics, and Distribution workforce initiatives.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_context_strip():
    st.markdown(
        """
        <div class="context-strip">
            <div class="context-item">
                <strong>What this monitors</strong>
                Port-relevant workforce readiness, placements, credentials, funding, and employer partnerships.
            </div>
            <div class="context-item">
                <strong>Who it supports</strong>
                Port Department leadership, Port Policy & Planning, Business Solutions, and reporting teams.
            </div>
            <div class="context-item">
                <strong>How to read it</strong>
                Higher readiness and placement rates signal stronger alignment with maritime, freight, and TLD labor needs.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpis(kpis):
    cols = st.columns(6)
    cols[0].metric("Workforce Participants", f"{kpis.get('total_participants', 0):,}")
    cols[1].metric("Completion Rate", f"{kpis.get('avg_completion_rate', 0):.1f}%")
    cols[2].metric("TLD Placement Rate", f"{kpis.get('avg_job_placement_rate', 0):.1f}%")
    cols[3].metric("Funding Allocation", f"${kpis.get('total_funding', 0):,.0f}")
    cols[4].metric("Active Initiatives", f"{kpis.get('active_programs', 0):,}")
    cols[5].metric("Maritime Placement Rate", f"{kpis.get('maritime_placement_rate', 0):.1f}%")


def chart_note(text):
    st.markdown(f'<div class="chart-note">{text}</div>', unsafe_allow_html=True)


def download_buttons(filtered_df, report_path, summary_text):
    col1, col2, col3 = st.columns(3)
    if report_path and report_path.exists():
        col1.download_button(
            "Download Excel Report",
            data=report_path.read_bytes(),
            file_name="panynj_workforce_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    col2.download_button(
        "Download Executive Summary",
        data=summary_text,
        file_name="executive_summary.txt",
        mime="text/plain",
    )
    csv_data = filtered_df.drop(columns=["quarter_sort"], errors="ignore").to_csv(index=False).encode("utf-8")
    col3.download_button(
        "Download Cleaned Workforce Dataset",
        data=csv_data,
        file_name="cleaned_panynj_workforce_dataset.csv",
        mime="text/csv",
    )


def main():
    inject_styles()
    render_header()
    render_context_strip()

    raw_df, load_errors = load_data()
    if load_errors:
        for error in load_errors:
            st.error(error)
        st.stop()

    valid_df, validation_errors, validation_warnings = validate_data(raw_df)
    if validation_errors:
        for error in validation_errors:
            st.error(error)
        st.info(f"See error log: {ERROR_LOG_FILE}")
        st.stop()
    for warning in validation_warnings:
        st.warning(warning)

    df = clean_data(valid_df)
    filtered_df = filter_data(df)
    if filtered_df.empty:
        st.warning("No records match the selected workforce analytics filters.")
        st.stop()

    kpis = calculate_kpis(filtered_df)
    render_kpis(kpis)

    st.markdown('<div class="section-title">Executive Operating View</div>', unsafe_allow_html=True)
    figures = generate_visualizations(filtered_df)
    left, right = st.columns(2)
    left.plotly_chart(figures["participation_trends"], use_container_width=True)
    with left:
        chart_note("Shows whether the Port Department's workforce pipeline is expanding, stable, or contracting by reporting quarter.")
    right.plotly_chart(figures["placement_trends"], use_container_width=True)
    with right:
        chart_note("Tracks whether programs are converting training participation into transportation, logistics, and distribution employment outcomes.")

    left, right = st.columns(2)
    left.plotly_chart(figures["top_programs"], use_container_width=True)
    with left:
        chart_note("Use these initiatives as models for scalable employer partnerships, credential alignment, and port-related placement practices.")
    right.plotly_chart(figures["bottom_programs"], use_container_width=True)
    with right:
        chart_note("These initiatives may need leadership review, stronger employer engagement, or redesigned training-to-placement pathways.")

    st.markdown('<div class="section-title">Funding, Sector, and Readiness Analysis</div>', unsafe_allow_html=True)
    left, right = st.columns(2)
    left.plotly_chart(figures["funding_vs_placement"], use_container_width=True)
    with left:
        chart_note("Compares investment levels with placement results; large circles with low placement rates are priority review points.")
    right.plotly_chart(figures["specialization_distribution"], use_container_width=True)
    with right:
        chart_note("Shows whether workforce coverage is balanced across maritime logistics, port operations, freight, rail, trucking, warehousing, and supply chain needs.")

    st.plotly_chart(figures["readiness_heatmap"], use_container_width=True)
    chart_note("Darker cells indicate stronger readiness. This view helps spot which initiatives are improving or losing strength across quarters.")

    left, right = st.columns(2)
    left.plotly_chart(figures["completion_vs_placement"], use_container_width=True)
    with left:
        chart_note("Separates completion from job placement so leadership can see whether training success is translating into employment outcomes.")
    right.plotly_chart(figures["county_participation"], use_container_width=True)
    with right:
        chart_note("Summarizes regional reach for workforce programs serving the Port of New York and New Jersey labor market.")

    trends = calculate_trends(filtered_df)
    top_programs = identify_top_programs(filtered_df)
    bottom_programs = identify_bottom_programs(filtered_df)
    funding = calculate_funding_efficiency(filtered_df)
    display_tables = prepare_display_tables(trends, top_programs, bottom_programs, funding)
    summary_text = generate_executive_summary(filtered_df, trends, top_programs, bottom_programs)
    report_path = export_excel_dashboard(filtered_df)

    st.markdown('<div class="section-title">Leadership Briefing Summary</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="insight-box">{summary_text}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Program Performance Tables</div>', unsafe_allow_html=True)
    tab1, tab2, tab3, tab4 = st.tabs(["Quarter Signals", "Strongest Initiatives", "Review Needed", "Funding Efficiency"])
    tab1.dataframe(display_tables["trends"], use_container_width=True, hide_index=True)
    tab2.dataframe(display_tables["top_programs"], use_container_width=True, hide_index=True)
    tab3.dataframe(display_tables["bottom_programs"], use_container_width=True, hide_index=True)
    tab4.dataframe(display_tables["funding"], use_container_width=True, hide_index=True)

    st.markdown('<div class="section-title">Exports</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="status-note">Reports are generated locally for Port Department workforce planning, public-sector reporting, and executive review.</div>',
        unsafe_allow_html=True,
    )
    download_buttons(filtered_df, report_path, summary_text)


if __name__ == "__main__":
    main()
