"""
USPTO Examining Attorney Performance Dashboard
Single-file Streamlit web application for tracking docket metrics,
turnaround times, refusal grounds, and quality review notes.

Requirements:
- streamlit
- pandas
- plotly
- openpyxl
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="USPTO Examining Attorney Performance Dashboard",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    .main-title {
        font-size: 2rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.25rem;
    }
    .sub-title {
        font-size: 0.95rem;
        color: #64748B;
        margin-bottom: 1.5rem;
    }
    .stMetric {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px;
    }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Required Data Schema
# ---------------------------------------------------------
REQUIRED_COLUMNS = [
    "Serial Number",
    "Mark Text",
    "Int. Class(es)",
    "Date Assigned",
    "Action Date",
    "Turnaround (Days)",
    "Action Taken",
    "Primary Refusal Ground",
    "Applicant Rep Type",
    "Quality Review / Notes"
]

def generate_sample_data() -> pd.DataFrame:
    """Generates realistic USPTO Examining Attorney docket data for instant demonstration."""
    sample_records = [
        {
            "Serial Number": "97/842,101",
            "Mark Text": "QUANTUM CLOUD",
            "Int. Class(es)": "009, 042",
            "Date Assigned": "2026-08-01",
            "Action Date": "2026-08-15",
            "Turnaround (Days)": 14,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Clear search; identification of services in Class 42 conforms with ID Manual."
        },
        {
            "Serial Number": "97/842,102",
            "Mark Text": "APEX BREADWORKS",
            "Int. Class(es)": "030",
            "Date Assigned": "2026-08-05",
            "Action Date": "2026-08-22",
            "Turnaround (Days)": 17,
            "Action Taken": "Non-Final Office Action",
            "Primary Refusal Ground": "Section 2(d) Likelihood of Confusion",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Cited prior pending Reg. 5,892,104 APEX BAKERY for identical bread goods. Class 30 search complete."
        },
        {
            "Serial Number": "97/842,103",
            "Mark Text": "RAPID HEAL GEL",
            "Int. Class(es)": "005",
            "Date Assigned": "2026-08-10",
            "Action Date": "2026-08-28",
            "Turnaround (Days)": 18,
            "Action Taken": "Non-Final Office Action",
            "Primary Refusal Ground": "Section 2(e)(1) Merely Descriptive",
            "Applicant Rep Type": "Pro Se",
            "Quality Review / Notes": "Applicant pro se. Mark immediately describes function and composition of topical analgesic gel."
        },
        {
            "Serial Number": "97/842,104",
            "Mark Text": "CYBERGUARD SECURITY",
            "Int. Class(es)": "042",
            "Date Assigned": "2026-08-12",
            "Action Date": "2026-08-30",
            "Turnaround (Days)": 18,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Disclaimer of 'SECURITY' accepted via Examiner's Amendment prior to pub approval."
        },
        {
            "Serial Number": "97/842,105",
            "Mark Text": "VITAL BLENDS",
            "Int. Class(es)": "032",
            "Date Assigned": "2026-09-01",
            "Action Date": "2026-09-12",
            "Turnaround (Days)": 11,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Passed supervisory review; specimen shows mark in commerce on fruit juice cartons."
        },
        {
            "Serial Number": "97/842,106",
            "Mark Text": "ECOCLEAN SOLUTIONS",
            "Int. Class(es)": "003, 005",
            "Date Assigned": "2026-09-02",
            "Action Date": "2026-09-18",
            "Turnaround (Days)": 16,
            "Action Taken": "Non-Final Office Action",
            "Primary Refusal Ground": "Identification of Goods/Services",
            "Applicant Rep Type": "Pro Se",
            "Quality Review / Notes": "Multi-class ID indefinite. Provided suggested wording for non-toxic cleaning preparations."
        },
        {
            "Serial Number": "97/842,107",
            "Mark Text": "HYPERDRIVE AUTOMOTIVE",
            "Int. Class(es)": "012",
            "Date Assigned": "2026-09-04",
            "Action Date": "2026-09-19",
            "Turnaround (Days)": 15,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Search reveals no conflicting marks; drawing matches specimen."
        },
        {
            "Serial Number": "97/842,108",
            "Mark Text": "ZENITH CAPITAL PARTNERS",
            "Int. Class(es)": "036",
            "Date Assigned": "2026-09-05",
            "Action Date": "2026-09-24",
            "Turnaround (Days)": 19,
            "Action Taken": "Final Office Action",
            "Primary Refusal Ground": "Section 2(d) Likelihood of Confusion",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Finalized 2(d) refusal against ZENITH FINANCIAL GROUP. Applicant arguments unpersuasive."
        },
        {
            "Serial Number": "97/842,109",
            "Mark Text": "SOLARIS ENERGY LABS",
            "Int. Class(es)": "040, 042",
            "Date Assigned": "2026-09-08",
            "Action Date": "2026-09-26",
            "Turnaround (Days)": 18,
            "Action Taken": "Suspension",
            "Primary Refusal Ground": "Section 2(d) Likelihood of Confusion",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Suspended pending disposition of prior-filed Application Ser. 97/654,321."
        },
        {
            "Serial Number": "97/842,110",
            "Mark Text": "NEXUS FLOW",
            "Int. Class(es)": "009",
            "Date Assigned": "2026-09-10",
            "Action Date": "2026-09-27",
            "Turnaround (Days)": 17,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Pro Se",
            "Quality Review / Notes": "Pro se applicant filed acceptable substitute specimen in response to phone inquiry."
        },
        {
            "Serial Number": "97/842,111",
            "Mark Text": "PURE OMNI HYDRATION",
            "Int. Class(es)": "032",
            "Date Assigned": "2026-09-14",
            "Action Date": "2026-09-28",
            "Turnaround (Days)": 14,
            "Action Taken": "Non-Final Office Action",
            "Primary Refusal Ground": "Specimen Refusal",
            "Applicant Rep Type": "Pro Se",
            "Quality Review / Notes": "Specimen is a digital mock-up; issued requirement for verified substitute specimen."
        },
        {
            "Serial Number": "97/842,112",
            "Mark Text": "TRUE NORTH LOGISTICS",
            "Int. Class(es)": "039",
            "Date Assigned": "2026-09-15",
            "Action Date": "2026-09-29",
            "Turnaround (Days)": 14,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Approved without objections. Turnaround within target threshold."
        },
        {
            "Serial Number": "97/842,113",
            "Mark Text": "TITAN ROOFING & SIDING",
            "Int. Class(es)": "037",
            "Date Assigned": "2026-10-01",
            "Action Date": "2026-10-12",
            "Turnaround (Days)": 11,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Clean examination; standard character claim verified."
        },
        {
            "Serial Number": "97/842,114",
            "Mark Text": "GOLDEN HARVEST CRISPS",
            "Int. Class(es)": "029, 030",
            "Date Assigned": "2026-10-02",
            "Action Date": "2026-10-16",
            "Turnaround (Days)": 14,
            "Action Taken": "Non-Final Office Action",
            "Primary Refusal Ground": "Section 2(d) Likelihood of Confusion",
            "Applicant Rep Type": "Pro Se",
            "Quality Review / Notes": "Pro se applicant. Cited HARVEST GOLD snacks Reg. 4,112,900. ID clarification requested."
        },
        {
            "Serial Number": "97/842,115",
            "Mark Text": "BLUEWAVE ANALYTICS",
            "Int. Class(es)": "042",
            "Date Assigned": "2026-10-05",
            "Action Date": "2026-10-18",
            "Turnaround (Days)": 13,
            "Action Taken": "Approved for Pub",
            "Primary Refusal Ground": "None (Approved)",
            "Applicant Rep Type": "Represented",
            "Quality Review / Notes": "Fast turnaround. Quality check confirmed no 2(d) conflicts in tech class 42."
        }
    ]
    return pd.DataFrame(sample_records)


# ---------------------------------------------------------
# Sidebar: File Uploader & Controls
# ---------------------------------------------------------
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload USPTO Docket Excel (.xlsx)",
    type=["xlsx"],
    help="Upload an Excel spreadsheet matching the USPTO Examining Attorney schema."
)

use_sample_data = False
if uploaded_file is None:
    st.sidebar.info("💡 No file uploaded yet. You can use the built-in sample docket for demonstration.")
    if st.sidebar.button("Load USPTO Sample Docket", use_container_width=True):
        st.session_state["use_sample"] = True

if st.session_state.get("use_sample", False) and uploaded_file is None:
    use_sample_data = True

# Load and validate data
df = None
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.sidebar.error(f"Error reading Excel file: {e}")
elif use_sample_data:
    df = generate_sample_data()

# ---------------------------------------------------------
# Data Validation & Processing
# ---------------------------------------------------------
if df is not None:
    # Verify required schema
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    
    if missing_cols:
        st.error(
            f"⚠️ **Schema Validation Warning**: The uploaded spreadsheet is missing required columns:\n\n"
            f"Missing: `{', '.join(missing_cols)}`\n\n"
            f"Required columns are:\n" + "\n".join([f"- {c}" for c in REQUIRED_COLUMNS])
        )
        st.stop()

    # Datetime conversions
    try:
        df["Action Date"] = pd.to_datetime(df["Action Date"])
        df["Date Assigned"] = pd.to_datetime(df["Date Assigned"])
    except Exception as e:
        st.error(f"Error converting Action Date or Date Assigned to datetime: {e}")
        st.stop()

    # Numeric conversion for turnaround
    df["Turnaround (Days)"] = pd.to_numeric(df["Turnaround (Days)"], errors="coerce").fillna(0)

    # Clean text columns
    df["Applicant Rep Type"] = df["Applicant Rep Type"].fillna("Unknown").astype(str)
    df["Action Taken"] = df["Action Taken"].fillna("Unknown").astype(str)
    df["Primary Refusal Ground"] = df["Primary Refusal Ground"].fillna("None / Not Specified").astype(str)

    # ---------------------------------------------------------
    # Sidebar Filters
    # ---------------------------------------------------------
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filter Docket")

    # 1. Applicant Rep Type Multiselect
    available_rep_types = sorted(df["Applicant Rep Type"].unique().tolist())
    selected_rep_types = st.sidebar.multiselect(
        "Applicant Rep Type",
        options=available_rep_types,
        default=available_rep_types,
        help="Filter marks by Pro Se vs. Attorney Represented applicants."
    )

    # Extract Year and Month from Action Date
    df["Year"] = df["Action Date"].dt.year
    df["Month_Num"] = df["Action Date"].dt.month
    df["Month_Name"] = df["Action Date"].dt.strftime("%B")

    # 2. Dynamic Year select box
    available_years = sorted(df["Year"].dropna().unique().tolist(), reverse=True)
    year_options = ["All Years"] + [str(y) for y in available_years]
    selected_year = st.sidebar.selectbox("Action Date Year", options=year_options, index=0)

    # 3. Dynamic Month select box based on selected Year
    if selected_year != "All Years":
        months_in_year = (
            df[df["Year"] == int(selected_year)][["Month_Num", "Month_Name"]]
            .drop_duplicates()
            .sort_values("Month_Num")["Month_Name"]
            .tolist()
        )
        month_options = ["All Months"] + months_in_year
    else:
        month_order = ["January", "February", "March", "April", "May", "June", 
                       "July", "August", "September", "October", "November", "December"]
        present_months = [m for m in month_order if m in df["Month_Name"].unique()]
        month_options = ["All Months"] + present_months

    selected_month = st.sidebar.selectbox("Action Date Month", options=month_options, index=0)

    # Apply filters dynamically
    filtered_df = df.copy()

    if selected_rep_types:
        filtered_df = filtered_df[filtered_df["Applicant Rep Type"].isin(selected_rep_types)]
    else:
        filtered_df = filtered_df.iloc[0:0]

    if selected_year != "All Years":
        filtered_df = filtered_df[filtered_df["Year"] == int(selected_year)]

    if selected_month != "All Months":
        filtered_df = filtered_df[filtered_df["Month_Name"] == selected_month]

    # Download Template Button in Sidebar
    st.sidebar.markdown("---")
    sample_buffer = io.BytesIO()
    with pd.ExcelWriter(sample_buffer, engine="openpyxl") as writer:
        generate_sample_data().to_excel(writer, index=False, sheet_name="Docket")
    st.sidebar.download_button(
        label="📥 Download Template Excel",
        data=sample_buffer.getvalue(),
        file_name="uspto_examining_attorney_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------------------------------------------------------
    # Main Dashboard
    # ---------------------------------------------------------
    st.markdown('<div class="main-title">USPTO Examining Attorney Performance Dashboard</div>', unsafe_allow_html=True)
    active_filters_text = []
    if selected_year != "All Years":
        active_filters_text.append(f"Year: {selected_year}")
    if selected_month != "All Months":
        active_filters_text.append(f"Month: {selected_month}")
    if selected_rep_types:
        active_filters_text.append(f"Rep Type: {', '.join(selected_rep_types)}")

    filter_summary = " | ".join(active_filters_text) if active_filters_text else "Showing all records"
    st.markdown(f'<div class="sub-title">Office of Trademark Examination · Performance & Docket Analytics ({filter_summary})</div>', unsafe_allow_html=True)

    # ---------------------------------------------------------
    # KPI Metrics
    # ---------------------------------------------------------
    total_marks = len(filtered_df)
    avg_turnaround = filtered_df["Turnaround (Days)"].mean() if total_marks > 0 else 0.0
    
    # Approval rate: Action Taken == 'Approved for Pub'
    approved_marks = filtered_df[filtered_df["Action Taken"] == "Approved for Pub"]
    approval_rate = (len(approved_marks) / total_marks * 100) if total_marks > 0 else 0.0

    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)

    with kpi_col1:
        st.metric(
            label="Total Marks Processed",
            value=f"{total_marks:,}",
            help="Total row count of trademark applications in the selected reporting period."
        )

    with kpi_col2:
        st.metric(
            label="Average Turnaround Time (Days)",
            value=f"{avg_turnaround:.1f} days",
            help="Mean turnaround days from Date Assigned to Action Date."
        )

    with kpi_col3:
        st.metric(
            label="Approval Rate",
            value=f"{approval_rate:.1f}%",
            help="Percentage of filtered applications where Action Taken is 'Approved for Pub'."
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # Visualizations
    # ---------------------------------------------------------
    if total_marks > 0:
        viz_col1, viz_col2 = st.columns(2)

        # 1. Action Breakdown Bar Chart
        with viz_col1:
            st.subheader("📊 Action Breakdown")
            action_counts = (
                filtered_df["Action Taken"]
                .value_counts()
                .reset_index()
            )
            action_counts.columns = ["Action Taken", "Count"]

            fig_action = px.bar(
                action_counts,
                x="Action Taken",
                y="Count",
                color="Action Taken",
                text="Count",
                title="Frequency by Action Taken",
                color_discrete_sequence=px.colors.qualitative.Prism
            )
            fig_action.update_traces(textposition="outside")
            fig_action.update_layout(
                showlegend=False,
                xaxis_title="",
                yaxis_title="Count of Marks",
                height=380,
                margin=dict(l=20, r=20, t=40, b=40)
            )
            st.plotly_chart(fig_action, use_container_width=True)

        # 2. Refusal Analysis Donut / Pie Chart
        with viz_col2:
            st.subheader("⚖️ Refusal Analysis")
            refusal_counts = (
                filtered_df["Primary Refusal Ground"]
                .value_counts()
                .reset_index()
            )
            refusal_counts.columns = ["Primary Refusal Ground", "Count"]

            fig_refusal = px.pie(
                refusal_counts,
                names="Primary Refusal Ground",
                values="Count",
                title="Distribution of Primary Refusal Grounds",
                hole=0.45,
                color_discrete_sequence=px.colors.qualitative.Safe
            )
            fig_refusal.update_traces(
                textinfo="percent+label",
                textposition="auto",
                insidetextorientation="radial"
            )
            fig_refusal.update_layout(
                height=450,
                margin=dict(t=50, b=50, l=50, r=50),
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.12,
                    xanchor="center",
                    x=0.5
                )
            )
            st.plotly_chart(fig_refusal, use_container_width=True)

        # 3. Efficiency Trend: Chronological Turnaround Time
        st.subheader("📈 Efficiency Trend")
        # Sort chronologically by Action Date
        trend_df = (
            filtered_df.sort_values("Action Date")
            .copy()
        )
        trend_df["Action Date Str"] = trend_df["Action Date"].dt.strftime("%Y-%m-%d")

        fig_trend = px.line(
            trend_df,
            x="Action Date",
            y="Turnaround (Days)",
            markers=True,
            hover_data=["Serial Number", "Mark Text", "Action Taken"],
            title="Chronological Turnaround Time (Days) by Action Date",
            color_discrete_sequence=["#2563EB"]
        )

        # Add reference average line
        fig_trend.add_hline(
            y=avg_turnaround,
            line_dash="dash",
            line_color="#DC2626",
            annotation_text=f"Avg: {avg_turnaround:.1f} d",
            annotation_position="bottom right"
        )

        fig_trend.update_layout(
            xaxis_title="Action Date",
            yaxis_title="Turnaround (Days)",
            height=360,
            margin=dict(l=20, r=20, t=40, b=40)
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    else:
        st.warning("No records match the selected filters. Please adjust the sidebar filters.")

    # ---------------------------------------------------------
    # Raw Data Preview
    # ---------------------------------------------------------
    st.subheader("📋 Filtered Docket Raw Data")
    st.caption("Review examined trademark applications, procedural history, and examining attorney Quality Review / Notes.")

    display_df = filtered_df[REQUIRED_COLUMNS].copy()
    display_df["Action Date"] = display_df["Action Date"].dt.strftime("%Y-%m-%d")
    display_df["Date Assigned"] = display_df["Date Assigned"].dt.strftime("%Y-%m-%d")

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True
    )

else:
    # Initial Welcome State
    st.markdown('<div class="main-title">USPTO Examining Attorney Performance Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Office of Trademark Examination · Performance & Docket Analytics</div>', unsafe_allow_html=True)

    st.info("👋 Welcome! Please upload a USPTO Docket Excel (.xlsx) file in the sidebar, or click below to load sample examining attorney docket records.")
    
    col_a, col_b = st.columns([1, 2])
    with col_a:
        if st.button("🚀 Load Sample USPTO Docket", use_container_width=True):
            st.session_state["use_sample"] = True
            st.rerun()

    st.markdown("### Expected Spreadsheet Schema")
    st.markdown("Your uploaded Excel file must contain the following columns:")
    cols_display = pd.DataFrame({
        "Required Column": REQUIRED_COLUMNS,
        "Description": [
            "USPTO 8-digit Serial Number (e.g., 97/842,101)",
            "Word mark text or mark description",
            "International Class(es) (e.g., 009, 042)",
            "Date trademark application was assigned to attorney docket (YYYY-MM-DD)",
            "Date examination action was signed/issued (YYYY-MM-DD)",
            "Elapsed working/calendar days from Date Assigned to Action Date",
            "Procedural action issued (e.g., Approved for Pub, Non-Final Office Action, Final Office Action, Suspension)",
            "Primary legal ground for refusal (e.g., Section 2(d), Section 2(e)(1), Specimen Refusal, None)",
            "Representation status: 'Represented' or 'Pro Se'",
            "Supervisory Quality Review notes, legal rationale, or examiner commentary"
        ]
    })
    st.table(cols_display)
