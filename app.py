"""
USPTO Examining Attorney Performance Dashboard
Single-file Streamlit web application supporting multi-workbook Excel uploads,
turnaround compliance tracking, refusal ground analysis, and docket notes.

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

# ---------------------------------------------------------
# Sidebar: File Uploader (Multiple Files Supported)
# ---------------------------------------------------------
st.sidebar.header("📁 Data Source")
uploaded_files = st.sidebar.file_uploader(
    "Upload USPTO Docket Excel (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Upload one or more Excel workbooks matching the USPTO Examining Attorney schema."
)

# Safe Halt if No Files Uploaded
if not uploaded_files:
    st.markdown('<div class="main-title">USPTO Examining Attorney Performance Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Office of Trademark Examination · Performance & Docket Analytics</div>', unsafe_allow_html=True)
    st.info("Please upload one or more Excel files.")
    st.stop()

# ---------------------------------------------------------
# Multi-File Data Ingestion & Concat Loop
# ---------------------------------------------------------
dataframes = []
for file in uploaded_files:
    try:
        temp_df = pd.read_excel(file)
        dataframes.append(temp_df)
    except Exception as e:
        st.sidebar.error(f"Error reading file '{file.name}': {e}")

if not dataframes:
    st.error("No valid tabular data could be read from the uploaded files.")
    st.stop()

# Merge all uploaded workbooks into a single master DataFrame
df = pd.concat(dataframes, ignore_index=True)

# ---------------------------------------------------------
# Data Validation & Processing
# ---------------------------------------------------------
# Verify required schema exists
missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
if missing_cols:
    st.error(
        f"⚠️ **Schema Validation Warning**: The uploaded spreadsheet(s) are missing required columns:\n\n"
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

# Standardize data types
df["Turnaround (Days)"] = pd.to_numeric(df["Turnaround (Days)"], errors="coerce").fillna(0)
df["Applicant Rep Type"] = df["Applicant Rep Type"].fillna("Unknown").astype(str)
df["Action Taken"] = df["Action Taken"].fillna("Unknown").astype(str)
df["Primary Refusal Ground"] = df["Primary Refusal Ground"].fillna("None / Not Specified").astype(str)

# ---------------------------------------------------------
# Sidebar Filters
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🔍 Filter Docket")
st.sidebar.caption(f"Loaded {len(uploaded_files)} workbook(s) · {len(df):,} total marks")

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
    month_order = [
        "January", "February", "March", "April", "May", "June", 
        "July", "August", "September", "October", "November", "December"
    ]
    present_months = [m for m in month_order if m in df["Month_Name"].unique()]
    month_options = ["All Months"] + present_months

selected_month = st.sidebar.selectbox("Action Date Month", options=month_options, index=0)

# Apply dynamic filters
filtered_df = df.copy()

if selected_rep_types:
    filtered_df = filtered_df[filtered_df["Applicant Rep Type"].isin(selected_rep_types)]
else:
    filtered_df = filtered_df.iloc[0:0]

if selected_year != "All Years":
    filtered_df = filtered_df[filtered_df["Year"] == int(selected_year)]

if selected_month != "All Months":
    filtered_df = filtered_df[filtered_df["Month_Name"] == selected_month]

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
st.markdown(
    f'<div class="sub-title">Office of Trademark Examination · Multi-Docket Analytics ({filter_summary})</div>',
    unsafe_allow_html=True
)

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
        help="Total row count of trademark applications in the selected reporting period across all workbooks."
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
    trend_df = filtered_df.sort_values("Action Date").copy()

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
