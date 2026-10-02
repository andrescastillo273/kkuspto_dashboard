"""
USPTO Examining Attorney Performance Dashboard
Single-file Streamlit web application with persistent local file manager and
robust state-isolated file saving and deletion logic.

Requirements:
- streamlit
- pandas
- plotly
- openpyxl
Run with: streamlit run app.py
"""

import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ---------------------------------------------------------
# Directory Management
# ---------------------------------------------------------
REPORTS_DIR = "saved_reports"
if not os.path.exists(REPORTS_DIR):
    os.makedirs(REPORTS_DIR)

# ---------------------------------------------------------
# Page Configuration & Custom Styling
# ---------------------------------------------------------
st.set_page_config(
    page_title="USPTO Examining Attorney Performance Dashboard",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

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
    div[data-testid="stMetric"] {
        background-color: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0F172A;
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
# Uploader Key Reset & Save Logic (Prevents Auto-Resaving)
# ---------------------------------------------------------
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0

st.sidebar.header("📁 Upload & Save Reports")

# File uploader tied to dynamic session key
uploaded_files = st.sidebar.file_uploader(
    "Upload Excel Worksheets (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}",
    help="Select Excel worksheets to upload."
)

# Explicit Save Button: Writes files, increments key to wipe widget buffer, and reruns
if uploaded_files:
    if st.sidebar.button("💾 Save Uploaded Files", type="primary", use_container_width=True):
        saved_count = 0
        for uploaded_file in uploaded_files:
            destination_path = os.path.join(REPORTS_DIR, uploaded_file.name)
            with open(destination_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            saved_count += 1
        
        # Increment uploader key to instantly reset the uploader widget on rerun
        st.session_state["uploader_key"] += 1
        st.session_state["last_action_msg"] = f"Permanently saved {saved_count} file(s) to '{REPORTS_DIR}/'."
        st.rerun()

if "last_action_msg" in st.session_state:
    st.sidebar.success(st.session_state.pop("last_action_msg"))

# ---------------------------------------------------------
# Sidebar File Manager: Read, Select, and Permanently Delete
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🗂️ Saved Reports Manager")

# List files currently on disk in saved_reports
saved_files = sorted([f for f in os.listdir(REPORTS_DIR) if f.endswith(".xlsx")])

if saved_files:
    st.sidebar.caption(f"Found {len(saved_files)} saved report(s) in `{REPORTS_DIR}/`:")
    
    # Multiselect for active analysis (defaults to all currently saved files)
    selected_files = st.sidebar.multiselect(
        "Select files to analyze in dashboard:",
        options=saved_files,
        default=saved_files,
        help="Check the files to include in master analysis."
    )

    # Permanent Deletion Button
    if st.sidebar.button("🗑️ Delete Selected", type="secondary", use_container_width=True):
        if selected_files:
            deleted_count = 0
            for file_name in selected_files:
                file_path = os.path.join(REPORTS_DIR, file_name)
                if os.path.exists(file_path):
                    os.remove(file_path)
                    deleted_count += 1
            st.session_state["last_action_msg"] = f"Permanently deleted {deleted_count} file(s) from '{REPORTS_DIR}/'."
            st.rerun()
        else:
            st.sidebar.warning("No files selected for deletion.")
else:
    selected_files = []
    st.sidebar.info(f"No saved Excel reports found in '{REPORTS_DIR}/'.")

# ---------------------------------------------------------
# Failsafe: Halt if No Files are Selected or Saved
# ---------------------------------------------------------
if not selected_files:
    st.markdown('<div class="main-title">USPTO Examining Attorney Performance Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Office of Trademark Examination · Performance & Docket Analytics</div>', unsafe_allow_html=True)
    st.info("Please upload or select an Excel report from the sidebar to view metrics.")
    st.stop()

# ---------------------------------------------------------
# Master Data Processing: Read & Concatenate Active Files
# ---------------------------------------------------------
dataframes = []
for file_name in selected_files:
    file_path = os.path.join(REPORTS_DIR, file_name)
    if os.path.exists(file_path):
        try:
            temp_df = pd.read_excel(file_path)
            dataframes.append(temp_df)
        except Exception as e:
            st.sidebar.error(f"Error reading '{file_name}': {e}")

if not dataframes:
    st.error("No valid data could be loaded from the selected reports.")
    st.stop()

# Master DataFrame merged from current disk state
df = pd.concat(dataframes, ignore_index=True)

# ---------------------------------------------------------
# Data Schema Validation & Datetime Conversions
# ---------------------------------------------------------
missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
if missing_cols:
    st.error(
        f"⚠️ **Schema Validation Warning**: The selected spreadsheet(s) are missing required columns:\n\n"
        f"Missing: `{', '.join(missing_cols)}`\n\n"
        f"Required columns are:\n" + "\n".join([f"- {col}" for col in REQUIRED_COLUMNS])
    )
    st.stop()

# Explicitly convert Action Date and Date Assigned to datetime
try:
    df["Action Date"] = pd.to_datetime(df["Action Date"])
    df["Date Assigned"] = pd.to_datetime(df["Date Assigned"])
except Exception as e:
    st.error(f"Error converting Action Date or Date Assigned to datetime: {e}")
    st.stop()

# Clean and standardize types
df["Turnaround (Days)"] = pd.to_numeric(df["Turnaround (Days)"], errors="coerce").fillna(0)
df["Applicant Rep Type"] = df["Applicant Rep Type"].fillna("Unknown").astype(str)
df["Action Taken"] = df["Action Taken"].fillna("Unknown").astype(str)
df["Primary Refusal Ground"] = df["Primary Refusal Ground"].fillna("None / Not Specified").astype(str)

# ---------------------------------------------------------
# Interactive Sidebar Filters
# ---------------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.header("🔍 Interactive Filters")

# 1. Applicant Rep Type Dropdown (Pro Se vs. Represented)
rep_types = ["All"] + sorted(df["Applicant Rep Type"].unique().tolist())
selected_rep_type = st.sidebar.selectbox(
    "Applicant Rep Type",
    options=rep_types,
    index=0,
    help="Filter by Pro Se vs. Represented applicants."
)

# Extract Year and Month from Action Date
df["Year"] = df["Action Date"].dt.year
df["Month_Num"] = df["Action Date"].dt.month
df["Month_Name"] = df["Action Date"].dt.strftime("%B")

# 2. Dynamic Year select box
available_years = sorted(df["Year"].dropna().unique().tolist(), reverse=True)
year_options = ["All Years"] + [str(y) for y in available_years]
selected_year = st.sidebar.selectbox("Action Date: Year", options=year_options, index=0)

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

selected_month = st.sidebar.selectbox("Action Date: Month", options=month_options, index=0)

# Apply dynamic filters
filtered_df = df.copy()

if selected_rep_type != "All":
    filtered_df = filtered_df[filtered_df["Applicant Rep Type"] == selected_rep_type]

if selected_year != "All Years":
    filtered_df = filtered_df[filtered_df["Year"] == int(selected_year)]

if selected_month != "All Months":
    filtered_df = filtered_df[filtered_df["Month_Name"] == selected_month]

# ---------------------------------------------------------
# Main Dashboard Header
# ---------------------------------------------------------
st.markdown('<div class="main-title">USPTO Examining Attorney Performance Dashboard</div>', unsafe_allow_html=True)

filter_desc = []
if selected_rep_type != "All":
    filter_desc.append(f"Rep: {selected_rep_type}")
if selected_year != "All Years":
    filter_desc.append(f"Year: {selected_year}")
if selected_month != "All Months":
    filter_desc.append(f"Month: {selected_month}")

filter_summary = " · ".join(filter_desc) if filter_desc else "All Records"
st.markdown(
    f'<div class="sub-title">Trademark Examining Operation · Active Reports: {len(selected_files)} file(s) ({filter_summary})</div>',
    unsafe_allow_html=True
)

# ---------------------------------------------------------
# KPI Metrics
# ---------------------------------------------------------
total_marks = len(filtered_df)
avg_turnaround = filtered_df["Turnaround (Days)"].mean() if total_marks > 0 else 0.0

# Approval Rate: Action Taken == 'Approved for Pub'
approved_count = len(filtered_df[filtered_df["Action Taken"] == "Approved for Pub"])
approval_rate = (approved_count / total_marks * 100) if total_marks > 0 else 0.0

kpi_col1, kpi_col2, kpi_col3 = st.columns(3)

with kpi_col1:
    st.metric(
        label="Total Marks Processed",
        value=f"{total_marks:,}",
        help="Total row count of filtered trademark applications."
    )

with kpi_col2:
    st.metric(
        label="Average Turnaround Time in Days",
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

    # 2. Refusal Analysis: Pie / Donut Chart
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
            title="Primary Refusal Ground Distribution",
            hole=0.45,
            color_discrete_sequence=px.colors.qualitative.Safe
        )
        fig_refusal.update_traces(
            textinfo="percent+label",
            textposition="auto",
            insidetextorientation="radial"
        )
        # Explicit margins requested so slices and labels fit without cut off
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

    # 3. Efficiency Trend: Chronological Line Chart
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

    # Reference Average Line
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
st.caption("Reference examining attorney Quality Review / Notes and individual application docket history.")

display_df = filtered_df[REQUIRED_COLUMNS].copy()
display_df["Action Date"] = display_df["Action Date"].dt.strftime("%Y-%m-%d")
display_df["Date Assigned"] = display_df["Date Assigned"].dt.strftime("%Y-%m-%d")

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True
)
