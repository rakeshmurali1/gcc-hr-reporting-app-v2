
import streamlit as st
import pandas as pd
import datetime
from openpyxl import load_workbook
from io import BytesIO
import altair as alt

TEMPLATE_PATH = "updated_template.xlsx"

def load_template():
    return load_workbook(TEMPLATE_PATH)

def save_to_bytes(wb):
    excel_io = BytesIO()
    wb.save(excel_io)
    excel_io.seek(0)
    return excel_io

def append_to_biweekly(wb, metrics):
    ws = wb["BiWeeklyData"]
    next_row = ws.max_row + 1
    for col, value in enumerate(metrics, start=1):
        ws.cell(row=next_row, column=col, value=value)

st.set_page_config(page_title="GCC HR Tracker", layout="wide")
st.title("📈 GCC HR Dashboard – Bi-Weekly Tracker")

st.sidebar.header("Report Info")
report_date = st.sidebar.date_input("Reporting Date", value=datetime.date.today())
period = st.sidebar.text_input("Bi-Weekly Period (e.g., Aug 1–15)", "Aug 1–15, 2025")

st.header("Executive Summary")
total_headcount = st.number_input("Total Headcount", 0)
total_attrition = st.number_input("Attrition Rate (YTD %)", 0.0, 100.0)
avg_tenure = st.number_input("Avg. Tenure (Years)", 0.0, 50.0)
engagement_index = st.number_input("Engagement Index %", 0.0, 100.0)
enps = st.number_input("eNPS", -100, 100)
new_hires = st.number_input("New Hires This Period", 0)
vol_attr = st.number_input("Voluntary Attrition", 0)
inv_attr = st.number_input("Involuntary Attrition", 0)

if st.button("📥 Submit & Update Excel"):
    wb = load_template()
    append_to_biweekly(wb, [
        report_date, period, total_headcount, total_attrition,
        avg_tenure, engagement_index, enps,
        new_hires, vol_attr, inv_attr
    ])
    output = save_to_bytes(wb)
    st.success("Submitted and Excel updated ✅")
    st.download_button("📤 Download Updated Excel", output, "GCC_HR_Report_Updated.xlsx")

st.divider()
st.subheader("📊 Trend Visualization")

try:
    df = pd.read_excel(TEMPLATE_PATH, sheet_name="BiWeeklyData")
    if not df.empty:
        df.columns = [
            "Date", "Period", "Headcount", "AttritionRate", "AvgTenure",
            "EngagementIndex", "eNPS", "NewHires", "VolAttr", "InvAttr"
        ]
        df["Date"] = pd.to_datetime(df["Date"])

        col1, col2 = st.columns(2)
        with col1:
            st.altair_chart(alt.Chart(df).mark_line(point=True).encode(
                x="Date", y="Headcount", tooltip=["Period", "Headcount"]
            ).properties(title="Headcount Trend", width=400))

        with col2:
            st.altair_chart(alt.Chart(df).mark_line(point=True).encode(
                x="Date", y="AttritionRate", tooltip=["Period", "AttritionRate"]
            ).properties(title="Attrition Rate Trend", width=400))

        st.altair_chart(alt.Chart(df).mark_line(point=True).encode(
            x="Date", y="EngagementIndex", tooltip=["Period", "EngagementIndex"]
        ).properties(title="Employee Engagement Trend", width=800))
    else:
        st.info("No data available yet in BiWeeklyData sheet.")
except Exception as e:
    st.error(f"Error loading trend data: {e}")
