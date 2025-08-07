
import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="GCC HR Reporting | SRKay", layout="wide")
st.title("📊 GCC HR Reporting Dashboard")

st.header("1. Executive Summary")
report_period = st.text_input("Report Period", "Q2 2025")
report_date = st.date_input("Reporting Date", datetime.today())
headcount = st.number_input("Total Headcount", min_value=0)
attrition_rate = st.text_input("Total Attrition Rate (YTD)")
key_highlights = st.text_area("Key Highlights")

if st.button("📤 Generate Report"):
    # Load the Excel template
    template = "template.xlsx"
    wb = load_workbook(template)
    ws = wb["Executive Summary"]
    ws["B1"] = "Report Period"; ws["C1"] = report_period
    ws["B2"] = "Reporting Date"; ws["C2"] = report_date.strftime("%d-%m-%Y")
    ws["B3"] = "Total Headcount"; ws["C3"] = headcount
    ws["B4"] = "Total Attrition Rate (YTD)"; ws["C4"] = attrition_rate
    ws["B5"] = "Key Highlights"; ws["C5"] = key_highlights

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.download_button(
        label="📥 Download GCC HR Report",
        data=output,
        file_name=f"GCC_HR_Report_{datetime.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
