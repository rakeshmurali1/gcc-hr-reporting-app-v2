import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import datetime

# Load Excel Template
def load_template():
    return load_workbook('template.xlsx')

# Update Excel with form data
def update_template(wb, data):
    ws = wb.active
    
    ws['B2'] = data['report_period']
    ws['B3'] = data['report_date']
    ws['B4'] = data['total_headcount']
    ws['B5'] = data['total_attrition']
    
    row = 7
    for line in data['highlights'].split("\n"):
        ws[f'B{row}'] = line
        row += 1

    # Other sheets can be accessed and updated similarly
    return wb

# Download updated workbook
def generate_excel(wb):
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# Streamlit UI
st.set_page_config(page_title="GCC HR Reporting Dashboard", page_icon="📊", layout="wide")
st.title("📊 GCC HR Reporting Dashboard")

st.header("1. Executive Summary")
report_period = st.text_input("Report Period", "Q2 2025")
report_date = st.date_input("Reporting Date", value=datetime.today())
total_headcount = st.number_input("Total Headcount", min_value=0, step=1)
total_attrition = st.number_input("Total Attrition Rate (YTD)", min_value=0.0, step=0.1)
highlights = st.text_area("Key Highlights (one per line)")

st.divider()

st.header("2. Headcount Overview")
col1, col2, col3 = st.columns(3)
with col1:
    fte = st.number_input("FTEs", min_value=0, step=1)
    contractors = st.number_input("Contractors", min_value=0, step=1)
with col2:
    interns = st.number_input("Interns / Trainees", min_value=0, step=1)
    female = st.number_input("Female Employees", min_value=0, step=1)
with col3:
    pwd = st.number_input("Employees with Disabilities", min_value=0, step=1)
    location_dist = st.text_area("Location Distribution (e.g. Bangalore:10, Pune:5, Remote:2)")

st.divider()

st.header("3. Skill Mix / Competency Distribution")
skill_data = st.data_editor(pd.DataFrame({
    'Skill Domain': ["Data Engineering", "Data Science / ML", "DevOps / SRE", "Cybersecurity", "Business Analytics", "Product Mgmt"],
    'Headcount': [0]*6
}))

st.divider()

st.header("4. Role Mix")
role_data = st.data_editor(pd.DataFrame({
    'Role Category': ["Individual Contributors", "Team Leads / SMEs", "Mid-level Managers", "Senior Leadership", "Support Functions"],
    'Headcount': [0]*5
}))

st.divider()

st.header("5. Tenure / Age in Organization")
tenure_data = st.data_editor(pd.DataFrame({
    'Tenure Band': ["< 1 year", "1–3 years", "3–5 years", "5+ years"],
    'Headcount': [0]*4
}))

st.divider()

st.header("6. Attrition Analysis")
col1, col2, col3 = st.columns(3)
with col1:
    voluntary = st.number_input("Voluntary Attrition", min_value=0, step=1)
    involuntary = st.number_input("Involuntary Attrition", min_value=0, step=1)
with col2:
    exit_interview = st.selectbox("Exit Interviews Completed", ["Yes", "No"])
    reasons = st.text_area("Top Reasons for Exit")
with col3:
    backfill = st.text_input("Backfill Rate (%)", "80")

st.divider()

st.header("7. Diversity & Inclusion")
col1, col2, col3 = st.columns(3)
with col1:
    gender_div = st.text_input("Gender Diversity (% Female)", "33")
    age_avg = st.text_input("Avg. Age of Employees", "29")
with col2:
    returning_mothers = st.text_input("% Returning Mothers", "75")
    lgbtq = st.text_input("LGBTQ+ Inclusion", "3")
with col3:
    disability = st.text_input("Disability Inclusion", "2")

st.divider()

st.header("8. Hiring Overview")
col1, col2, col3 = st.columns(3)
with col1:
    new_hires = st.number_input("New Hires (This Period)", min_value=0)
    internal_moves = st.number_input("Internal Movements / Promotions", min_value=0)
with col2:
    lateral_ratio = st.text_input("Lateral vs Fresher Hiring (%)", "70:30")
    offer_to_join = st.text_input("Offer-to-Join Ratio", "85%")
with col3:
    avg_time_fill = st.text_input("Avg. Time to Fill (in Days)", "25")

st.divider()

st.header("9. Learning & Development")
col1, col2, col3 = st.columns(3)
with col1:
    avg_training = st.text_input("Avg. Training Hours per Employee", "12")
    trained_pct = st.text_input("% Employees Trained in Last 6 Months", "90")
with col2:
    certs = st.text_input("Certifications Completed", "50")
    mgr_coverage = st.text_input("Managerial Development Coverage", "60")

st.divider()

st.header("10. Employee Engagement")
col1, col2, col3 = st.columns(3)
with col1:
    engagement_idx = st.text_input("Engagement Index (%)", "78")
    enps = st.text_input("eNPS", "45")
with col2:
    survey_rate = st.text_input("Survey Response Rate (%)", "87")
    awards = st.text_input("Recognition Awards Issued", "35")

st.divider()

st.header("11. Compliance & Policy Metrics")
col1, col2, col3 = st.columns(3)
with col1:
    posh = st.selectbox("POSH Committee Coverage", ["Yes", "No"])
with col2:
    bgv = st.selectbox("Background Verifications Completed", ["Yes", "No"])
with col3:
    mandatory_training = st.selectbox("Mandatory Training Completion", ["Yes", "No"])
    grievances = st.text_input("Grievances Resolved", "5")

# Button
if st.button("📥 Generate Report"):
    wb = load_template()
    report_data = {
        'report_period': report_period,
        'report_date': report_date.strftime('%Y/%m/%d'),
        'total_headcount': total_headcount,
        'total_attrition': total_attrition,
        'highlights': highlights
    }
    wb = update_template(wb, report_data)
    file = generate_excel(wb)
    st.success("✅ Report ready!")
    st.download_button("⬇️ Download Excel Report", file_name="GCC_HR_Report.xlsx", data=file, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
