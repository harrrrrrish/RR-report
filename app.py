import streamlit as st
import pandas as pd

st.set_page_config(page_title="CT Failure Analysis 2025", layout="wide")
st.title("📊 PORTABLE CT FAILURE ANALYSIS FOR YEAR 2025")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    # Expected Titles (FIXED STRUCTURE)
    section_titles = [
        "1. CT INSTALLATION BASE OVERALL",
        "2. YEAR-WISE INSTALLATION",
        "3. REGION VS MODEL",
        "4. DEFECTIVES – REGION VS MODEL",
        "5. DEFECTIVES – MODEL VS PRODUCT TYPE",
        "6. DEFECTIVES – WARRANTY",
        "7. ACTION TAKEN",
        "8. BODYTOM NL4000 FAILURES (FULL)",
        "9. CERETOM NL3000 FAILURES (FULL)",
        "10. OMNITOM NL5000 FAILURES (FULL)",
        "11. SCRAPPED ITEMS"
    ]

    all_data = {}

    # ---------------- DISPLAY ALL TABLES ----------------
    for i, sheet in enumerate(sheets):

        df = pd.read_excel(uploaded_file, sheet_name=sheet)

        # Preserve EXACT formatting
        df.columns = df.columns.astype(str)

        title = section_titles[i] if i < len(section_titles) else f"Sheet {i+1}"

        st.header(title)
        st.dataframe(df, use_container_width=True)

        all_data[title] = df

    # ---------------- KPI CALCULATION ----------------
    st.header("📌 FINAL INSIGHTS")

    total_installations = 0
    total_failures = 0
    total_scrapped = 0

    # Page 1 → Installations
    if "1. CT INSTALLATION BASE OVERALL" in all_data:
        df1 = all_data["1. CT INSTALLATION BASE OVERALL"]
        if "Grand Total" in df1.columns:
            total_installations = df1["Grand Total"].iloc[-1]

    # Page 4 or 5 → Failures
    if "4. DEFECTIVES – REGION VS MODEL" in all_data:
        df4 = all_data["4. DEFECTIVES – REGION VS MODEL"]
        if "Total" in df4.columns:
            total_failures = df4["Total"].iloc[-1]

    # Page 11 → Scrapped
    if "11. SCRAPPED ITEMS" in all_data:
        df11 = all_data["11. SCRAPPED ITEMS"]
        if "Count" in df11.columns:
            total_scrapped = df11["Count"].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Installations", total_installations)
    col2.metric("Total Failures", total_failures)
    col3.metric("Total Scrapped", total_scrapped)

else:
    st.info("Upload Excel file with same structure (11 sheets)")
