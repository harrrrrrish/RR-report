import streamlit as st
import pandas as pd

st.set_page_config(layout="wide")
st.title("📊 PORTABLE CT FAILURE ANALYSIS FOR YEAR 2025")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

def show_section(title, df, note=None):
    st.markdown(f"## {title}")
    if note:
        st.caption(note)
    st.dataframe(df, use_container_width=True)

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)

    # -------- SECTION 1 --------
    df1 = pd.read_excel(xls, sheet_name=0)
    show_section(
        "1. CT INSTALLATION BASE OVERALL",
        df1,
        "*Includes Open Installations"
    )

    # -------- SECTION 2 --------
    df2 = pd.read_excel(xls, sheet_name=1)
    show_section(
        "2. CT OVERALL INSTALLATION BASE - YEAR WISE",
        df2,
        "*Includes Open Installations"
    )

    # -------- SECTION 3 --------
    df3 = pd.read_excel(xls, sheet_name=2)
    show_section(
        "3. CT INSTALLATION BASE – REGION VS MODEL",
        df3,
        "*Includes Open Installations"
    )

    # -------- SECTION 4 --------
    df4 = pd.read_excel(xls, sheet_name=3)
    show_section(
        "4. CT OVERALL DEFECTIVES – REGION VS MODEL",
        df4
    )

    # -------- SECTION 5 --------
    df5 = pd.read_excel(xls, sheet_name=4)
    show_section(
        "5. CT OVERALL DEFECTIVES – MODEL VS PRODUCT TYPE",
        df5
    )

    # -------- SECTION 6 --------
    df6 = pd.read_excel(xls, sheet_name=5)
    show_section(
        "6. CT OVERALL DEFECTIVES – MODEL VS WARRANTY",
        df6
    )

    # -------- SECTION 7 --------
    df7 = pd.read_excel(xls, sheet_name=6)
    show_section(
        "7. CT OVERALL DEFECTIVES – ACTION TAKEN",
        df7
    )

    # -------- SECTION 8 --------
    df8 = pd.read_excel(xls, sheet_name=7)
    show_section(
        "8. BODYTOM NL – 4000 OVERALL FAILURE",
        df8
    )

    # -------- SECTION 9 --------
    df9 = pd.read_excel(xls, sheet_name=8)
    show_section(
        "9. CERETOM NL – 3000 OVERALL FAILURE",
        df9
    )

    # -------- SECTION 10 --------
    df10 = pd.read_excel(xls, sheet_name=9)
    show_section(
        "10. OMNITOM NL – 5000 OVERALL FAILURE",
        df10
    )

    # -------- SECTION 11 --------
    df11 = pd.read_excel(xls, sheet_name=10)
    show_section(
        "11. CT OVERALL SCRAPPED ITEMS",
        df11
    )

    # -------- KPI --------
    st.markdown("## 📌 FINAL INSIGHTS")

    total_install = df1.iloc[-1, -1]
    total_failure = df4.iloc[-1, -2]
    total_scrap = df11.iloc[:, -1].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Installations", total_install)
    c2.metric("Total Failures", total_failure)
    c3.metric("Total Scrapped", total_scrap)

else:
    st.info("Upload Excel with same structure (11 sheets)")
