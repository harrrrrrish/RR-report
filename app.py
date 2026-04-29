import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")
st.title("📊 CT FAILURE INTELLIGENCE DASHBOARD")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    # ---------------- CLEANING ----------------
    df.columns = df.columns.str.strip()

    # Convert dates
    date_cols = ["ENTRY DATE", "FRN_DATE", "CURRENT DATE"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # ---------------- KPI ----------------
    st.header("📌 KEY METRICS")

    total_cases = len(df)
    total_scrap = len(df[df["TYPE_OF_WORK"] == "SCRAPPED"])
    total_completed = len(df[df["TYPE_OF_WORK"] == "COMPLETED"])
    total_pending = len(df[df["STATUS"] == "Pending"])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Cases", total_cases)
    c2.metric("Scrapped", total_scrap)
    c3.metric("Completed", total_completed)
    c4.metric("Pending", total_pending)

    # ---------------- REGION ANALYSIS ----------------
    st.header("🌍 Region-wise Failures")

    region_df = df["REGION"].value_counts().reset_index()
    region_df.columns = ["Region", "Count"]

    fig1 = px.bar(region_df, x="Region", y="Count", title="Failures by Region")
    st.plotly_chart(fig1, use_container_width=True)

    # ---------------- MODEL ANALYSIS ----------------
    st.header("🏥 Model-wise Failures")

    model_df = df["PRODUCT_MODEL"].value_counts().reset_index()
    model_df.columns = ["Model", "Count"]

    fig2 = px.bar(model_df, x="Model", y="Count", title="Failures by Model")
    st.plotly_chart(fig2, use_container_width=True)

    # ---------------- DEFECT TYPE ----------------
    st.header("⚙️ Defect Type Analysis")

    defect_df = df["DEF_TYPE"].value_counts().reset_index()
    defect_df.columns = ["Defect Type", "Count"]

    fig3 = px.pie(defect_df, names="Defect Type", values="Count")
    st.plotly_chart(fig3, use_container_width=True)

    # ---------------- SCRAP ANALYSIS ----------------
    st.header("🗑 Scrap Components")

    scrap_df = df[df["TYPE_OF_WORK"] == "SCRAPPED"]

    if not scrap_df.empty:
        comp_df = scrap_df["MOD_BRD_NAME"].value_counts().reset_index()
        comp_df.columns = ["Component", "Count"]

        fig4 = px.bar(comp_df, x="Component", y="Count", title="Most Scrapped Components")
        st.plotly_chart(fig4, use_container_width=True)

    # ---------------- ENGINEER PERFORMANCE ----------------
    st.header("👨‍🔧 Engineer Performance")

    eng_df = df["ENGINEER"].value_counts().reset_index()
    eng_df.columns = ["Engineer", "Cases"]

    fig5 = px.bar(eng_df, x="Engineer", y="Cases")
    st.plotly_chart(fig5, use_container_width=True)

    # ---------------- PENDING DAYS ----------------
    st.header("⏳ Pending Days Analysis")

    if "PENDING DAYS(URP)" in df.columns:
        df["PENDING DAYS(URP)"] = pd.to_numeric(df["PENDING DAYS(URP)"], errors='coerce')

        fig6 = px.histogram(df, x="PENDING DAYS(URP)", nbins=20)
        st.plotly_chart(fig6, use_container_width=True)

    # ---------------- RAW DATA ----------------
    st.header("📄 Full Data Table")
    st.dataframe(df, use_container_width=True)

else:
    st.info("Upload your Excel file")
