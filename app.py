import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# DOWNLOAD FUNCTION
# =========================================================
def download_excel(dataframe, filename):

    output = pd.ExcelWriter(
        filename,
        engine="openpyxl"
    )

    dataframe.to_excel(
        output,
        index=False
    )

    output.close()

    with open(filename, "rb") as f:

        st.download_button(
            label="⬇ Download Excel Report",
            data=f,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="RR Analysis Dashboard",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# CUSTOM CSS
# =========================================================
st.markdown("""
<style>

.main {
    background-color: #f5f7fa;
}

.block-container {
    padding-top: 1rem;
}

h1, h2, h3 {
    color: #003366;
    font-weight: 700;
}

[data-testid="stSidebar"] {
    background-color: #0f172a;
}

[data-testid="stSidebar"] * {
    color: white;
}

.metric-card {
    background: white;
    padding: 20px;
    border-radius: 15px;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.08);
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# TITLE
# =========================================================
st.title("📊 RR ANALYSIS DASHBOARD")
st.markdown("### Upload Excel File for Automatic Analysis")

# =========================================================
# FILE UPLOAD
# =========================================================
uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

# =========================================================
# MAIN PROCESS
# =========================================================
if uploaded_file:

    # =====================================================
    # LOAD DATA
    # =====================================================
    df = pd.read_excel(uploaded_file)

    # =====================================================
    # CLEAN COLUMNS
    # =====================================================
    df.columns = df.columns.str.strip()

    # =====================================================
    # SIDEBAR MENU
    # =====================================================
    st.sidebar.title("📁 Navigation")

    menu = st.sidebar.radio(
        "Select Report",
        [
            "Dashboard Overview",
            "Region vs Model",
            "Model vs Product Type",
            "Product Type vs Warranty",
            "Model vs Type of Work",
            "Modelwise Mod Board vs Warranty",
            "CT Overall Scrapped Items"
        ]
    )

    # =====================================================
    # DASHBOARD OVERVIEW
    # =====================================================
    if menu == "Dashboard Overview":

        st.header("📌 Dashboard Overview")

        total_records = len(df)
        total_models = df["PRODUCT_MODEL"].nunique()
        total_regions = df["REGION"].nunique()
        total_scrap = len(
            df[df["REPORT_TYPE"].astype(str).str.lower() == "scraplist"]
        )

        col1, col2, col3, col4 = st.columns(4)

        col1.metric("Total Records", total_records)
        col2.metric("Total Models", total_models)
        col3.metric("Total Regions", total_regions)
        col4.metric("Scrapped Items", total_scrap)

        st.markdown("---")

        # ==========================================
        # MODEL DISTRIBUTION
        # ==========================================
        model_counts = (
            df["PRODUCT_MODEL"]
            .value_counts()
            .reset_index()
        )

        model_counts.columns = ["PRODUCT_MODEL", "COUNT"]

        fig = px.bar(
            model_counts,
            x="PRODUCT_MODEL",
            y="COUNT",
            text_auto=True,
            title="Product Model Distribution"
        )

        st.plotly_chart(fig, use_container_width=True)

        # ==========================================
        # REGION DISTRIBUTION
        # ==========================================
        region_counts = (
            df["REGION"]
            .value_counts()
            .reset_index()
        )

        region_counts.columns = ["REGION", "COUNT"]

        fig2 = px.pie(
            region_counts,
            names="REGION",
            values="COUNT",
            title="Region Distribution"
        )

        st.plotly_chart(fig2, use_container_width=True)

    # =====================================================
    # REGION VS MODEL
    # =====================================================
    elif menu == "Region vs Model":

        st.header("📍 CT OVERALL DEFECTIVES – REGION VS MODEL")

        df_unique = df.drop_duplicates(
            subset=["REGION", "PRODUCT_MODEL", "UNIT_SLNO"]
        )

        summary_table = pd.pivot_table(
            df_unique,
            index="REGION",
            columns="PRODUCT_MODEL",
            values="UNIT_SLNO",
            aggfunc=pd.Series.nunique,
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        summary_table = summary_table.reset_index()

        summary_table.rename(
            columns={"REGION": "Region / Model"},
            inplace=True
        )

        overall_total = summary_table.loc[
            summary_table["Region / Model"] == "Grand Total",
            "Grand Total"
        ].values[0]

        summary_table["Contribution"] = (
            summary_table["Grand Total"] / overall_total * 100
        ).round(1)

        summary_table.loc[
            summary_table["Region / Model"] == "Grand Total",
            "Contribution"
        ] = 100.0

        st.dataframe(summary_table, use_container_width=True)
        
        download_excel(
            summary_table,
            "CT_Overall_Defectives_Region_vs_Model.xlsx"
        )

        chart_df = summary_table[
            summary_table["Region / Model"] != "Grand Total"
        ]

        fig = px.bar(
            chart_df,
            x="Region / Model",
            y="Grand Total",
            color="Grand Total",
            text_auto=True,
            title="Region Wise Total Defectives"
        )

        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # MODEL VS PRODUCT TYPE
    # =====================================================
    elif menu == "Model vs Product Type":

        st.header("🛠 MODEL VS PRODUCT TYPE")

        df_unique = df.drop_duplicates(
            subset=["PRODUCT_MODEL", "DEF_TYPE", "UNIT_SLNO"]
        )

        summary_table = pd.pivot_table(
            df_unique,
            index="PRODUCT_MODEL",
            columns="DEF_TYPE",
            values="UNIT_SLNO",
            aggfunc=pd.Series.nunique,
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        summary_table = summary_table.reset_index()

        summary_table.rename(
            columns={"PRODUCT_MODEL": "Model / Type"},
            inplace=True
        )

        overall_total = summary_table.loc[
            summary_table["Model / Type"] == "Grand Total",
            "Grand Total"
        ].values[0]

        summary_table["Contribution %"] = (
            summary_table["Grand Total"] / overall_total * 100
        ).round(1)

        summary_table.loc[
            summary_table["Model / Type"] == "Grand Total",
            "Contribution %"
        ] = 100.0

        st.dataframe(summary_table, use_container_width=True)

        download_excel(
            summary_table,
            "CT_Overall_Defectives_Model_vs_ProductType.xlsx"
        )

        chart_df = summary_table[
            summary_table["Model / Type"] != "Grand Total"
        ]

        fig = px.bar(
            chart_df,
            x="Model / Type",
            y="Grand Total",
            color="Grand Total",
            text_auto=True,
            title="Model vs Product Type"
        )

        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # PRODUCT TYPE VS WARRANTY
    # =====================================================
    elif menu == "Product Type vs Warranty":

        st.header("🧾 PRODUCT TYPE VS WARRANTY")

        df_unique = df.drop_duplicates(
            subset=["PRODUCT_MODEL", "UNIT_STATUS", "UNIT_SLNO"]
        )

        summary_table = pd.pivot_table(
            df_unique,
            index="PRODUCT_MODEL",
            columns="UNIT_STATUS",
            values="UNIT_SLNO",
            aggfunc=pd.Series.nunique,
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        summary_table = summary_table.reset_index()

        summary_table.rename(
            columns={"PRODUCT_MODEL": "Product Type / Warranty"},
            inplace=True
        )

        overall_total = summary_table.loc[
            summary_table["Product Type / Warranty"] == "Grand Total",
            "Grand Total"
        ].values[0]

        summary_table["Contribution %"] = (
            summary_table["Grand Total"] / overall_total * 100
        ).round(1)

        summary_table.loc[
            summary_table["Product Type / Warranty"] == "Grand Total",
            "Contribution %"
        ] = 100.0

        st.dataframe(summary_table, use_container_width=True)
        
        download_excel(
            summary_table,
            "CT_ProductType_vs_Warranty.xlsx"
        )

        chart_df = summary_table[
            summary_table["Product Type / Warranty"] != "Grand Total"
        ]

        fig = px.bar(
            chart_df,
            x="Product Type / Warranty",
            y="Grand Total",
            color="Grand Total",
            text_auto=True,
            title="Warranty Distribution"
        )

        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # MODEL VS TYPE OF WORK
    # =====================================================
    elif menu == "Model vs Type of Work":

        st.header("🏭 MODEL VS TYPE OF WORK")

        df_unique = df.drop_duplicates(
            subset=["PRODUCT_MODEL", "TYPE_OF_WORK", "UNIT_SLNO"]
        )

        summary_table = pd.pivot_table(
            df_unique,
            index="PRODUCT_MODEL",
            columns="TYPE_OF_WORK",
            values="UNIT_SLNO",
            aggfunc=pd.Series.nunique,
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        summary_table = summary_table.reset_index()

        summary_table.rename(
            columns={"PRODUCT_MODEL": "Model / Type of Work"},
            inplace=True
        )

        overall_total = summary_table.loc[
            summary_table["Model / Type of Work"] == "Grand Total",
            "Grand Total"
        ].values[0]

        summary_table["Contribution %"] = (
            summary_table["Grand Total"] / overall_total * 100
        ).round(1)

        summary_table.loc[
            summary_table["Model / Type of Work"] == "Grand Total",
            "Contribution %"
        ] = 100.0

        st.dataframe(summary_table, use_container_width=True)

        download_excel(
            summary_table,
            "CT_Model_vs_TypeOfWork.xlsx"
        )

        chart_df = summary_table[
            summary_table["Model / Type of Work"] != "Grand Total"
        ]

        fig = px.bar(
            chart_df,
            x="Model / Type of Work",
            y="Grand Total",
            color="Grand Total",
            text_auto=True,
            title="Type Of Work Distribution"
        )

        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # MODELWISE MOD BOARD VS WARRANTY
    # =====================================================
    elif menu == "Modelwise Mod Board vs Warranty":

        st.header("🔧 MODELWISE MOD_BRD_NAME VS WARRANTY")

        models = [
            "BODYTOM NL 4000",
            "CERETOM NL 3000",
            "MAT5000 IV",
            "OMNITOM NL 5000"
        ]

        selected_model = st.selectbox(
            "Select Product Model",
            models
        )

        model_df = df[
            df["PRODUCT_MODEL"]
            .astype(str)
            .str.upper()
            .str.strip() == selected_model.upper()
        ]

        df_unique = model_df.drop_duplicates(
            subset=["MOD_BRD_NAME", "UNIT_STATUS", "UNIT_SLNO"]
        )

        summary_table = pd.pivot_table(
            df_unique,
            index="MOD_BRD_NAME",
            columns="UNIT_STATUS",
            values="UNIT_SLNO",
            aggfunc=pd.Series.nunique,
            fill_value=0,
            margins=True,
            margins_name="Grand Total"
        )

        summary_table = summary_table.reset_index()

        summary_table.rename(
            columns={"MOD_BRD_NAME": "Module Board / Warranty"},
            inplace=True
        )

        overall_total = summary_table.loc[
            summary_table["Module Board / Warranty"] == "Grand Total",
            "Grand Total"
        ].values[0]

        summary_table["Contribution %"] = (
            summary_table["Grand Total"] / overall_total * 100
        ).round(1)

        summary_table.loc[
            summary_table["Module Board / Warranty"] == "Grand Total",
            "Contribution %"
        ] = 100.0

        st.dataframe(summary_table, use_container_width=True)
        
        download_excel(
            summary_table,
            f"{selected_model}_ModBoard_vs_Warranty.xlsx"
        )

        chart_df = summary_table[
            summary_table["Module Board / Warranty"] != "Grand Total"
        ]

        fig = px.bar(
            chart_df.head(15),
            x="Module Board / Warranty",
            y="Grand Total",
            color="Grand Total",
            text_auto=True,
            title=f"{selected_model} - Top Module Boards"
        )

        st.plotly_chart(fig, use_container_width=True)

    # =====================================================
    # SCRAPPED ITEMS
    # =====================================================
    elif menu == "CT Overall Scrapped Items":

        st.header("♻️ CT OVERALL SCRAPPED ITEMS")

        scrap_df = df[
            df["REPORT_TYPE"]
            .astype(str)
            .str.lower() == "scraplist"
        ].copy()

        models = [
            "CERETOM NL 3000",
            "BODYTOM NL 4000",
            "MAT5000 IV",
            "OMNITOM NL 5000"
        ]

        scrap_df = scrap_df[
            scrap_df["PRODUCT_MODEL"]
            .astype(str)
            .str.strip()
            .isin(models)
        ]

        scrap_df["DEF_MOD_BRD_NAME"] = (
            scrap_df["DEF_MOD_BRD_NAME"]
            .astype(str)
            .str.strip()
            .str.title()
        )

        scrap_df["DEF_TYPE"] = (
            scrap_df["DEF_TYPE"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        final_rows = []

        overall_consumables = 0
        overall_spare = 0
        overall_total = 0

        for model in models:

            model_df = scrap_df[
                scrap_df["PRODUCT_MODEL"]
                .astype(str)
                .str.strip() == model
            ]

            if model_df.empty:
                continue

            pivot = pd.pivot_table(
                model_df,
                index="DEF_MOD_BRD_NAME",
                columns="DEF_TYPE",
                values="SC_REF_NO",
                aggfunc="count",
                fill_value=0
            )

            for col in ["CONSUMABLES", "SPARE"]:
                if col not in pivot.columns:
                    pivot[col] = 0

            pivot = pivot[["CONSUMABLES", "SPARE"]]

            pivot["Grand Total"] = (
                pivot["CONSUMABLES"] +
                pivot["SPARE"]
            )

            pivot = pivot.sort_values(
                by="Grand Total",
                ascending=False
            )

            pivot = pivot.reset_index()

            pivot.rename(
                columns={
                    "DEF_MOD_BRD_NAME": "Defective Items"
                },
                inplace=True
            )

            model_consumables = int(pivot["CONSUMABLES"].sum())
            model_spare = int(pivot["SPARE"].sum())
            model_total = int(pivot["Grand Total"].sum())

            overall_consumables += model_consumables
            overall_spare += model_spare
            overall_total += model_total

            st.subheader(model)

            st.dataframe(
                pivot,
                use_container_width=True
            )
            
            download_excel(
                pivot,
                f"{model}_Scrapped_Items.xlsx"
            )

            fig = px.bar(
                pivot,
                x="Defective Items",
                y="Grand Total",
                color="Grand Total",
                text_auto=True,
                title=f"{model} - Scrapped Items"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )

        # ==========================================
        # OVERALL SUMMARY
        # ==========================================
        st.markdown("---")

        st.subheader("📌 Overall Summary")

        col1, col2, col3 = st.columns(3)

        col1.metric(
            "Consumables",
            overall_consumables
        )

        col2.metric(
            "Spares",
            overall_spare
        )

        col3.metric(
            "Grand Total",
            overall_total
        )
        overall_summary = pd.DataFrame({
            "Category": ["Consumables", "Spares", "Grand Total"],
            "Count": [
                overall_consumables,
                overall_spare,
                overall_total
            ]
        })
        
        download_excel(
            overall_summary,
            "CT_Overall_Scrap_Summary.xlsx"
        )

else:
    st.info("📂 Please upload an Excel file to continue.")
