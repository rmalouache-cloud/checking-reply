import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image
from fpdf import FPDF
import tempfile

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Container Dashboard", layout="wide")

# =========================
# 📝 STUDY INPUT (FORCED DISPLAY)
# =========================
st.subheader("📝 Study Information")

colA, colB, colC = st.columns(3)

with colA:
    packing_type = st.selectbox(
        "Type of Packing List",
        ["Panel", "SP", "SP/MainBoard", "OC"]
    )

with colB:
    model = st.text_input("Model (ex: Mini LED)")

with colC:
    odf = st.text_input("ODF (ex: IDL2500)")

# =========================
# LOGOS + HEADER
# =========================
container_logo = Image.open("conteneur_logo.png")
stream_logo = Image.open("stream_logo.png")

col1, col2, col3 = st.columns([1, 5, 1])

with col1:
    st.image(container_logo, width=400)

with col2:
    st.title(" Container Filling Industrial Dashboard")
    st.caption("Supply Chain Analysis - BOM & Packing Control")

with col3:
    st.image(stream_logo, width=800)

# =========================
# 📘 USER GUIDE (FR + EN)
# =========================
with st.expander("📘 Manuel d'utilisation / User Guide"):

    st.markdown("""
# 🇫🇷 Manuel d'utilisation

### 🎯 Objectif
Analyser le taux de remplissage des conteneurs à partir d’un fichier Excel.

### 📂 Format requis
- CONTAINER NO  
- CTNER.SIZE (20GP, 40GP, 40HQ)  
- CBM  

### 🚀 Étapes
1. Upload fichier Excel  
2. Vérifier les données  
3. Lire les résultats  
4. Visualiser graphique  
5. Télécharger Excel ou PDF  

### 📏 Règles
- 20GP → 33  
- 40GP → 67  
- 40HQ → 76  

👉 OK ≥ 70%  
👉 NON CONFORME < 70%  

---

# 🇬🇧 User Guide

### 🎯 Purpose
Analyze container filling rate from an Excel file.

### 📂 Required columns
- CONTAINER NO  
- CTNER.SIZE  
- CBM  

### 🚀 Steps
1. Upload Excel file  
2. Check preview  
3. Analyze results  
4. View chart  
5. Download Excel/PDF  
""")

# =========================
# UPLOAD FILE
# =========================
file = st.file_uploader("Upload Packing Excel file", type=["xlsx"])

if file is not None:

    df = pd.read_excel(file)
    df.columns = df.columns.str.strip()

    # =========================
    # DYNAMIC TITLE
    # =========================
    custom_title = f"Container Filling Industrial Dashboard of {packing_type} of {model}__{odf}"

    st.markdown(f"### 📌 {custom_title}")

    st.write("📄 Data Preview")
    st.dataframe(df)

    # =========================
    # DETECT CBM COLUMN
    # =========================
    cbm_col = None
    for col in df.columns:
        if "CBM" in col.upper():
            cbm_col = col
            break

    if cbm_col is None:
        st.error("❌ CBM column not found")
    else:

        # =========================
        # GROUP BY CONTAINER
        # =========================
        summary = df.groupby(
            ["CONTAINER NO", "CTNER.SIZE"], as_index=False
        ).agg({cbm_col: "sum"})

        summary.rename(columns={cbm_col: "TOTAL_VOLUME"}, inplace=True)

        # =========================
        # CAPACITY
        # =========================
        capacity_map = {
            "20GP": 33,
            "40GP": 67,
            "40HQ": 76
        }

        summary["CAPACITY"] = summary["CTNER.SIZE"].map(capacity_map)

        # =========================
        # CALCULATION
        # =========================
        summary["FILL_RATE_%"] = summary["TOTAL_VOLUME"] * 100 / summary["CAPACITY"]

        # =========================
        # STATUS
        # =========================
        summary["STATUS"] = summary["FILL_RATE_%"].apply(
            lambda x: "OK" if x >= 70 else "NON CONFORME"
        )

        # =========================
        # COLOR TABLE
        # =========================
        def color_status(val):
            if val == "OK":
                return "background-color: lightgreen"
            else:
                return "background-color: lightcoral"

        styled = summary.style.map(color_status, subset=["STATUS"])

        st.subheader("📊 Result Table")
        st.write(styled)

        # =========================
        # CHART
        # =========================
        st.subheader("📈 Filling Rate Chart")

        fig, ax = plt.subplots(figsize=(7, 3))

        ax.bar(summary["CONTAINER NO"], summary["FILL_RATE_%"])
        ax.axhline(70, linestyle="--", color="red")

        ax.set_ylabel("Filling Rate %")
        ax.set_xlabel("Container")
        ax.set_title("Container Filling Rate")

        plt.xticks(rotation=45, ha='right')

        st.pyplot(fig)

        # =========================
        # EXPORT EXCEL
        # =========================
        excel_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        summary.to_excel(excel_file.name, index=False)

        st.download_button(
            label="📥 Download Excel Result",
            data=open(excel_file.name, "rb"),
            file_name="container_analysis.xlsx"
        )

        # =========================
        # EXPORT PDF
        # =========================
        chart_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
        fig.savefig(chart_path, bbox_inches="tight")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", "B", 14)

        # ✅ TITLE DYNAMIQUE
        pdf.cell(200, 10, txt=custom_title, ln=True, align="C")
        pdf.ln(5)

        pdf.set_font("Arial", "B", 9)

        headers = [
            "CONTAINER NO",
            "SIZE",
            "TOTAL_VOL",
            "CAPACITY",
            "FILL_RATE %",
            "STATUS"
        ]

        col_widths = [35, 20, 30, 25, 25, 30]

        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, h, border=1, align="C")
        pdf.ln()

        pdf.set_font("Arial", size=9)

        for i, row in summary.iterrows():
            pdf.cell(col_widths[0], 8, str(row["CONTAINER NO"]), border=1)
            pdf.cell(col_widths[1], 8, str(row["CTNER.SIZE"]), border=1)
            pdf.cell(col_widths[2], 8, f"{row['TOTAL_VOLUME']:.2f}", border=1)
            pdf.cell(col_widths[3], 8, str(row["CAPACITY"]), border=1)
            pdf.cell(col_widths[4], 8, f"{row['FILL_RATE_%']:.2f}%", border=1)
            pdf.cell(col_widths[5], 8, str(row["STATUS"]), border=1)
            pdf.ln()

        pdf.ln(5)
        pdf.set_font("Arial", "B", 11)
        pdf.cell(200, 10, txt="Filling Rate Chart", ln=True)

        pdf.image(chart_path, w=180)

        pdf_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf.output(pdf_file.name)

        with open(pdf_file.name, "rb") as f:
            st.download_button(
                label="📄 Download PDF Dashboard",
                data=f,
                file_name="container_dashboard.pdf"
            )

        st.success("✅ Analysis completed successfully")
