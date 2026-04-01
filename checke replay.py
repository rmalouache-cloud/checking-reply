import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Verification Tool")

# ==============================
# 1. UPLOAD
# ==============================
main_file = st.file_uploader("📂 Upload Main File", type=["xlsx"])
frs_files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# ==============================
# EXTRACTION FRS FILE NAME
# ==============================
def extract_frs(text):
    text = str(text)

    match = re.search(r'"(.*?)"', text)
    if match:
        return match.group(1)

    match = re.search(r'(1RT\w+)', text)
    if match:
        return match.group(1)

    return None

# ==============================
# PROCESS
# ==============================
if main_file and frs_files:

    main_sheets = pd.read_excel(main_file, sheet_name=None)

    st.subheader("📝 Enter ODF per Model")

    odf_inputs = {}
    for sheet in main_sheets.keys():
        odf_inputs[sheet] = st.text_input(f"ODF for {sheet}")

    if st.button("▶️ Start Verification"):

        st.info("⏳ Processing...")

        # ==========================
        # LOAD FRS FILES
        # ==========================
        frs_dict = {}

        for file in frs_files:
            name = file.name.replace(".xlsx", "")
            df = pd.concat(pd.read_excel(file, sheet_name=None).values())

            # Nettoyage colonnes
            df.columns = df.columns.str.strip().str.upper()

            # Nettoyage données
            df["PART NO."] = df["PART NO."].astype(str).str.strip().str.upper()
            df["ODF"] = df["ODF"].astype(str).str.strip().str.upper()

            frs_dict[name] = df

        results = []

        # ==========================
        # LOOP SHEETS
        # ==========================
        for sheet_name, df in main_sheets.items():

            st.markdown(f"---\n### 📄 Sheet: {sheet_name}")

            # Nettoyage colonnes
            df.columns = df.columns.str.strip().str.upper()
            df = df.fillna("")

            st.write("📌 Colonnes :", df.columns.tolist())

            odf = odf_inputs.get(sheet_name, "").strip().upper()

            if odf == "":
                st.warning(f"⚠️ ODF manquant pour {sheet_name}")
                continue

            # ==========================
            # FILTRE REMARKS
            # ==========================
            if "REMARKS" not in df.columns:
                st.error("❌ Colonne REMARKS manquante")
                continue

            df = df[df["REMARKS"].astype(str).str.upper().str.contains("MISSING|SHORTAGE", na=False)]

            st.write("🔎 After REMARKS filter:", df.shape)

            # ==========================
            # FILTRE MOKA
            # ==========================
            if "MOKA REPLY" not in df.columns:
                st.error("❌ Colonne MOKA REPLY manquante")
                continue

            df = df[df["MOKA REPLY"].astype(str).str.lower().str.contains("enough", na=False)]

            st.write("🔎 After MOKA filter:", df.shape)

            if df.empty:
                st.warning("⚠️ No data after filtering")
                continue

            # ==========================
            # TRAITEMENT LIGNES
            # ==========================
            for _, row in df.iterrows():

                # PART N
                part_no = str(row.get("PART N", "")).strip().upper()

                # QTY FOR dynamic
                qty_needed_col = [col for col in df.columns if "QTY FOR" in col]

                if not qty_needed_col:
                    continue

                qty_needed = row[qty_needed_col[0]]

                # PACKING LIST
                qty_sent = row.get("PACKING LIST QTY", 0)

                # OVERSENT REPLY
                oversent_reply = row.get("OVERSENT QTY", 0)

                # MOKA REPLY
                moka = row.get("MOKA REPLY", "")

                frs_name = extract_frs(moka)

                if not frs_name or frs_name not in frs_dict:
                    continue

                frs_df = frs_dict[frs_name]

                # Match FRS
                match = frs_df[
                    (frs_df["PART NO."] == part_no) &
                    (frs_df["ODF"] == odf)
                ]

                if match.empty:
                    results.append([sheet_name, part_no, "NOT FOUND"])
                    continue

                idx = match.index[0]

                # OLD OVERSENT (ligne précédente)
                if idx > 0:
                    old_oversent = frs_df.iloc[idx - 1].get("OVERSENT QTY", 0)
                else:
                    old_oversent = 0

                # CALCUL
                calc = (old_oversent - qty_needed) + qty_sent

                # CHECK
                if calc == oversent_reply:
                    check = "OK"
                else:
                    check = "NON CONFORME"

                results.append([
                    sheet_name,
                    part_no,
                    qty_needed,
                    old_oversent,
                    qty_sent,
                    calc,
                    oversent_reply,
                    check
                ])

        # ==========================
        # RESULT
        # ==========================
        if results:
            result_df = pd.DataFrame(results, columns=[
                "MODEL",
                "PART NO",
                "QTY NEEDED",
                "OLD OVERSENT",
                "QTY SENT",
                "OVERSENT CALC",
                "OVERSENT REPLY",
                "CHECK"
            ])

            st.success("✅ Completed")

            st.dataframe(result_df)

            # DOWNLOAD
            output = BytesIO()
            result_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "📥 Download Excel",
                data=output,
                file_name="Oversent_Check.xlsx"
            )

        else:
            st.warning("⚠️ No results found")

else:
    st.info("📌 Upload files to start")
