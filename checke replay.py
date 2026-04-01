import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Verification Tool")

# ==============================
# UPLOAD FILES
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
# START
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

            # Nettoyage
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
                st.error("❌ REMARKS missing")
                continue

            df1 = df[df["REMARKS"].astype(str).str.upper().str.contains("MISSING|SHORTAGE", na=False)]
            st.write("🔎 After REMARKS filter:", df1.shape)

            # ==========================
            # FILTRE MOKA
            # ==========================
            if "MOKA REPLY" not in df.columns:
                st.error("❌ MOKA REPLY missing")
                continue

            df2 = df1[df1["MOKA REPLY"].astype(str).str.lower().str.contains("enough", na=False)]
            st.write("🔎 After MOKA filter:", df2.shape)

            if df2.empty:
                st.warning("⚠️ No data after filters")
                continue

            # ==========================
            # TRAITEMENT
            # ==========================
            for _, row in df2.iterrows():

                # PART N
                part_no = str(row.get("PART N", "")).strip().upper()

                # QTY NEEDED (colonne dynamique)
                qty_col = [col for col in df.columns if "QTY FOR" in col]

                if not qty_col:
                    continue

                qty_needed = row[qty_col[0]]

                # PACKING LIST
                qty_sent = row.get("PACKING LIST QTY", 0)

                # OVERSENT REPLY
                oversent_reply = row.get("OVERSENT QTY", 0)

                # MOKA REPLY
                moka = row.get("MOKA REPLY", "")

                # EXTRACTION FRS FILE
                frs_name = extract_frs(moka)

                if not frs_name or frs_name not in frs_dict:
                    continue

                frs_df = frs_dict[frs_name]

                # ==========================
                # NORMALISATION MATCH
                # ==========================
                frs_df["PART NO."] = frs_df["PART NO."].astype(str).str.strip().str.upper()
                frs_df["ODF"] = frs_df["ODF"].astype(str).str.strip().str.upper()

                part_no_clean = part_no.strip().upper()
                odf_clean = odf.strip().upper()

                # DEBUG
                st.write("🔍 Searching:", part_no_clean, odf_clean)

                # ==========================
                # MATCH
                # ==========================
                match = frs_df[
                    (frs_df["PART NO."] == part_no_clean) &
                    (frs_df["ODF"] == odf_clean)
                ]

                if match.empty:
                    results.append([sheet_name, part_no, "NOT FOUND"])
                    continue

                idx = match.index[0]

                # OLD OVERSENT
                if idx > 0:
                    old_oversent = frs_df.iloc[idx - 1].get("OVERSENT QTY", 0)
                else:
                    old_oversent = 0

                # ==========================
                # CALCUL
                # ==========================
                calc = (old_oversent - qty_needed) + qty_sent

                # ==========================
                # CHECK
                # ==========================
                check = "OK" if calc == oversent_reply else "NON CONFORME"

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

            st.success("✅ Done")

            st.dataframe(result_df)

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
