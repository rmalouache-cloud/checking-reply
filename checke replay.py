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
# EXTRACTION FRS NAME
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

    # Lire toutes les feuilles
    main_sheets = pd.read_excel(main_file, sheet_name=None)

    st.subheader("📝 Saisir ODF par modèle")

    # Saisie ODF
    odf_inputs = {}
    for sheet in main_sheets.keys():
        odf_inputs[sheet] = st.text_input(f"ODF for {sheet}")

    if st.button("▶️ Start Verification"):

        st.info("⏳ Traitement en cours...")

        # ==========================
        # Charger FRS
        # ==========================
        frs_dict = {}

        for file in frs_files:
            name = file.name.replace(".xlsx", "")
            df = pd.concat(pd.read_excel(file, sheet_name=None).values())

            df.columns = df.columns.str.strip().str.upper()

            df["PART NO."] = df["PART NO."].astype(str).str.strip().str.upper()
            df["ODF"] = df["ODF"].astype(str).str.strip().str.upper()

            frs_dict[name] = df

        results = []

        # ==========================
        # TRAITEMENT PAR FEUILLE
        # ==========================
        for sheet_name, df in main_sheets.items():

            st.markdown(f"---\n### 📄 Sheet: {sheet_name}")

            df.columns = df.columns.str.strip().str.upper()
            df = df.fillna("")

            # DEBUG colonnes
            st.write("📌 Colonnes :", df.columns.tolist())

            odf = odf_inputs.get(sheet_name, "").strip().upper()

            if odf == "":
                st.warning(f"⚠️ ODF manquant pour {sheet_name}")
                continue

            # ==========================
            # FILTRE 1 : REMARKS
            # ==========================
            if "REMARKS" not in df.columns:
                st.error("❌ Colonne REMARKS manquante")
                continue

            df1 = df[df["REMARKS"].astype(str).str.upper().str.contains("MISSING|SHORTAGE", na=False)]

            st.write("🔎 Après filtre REMARKS :", df1.shape)

            # ==========================
            # FILTRE 2 : MOKA REPLY
            # ==========================
            if "MOKA REPLY" not in df.columns:
                st.error("❌ Colonne MOKA REPLY manquante")
                continue

            df2 = df1[df1["MOKA REPLY"].astype(str).str.lower().str.contains("enough", na=False)]

            st.write("🔎 Après filtre MOKA :", df2.shape)

            if df2.empty:
                st.warning("⚠️ Aucune ligne après filtrage")
                continue

            # ==========================
            # TRAITEMENT LIGNES
            # ==========================
            for _, row in df2.iterrows():

                part_no = str(row.get("PART NO", "")).strip().upper()

                qty_needed = row.iloc[4] if len(row) > 4 else 0
                qty_sent = row.get("PACKING LIST QTY", 0)
                oversent_reply = row.iloc[8] if len(row) > 8 else 0

                moka = row.get("MOKA REPLY", "")
                frs_name = extract_frs(moka)

                if not frs_name or frs_name not in frs_dict:
                    continue

                frs_df = frs_dict[frs_name]

                match = frs_df[
                    (frs_df["PART NO."] == part_no) &
                    (frs_df["ODF"] == odf)
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

                # CALCUL
                calc = (old_oversent - qty_needed) + qty_sent

                # CHECK
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
        # RESULTATS
        # ==========================
        if results:
            result_df = pd.DataFrame(results, columns=[
                "MODEL",
                "PART NO",
                "QTY NEEDED",
                "OLD OVERSENT",
                "QTY SENT",
                "OVERSENT CALCULE",
                "OVERSENT REPLY",
                "CHECK"
            ])

            st.success("✅ Traitement terminé")

            st.dataframe(result_df)

            # DOWNLOAD
            output = BytesIO()
            result_df.to_excel(output, index=False)
            output.seek(0)

            st.download_button(
                "📥 Télécharger Excel",
                data=output,
                file_name="Oversent_Check.xlsx"
            )
        else:
            st.warning("⚠️ Aucun résultat trouvé")

else:
    st.info("📌 Upload fichiers pour commencer")
