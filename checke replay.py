import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 Oversent Verification Tool")

# ==============================
# 1. UPLOAD FILES
# ==============================
main_file = st.file_uploader("📂 Upload Main File", type=["xlsx"])
frs_files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# ==============================
# EXTRACTION NOM FICHIER FRS
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
# MAIN PROCESS
# ==============================
if main_file and frs_files:

    # 2. LIRE LES FEUILLES (MODÈLES)
    main_sheets = pd.read_excel(main_file, sheet_name=None)

    st.subheader("📝 Saisir ODF pour chaque modèle")

    # 3. SAISIR ODF
    odf_inputs = {}
    for sheet in main_sheets.keys():
        odf_inputs[sheet] = st.text_input(f"ODF for {sheet}")

    # BOUTON LANCEMENT
    if st.button("▶️ Start Verification"):

        st.info("⏳ Traitement en cours...")

        # Charger fichiers FRS
        frs_dict = {}

        for file in frs_files:
            name = file.name.replace(".xlsx", "")
            df = pd.concat(pd.read_excel(file, sheet_name=None).values())
            df = df.fillna(0)

            df["PART NO."] = df["PART NO."].astype(str).str.strip().str.upper()
            df["ODF"] = df["ODF"].astype(str).str.strip().str.upper()

            frs_dict[name] = df

        results = []

        # ==============================
        # TRAITEMENT PAR FEUILLE
        # ==============================
        for sheet_name, df in main_sheets.items():

            df = df.fillna("")
            odf = odf_inputs.get(sheet_name, "").strip().upper()

            if odf == "":
                st.warning(f"⚠️ ODF manquant pour {sheet_name}")
                continue

            # 4. FILTRE REMARKS
            df = df[df["Remarks"].str.contains("Missing|Shortage", case=False, na=False)]

            # 5. FILTRE MOKA REPLY
            df = df[df["moka reply"].str.contains(
                "it's enough for your production",
                case=False,
                na=False
            )]

            for _, row in df.iterrows():

                # 6. PART NO
                part_no = str(row.get("PART NO", "")).strip().upper()

                # 7. QTY NEEDED (colonne E)
                qty_needed = row.iloc[4]

                # 9. QTY SENT
                qty_sent = row.get("Packing list qty", 0)

                # 11. OVERSENT REPLY (colonne I)
                oversent_reply = row.iloc[8]

                moka = row.get("moka reply", "")
                frs_name = extract_frs(moka)

                if not frs_name or frs_name not in frs_dict:
                    continue

                frs_df = frs_dict[frs_name]

                # 8. CHERCHER PN + ODF
                match = frs_df[
                    (frs_df["PART NO."] == part_no) &
                    (frs_df["ODF"] == odf)
                ]

                if match.empty:
                    results.append([sheet_name, part_no, "NOT FOUND"])
                    continue

                idx = match.index[0]

                # 🔥 OLD OVERSENT (ligne précédente)
                if idx > 0:
                    old_oversent = frs_df.iloc[idx - 1]["OVERSENT QTY"]
                else:
                    old_oversent = 0

                # 10. CALCUL
                calc = (old_oversent - qty_needed) + qty_sent

                # 12. COMPARAISON
                if calc == oversent_reply:
                    check = "OK"
                else:
                    check = "NON CONFORME"

                # RESULT
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

        # ==============================
        # TABLE FINAL
        # ==============================
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

        st.success("✅ Vérification terminée")

        st.dataframe(result_df)

        # DOWNLOAD
        output = BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            "📥 Télécharger le résultat",
            data=output,
            file_name="Oversent_Check.xlsx"
        )

else:
    st.info("📌 Upload fichiers pour commencer")
