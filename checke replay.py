import streamlit as st
import pandas as pd
from io import BytesIO

# ==============================
# CONFIG
# ==============================
st.set_page_config(page_title="BOM vs Packing Checker", layout="wide")

st.title("📊 BOM vs Packing Comparator")

# ==============================
# UPLOAD FILES
# ==============================
packing_file = st.file_uploader("📂 Upload Packing Excel file", type=["xlsx"])
bom_file = st.file_uploader("📂 Upload BOM Excel file", type=["xlsx"])

# ==============================
# PROCESS
# ==============================
if packing_file is not None and bom_file is not None:

    try:
        # Lire les fichiers
        packing_sheets = pd.read_excel(packing_file, sheet_name=None)
        bom_sheets = pd.read_excel(bom_file, sheet_name=None)

        # Prendre la première feuille
        packing_df = list(packing_sheets.values())[0]
        bom_df = list(bom_sheets.values())[0]

        st.success("✅ Fichiers chargés avec succès")

        # ==============================
        # CHOIX DES COLONNES
        # ==============================
        st.subheader("⚙️ Configuration des colonnes")

        packing_col = st.selectbox(
            "Choisir la colonne PN (Packing)",
            packing_df.columns
        )

        bom_col = st.selectbox(
            "Choisir la colonne PN (BOM)",
            bom_df.columns
        )

        # ==============================
        # COMPARAISON
        # ==============================
        packing_pn = set(packing_df[packing_col].astype(str))
        bom_pn = set(bom_df[bom_col].astype(str))

        missing_in_packing = bom_pn - packing_pn
        missing_in_bom = packing_pn - bom_pn

        # ==============================
        # AFFICHAGE
        # ==============================
        st.subheader("📊 Résultats")

        col1, col2 = st.columns(2)

        with col1:
            st.write("❌ PN manquants dans Packing")
            st.write(len(missing_in_packing))
            st.dataframe(pd.DataFrame(list(missing_in_packing), columns=["PN"]))

        with col2:
            st.write("❌ PN manquants dans BOM")
            st.write(len(missing_in_bom))
            st.dataframe(pd.DataFrame(list(missing_in_bom), columns=["PN"]))

        # ==============================
        # EXPORT EXCEL
        # ==============================
        output = BytesIO()

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(list(missing_in_packing), columns=["PN"]).to_excel(
                writer, sheet_name="Missing_in_Packing", index=False
            )
            pd.DataFrame(list(missing_in_bom), columns=["PN"]).to_excel(
                writer, sheet_name="Missing_in_BOM", index=False
            )

        output.seek(0)

        st.download_button(
            label="📥 Télécharger le rapport Excel",
            data=output,
            file_name="BOM_vs_Packing_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erreur : {e}")
else:
    st.info("📌 Veuillez uploader les deux fichiers pour lancer la comparaison")