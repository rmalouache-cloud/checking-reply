import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 FRS Oversent Verification Tool")

# ==============================
# UPLOAD FILES
# ==============================
main_file = st.file_uploader("📂 Upload Main Report", type=["xlsx"])
frs_files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# ==============================
# FUNCTION
# ==============================
def extract_frs(text):
    match = re.search(r'"(.*?)"', str(text))
    return match.group(1) if match else None

# ==============================
# PROCESS
# ==============================
if main_file and frs_files:

    # Charger main file
    main_sheets = pd.read_excel(main_file, sheet_name=None)

    # Charger fichiers FRS
    frs_dict = {}
    for file in frs_files:
        name = file.name.replace(".xlsx", "")
        df = pd.concat(pd.read_excel(file, sheet_name=None).values())
        df = df.fillna(0)
        frs_dict[name] = df

    results = []

    # ==============================
    # ANALYSE
    # ==============================
    for sheet_name, df in main_sheets.items():

        df = df.fillna("")

        for _, row in df.iterrows():

            part_no = row.get("PART NO", "")
            odf = row.get("ODF", "")
            moka = row.get("moka reply", "")
            main_oversent = row.get("Oversent qty", 0)

            frs_name = extract_frs(moka)

            if not frs_name or frs_name not in frs_dict:
                continue

            frs_df = frs_dict[frs_name]

            match = frs_df[
                (frs_df["PART NO."] == part_no) &
                (frs_df["ODF"] == odf)
            ]

            if match.empty:
                results.append([part_no, odf, "NOT FOUND"])
                continue

            r = match.iloc[0]

            qty_needed = r["QTY NEEDED IN THIS LOT"]
            qty_sent = r["QTY SENT IN THIS LOT"]
            old_oversent = r["QTY OVERSENT IN LAST TIME - QTY NEEDED IN THIS LOT"]
            frs_oversent = r["OVERSENT QTY"]

            # 🔥 CALCUL
            calc = old_oversent + qty_sent - qty_needed

            # RESULT
            if calc == frs_oversent:
                status = "OK"
            else:
                status = "ERROR"

            if calc < 0:
                status = "SHORTAGE"

            results.append([
                part_no,
                odf,
                calc,
                frs_oversent,
                main_oversent,
                status
            ])

    # ==============================
    # RESULT TABLE
    # ==============================
    result_df = pd.DataFrame(results, columns=[
        "PART NO",
        "ODF",
        "CALCULATED",
        "FRS OVERSENT",
        "MAIN OVERSENT",
        "STATUS"
    ])

    st.subheader("📊 Résultat de vérification")
    st.dataframe(result_df)

    # ==============================
    # DOWNLOAD
    # ==============================
    output = BytesIO()
    result_df.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        "📥 Télécharger le rapport",
        data=output,
        file_name="FRS_CHECK.xlsx"
    )

else:
    st.info("📌 Upload main file + FRS files")
