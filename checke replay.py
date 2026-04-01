import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.title("📊 FRS Oversent Checker (with ODF Input)")

# ==============================
# UPLOAD
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
# MAIN PROCESS
# ==============================
if main_file and frs_files:

    # Lire fichier principal
    main_sheets = pd.read_excel(main_file, sheet_name=None)

    st.subheader("📝 Enter ODF for each Model")

    # ==============================
    # SAISIE ODF PAR FEUILLE
    # ==============================
    odf_inputs = {}

    for sheet in main_sheets.keys():
        odf_inputs[sheet] = st.text_input(f"ODF for {sheet}")

    # ==============================
    # Charger FRS files
    # ==============================
    frs_dict = {}

    for file in frs_files:
        name = file.name.replace(".xlsx", "")
        df = pd.concat(pd.read_excel(file, sheet_name=None).values())
        df = df.fillna(0)

        # Nettoyage
        df["PART NO."] = df["PART NO."].astype(str).str.strip().str.upper()
        df["ODF"] = df["ODF"].astype(str).str.strip().str.upper()

        frs_dict[name] = df

    results = []

    # ==============================
    # ANALYSE
    # ==============================
    for sheet_name, df in main_sheets.items():

        df = df.fillna("")
        odf = odf_inputs.get(sheet_name, "").strip().upper()

        if odf == "":
            continue

        for _, row in df.iterrows():

            part_no = str(row.get("PART NO", "")).strip().upper()
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
                results.append([sheet_name, part_no, odf, "NOT FOUND"])
                continue

            r = match.iloc[0]

            qty_needed = r["QTY NEEDED IN THIS LOT"]
            qty_sent = r["QTY SENT IN THIS LOT"]
            old_oversent = r["QTY OVERSENT IN LAST TIME - QTY NEEDED IN THIS LOT"]
            frs_oversent = r["OVERSENT QTY"]

            # 🔥 CALCUL
            calc = old_oversent + qty_sent - qty_needed

            # STATUS
            if calc == frs_oversent:
                status = "OK"
            else:
                status = "ERROR"

            if calc < 0:
                status = "SHORTAGE"

            results.append([
                sheet_name,
                part_no,
                odf,
                calc,
                frs_oversent,
                main_oversent,
                status
            ])

    # ==============================
    # RESULT
    # ==============================
    result_df = pd.DataFrame(results, columns=[
        "MODEL",
        "PART NO",
        "ODF",
        "CALCULATED",
        "FRS OVERSENT",
        "MAIN OVERSENT",
        "STATUS"
    ])

    st.subheader("📊 Result")
    st.dataframe(result_df)

    # ==============================
    # DOWNLOAD
    # ==============================
    output = BytesIO()
    result_df.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        "📥 Download Report",
        data=output,
        file_name="FRS_CHECK.xlsx"
    )

else:
    st.info("📌 Upload files to start")
