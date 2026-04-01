import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Calculation Tool (New Logic)")

# =========================
# 1. UPLOAD FILES
# =========================
files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# =========================
# 2. MODEL + ODF
# =========================
model = st.text_input("📌 Enter Model Name")
odf = st.text_input("📌 Enter ODF").strip().upper()

# =========================
# 3. USER INPUT TABLE
# =========================
st.subheader("✍️ Enter Data Manually")

data = []

num_rows = st.number_input("Number of items", min_value=1, value=1)

for i in range(num_rows):

    st.markdown(f"### Item {i+1}")

    pn = st.text_input(f"PN {i+1}", key=f"pn_{i}")
    qty_needed = st.number_input(f"QTY NEEDED IN THIS LOT {i+1}", key=f"qty_{i}")
    qty_sent = st.number_input(f"QTY SENT IN THIS LOT {i+1}", key=f"sent_{i}")
    oversent_reply = st.number_input(f"OVERSENT REPLY {i+1}", key=f"reply_{i}")

    # =========================
    # SELECT FILE
    # =========================
    selected_file = st.selectbox(
        f"Select file for PN {i+1}",
        [f.name for f in files] if files else [],
        key=f"file_{i}"
    )

    data.append({
        "pn": pn.strip().upper(),
        "qty_needed": qty_needed,
        "qty_sent": qty_sent,
        "oversent_reply": oversent_reply,
        "file": selected_file
    })

# =========================
# PROCESS BUTTON
# =========================
if st.button("▶️ Calculate Oversent"):

    if not files:
        st.error("❌ Please upload files")
        st.stop()

    results = []

    # =========================
    # LOAD FILES
    # =========================
    frs_dict = {}

    for file in files:
        df = pd.read_excel(file)

        df.columns = df.columns.str.strip().str.upper()

        # Clean data
        for col in ["PART N", "ODF"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.upper()

        frs_dict[file.name] = df

    # =========================
    # PROCESS EACH ITEM
    # =========================
    for item in data:

        pn = item["pn"]
        qty_needed = item["qty_needed"]
        qty_sent = item["qty_sent"]
        oversent_reply = item["oversent_reply"]
        file_name = item["file"]

        if file_name not in frs_dict:
            results.append([pn, "FILE NOT FOUND"])
            continue

        df = frs_dict[file_name]

        if "PART N" not in df.columns or "ODF" not in df.columns:
            results.append([pn, "MISSING COLUMNS"])
            continue

        # =========================
        # FIND ALL MATCHING PN
        # =========================
        matches = df[df["PART N"] == pn]

        if matches.empty:
            results.append([pn, "PN NOT FOUND"])
            continue

        # =========================
        # FILTER BY ODF
        # =========================
        matches_odf = matches[matches["ODF"] == odf]

        if matches_odf.empty:
            results.append([pn, "ODF NOT FOUND"])
            continue

        idx = matches_odf.index[0]

        # =========================
        # GET PREVIOUS OVERSENT
        # =========================
        if idx > 0 and "OVERSENT QTY" in df.columns:
            last_oversent = df.iloc[idx - 1]["OVERSENT QTY"]
        else:
            last_oversent = 0

        # =========================
        # CALCUL
        # =========================
        calc_oversent = (last_oversent - qty_needed) + qty_sent

        # =========================
        # CHECK
        # =========================
        status = "OK" if calc_oversent == oversent_reply else "NON CONFORME"

        results.append([
            pn,
            last_oversent,
            qty_needed,
            qty_sent,
            calc_oversent,
            oversent_reply,
            status
        ])

    # =========================
    # DISPLAY RESULT
    # =========================
    if results:

        result_df = pd.DataFrame(results, columns=[
            "PART NO",
            "LAST OVERSENT",
            "QTY NEEDED",
            "QTY SENT",
            "CALCULATED OVERSENT",
            "OVERSENT REPLY",
            "STATUS"
        ])

        st.success("✅ Calculation Done")
        st.dataframe(result_df)

        # =========================
        # DOWNLOAD
        # =========================
        output = BytesIO()
        result_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            "📥 Download Result",
            data=output,
            file_name="oversent_result.xlsx"
        )

    else:
        st.warning("⚠️ No results found")
