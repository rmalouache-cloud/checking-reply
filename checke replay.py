import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Calculation Tool (Stable Version)")

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
# 3. USER INPUT
# =========================
st.subheader("✍️ Enter Data")

num_rows = st.number_input("Number of items", min_value=1, value=1)

data = []

for i in range(num_rows):

    st.markdown(f"### Item {i+1}")

    pn = st.text_input(f"PN {i+1}", key=f"pn_{i}")
    qty_needed = st.number_input(f"QTY NEEDED {i+1}", key=f"qty_{i}")
    qty_sent = st.number_input(f"QTY SENT {i+1}", key=f"sent_{i}")
    oversent_reply = st.number_input(f"OVERSENT REPLY {i+1}", key=f"reply_{i}")

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
# PROCESS
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

        # Nettoyage
        if "PART N" in df.columns:
            df["PART N"] = df["PART N"].astype(str).str.strip().str.upper()
        if "ODF" in df.columns:
            df["ODF"] = df["ODF"].astype(str).str.strip().str.upper()

        frs_dict[file.name] = df

    # =========================
    # PROCESS ITEMS
    # =========================
    for item in data:

        pn = item["pn"]
        qty_needed = item["qty_needed"]
        qty_sent = item["qty_sent"]
        oversent_reply = item["oversent_reply"]
        file_name = item["file"]

        result = {
            "MODEL": model,
            "PART NO": pn,
            "LAST OVERSENT": None,
            "QTY NEEDED": qty_needed,
            "QTY SENT": qty_sent,
            "CALCULATED OVERSENT": None,
            "OVERSENT REPLY": oversent_reply,
            "STATUS": ""
        }

        # =========================
        # VALIDATIONS
        # =========================
        if file_name not in frs_dict:
            result["STATUS"] = "FILE NOT FOUND"
            results.append(result)
            continue

        df = frs_dict[file_name]

        if "PART N" not in df.columns or "ODF" not in df.columns:
            result["STATUS"] = "MISSING COLUMNS"
            results.append(result)
            continue

        # =========================
        # FIND PN
        # =========================
        matches = df[df["PART N"] == pn]

        if matches.empty:
            result["STATUS"] = "PN NOT FOUND"
            results.append(result)
            continue

        # =========================
        # FILTER ODF
        # =========================
        matches_odf = matches[matches["ODF"] == odf]

        if matches_odf.empty:
            result["STATUS"] = "ODF NOT FOUND"
            results.append(result)
            continue

        idx = matches_odf.index[0]

        # =========================
        # GET LAST OVERSENT
        # =========================
        if "OVERSENT QTY" in df.columns and idx > 0:
            last_oversent = df.iloc[idx - 1]["OVERSENT QTY"]
        else:
            last_oversent = 0

        result["LAST OVERSENT"] = last_oversent

        # =========================
        # CALCUL
        # =========================
        calc = (last_oversent - qty_needed) + qty_sent
        result["CALCULATED OVERSENT"] = calc

        # =========================
        # CHECK
        # =========================
        if calc == oversent_reply:
            result["STATUS"] = "OK"
        else:
            result["STATUS"] = "NON CONFORME"

        results.append(result)

    # =========================
    # RESULT TABLE
    # =========================
    result_df = pd.DataFrame(results)

    st.success("✅ Calculation Completed")
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
