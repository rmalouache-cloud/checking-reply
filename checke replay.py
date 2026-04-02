import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Verification Tool")

# =========================
# UPLOAD FILES
# =========================
files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# =========================
# INPUTS
# =========================
model = st.text_input("📌 Model")
odf = st.text_input("📌 ODF").strip().upper()

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
if st.button("▶️ Calculate"):

    if not files:
        st.error("❌ Upload files first")
        st.stop()

    results = []
    frs_dict = {}

    # =========================
    # LOAD FRS FILES
    # =========================
    for file in files:
        df = pd.read_excel(file)

        # Nettoyage colonnes
        df.columns = df.columns.str.strip().str.upper()

        # Détection automatique
        part_col = next((c for c in df.columns if "PART" in c), None)
        odf_col = next((c for c in df.columns if "ODF" in c), None)
        oversent_col = next((c for c in df.columns if "OVERSENT" in c), None)

        # Nettoyage données
        if part_col:
            df[part_col] = df[part_col].astype(str).str.strip().str.upper()

        if odf_col:
            df[odf_col] = df[odf_col].astype(str).str.strip().str.upper()

        # Reset index (important)
        df = df.reset_index(drop=True)

        frs_dict[file.name] = {
            "df": df,
            "part_col": part_col,
            "odf_col": odf_col,
            "oversent_col": oversent_col
        }

    # =========================
    # PROCESS EACH ITEM
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

        if file_name not in frs_dict:
            result["STATUS"] = "FILE NOT FOUND"
            results.append(result)
            continue

        df_info = frs_dict[file_name]
        df = df_info["df"]
        part_col = df_info["part_col"]
        odf_col = df_info["odf_col"]
        oversent_col = df_info["oversent_col"]

        if not part_col or not odf_col or not oversent_col:
            result["STATUS"] = "COLUMN ERROR"
            results.append(result)
            continue

        # =========================
        # FIND ALL SAME PN
        # =========================
        same_pn = df[df[part_col] == pn]

        if same_pn.empty:
            result["STATUS"] = "PN NOT FOUND"
            results.append(result)
            continue

        # =========================
        # FIND CURRENT ODF
        # =========================
        matches_odf = same_pn[same_pn[odf_col] == odf]

        if matches_odf.empty:
            result["STATUS"] = "ODF NOT FOUND"
            results.append(result)
            continue

        current_idx = matches_odf.index[0]

        # =========================
        # 🔥 CORRECT LOGIC FOR LAST OVERSENT
        # =========================
        previous_rows = same_pn[same_pn.index < current_idx]

        if not previous_rows.empty:
            last_row = previous_rows.iloc[-1]
            last_oversent = last_row[oversent_col]
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
    # RESULT
    # =========================
    df_result = pd.DataFrame(results)

    st.success("✅ Calculation Finished")
    st.dataframe(df_result)

    # DOWNLOAD
    output = BytesIO()
    df_result.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        "📥 Download Excel",
        data=output,
        file_name="Oversent_Result.xlsx"
    )
