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
# GLOBAL INPUT
# =========================
model = st.text_input("📌 Model")
odf = st.text_input("📌 ODF").strip().upper()

# =========================
# TABLE SIZE
# =========================
st.subheader("🧾 Create Your Table")

num_rows = st.number_input("Number of articles", min_value=1, value=5)

# =========================
# TABLE (SANS FILE NAME)
# =========================
df_input = pd.DataFrame({
    "PART NO": [""] * num_rows,
    "QTY NEEDED": [0] * num_rows,
    "QTY SENT": [0] * num_rows,
    "OVERSENT REPLY": [0] * num_rows
})

edited_df = st.data_editor(df_input, use_container_width=True)

# =========================
# FILE SELECT PAR LIGNE
# =========================
st.subheader("📂 Select file for each row")

file_selection = []

file_names = [f.name for f in files] if files else []

for i in range(len(edited_df)):
    selected = st.selectbox(
        f"File for row {i+1}",
        file_names,
        key=f"file_{i}"
    )
    file_selection.append(selected)

# =========================
# PROCESS
# =========================
if st.button("▶️ Calculate"):

    if not files:
        st.error("❌ Upload files first")
        st.stop()

    # LOAD FILES
    frs_dict = {}

    for file in files:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.upper()

        part_col = next((c for c in df.columns if "PART" in c), None)
        odf_col = next((c for c in df.columns if "ODF" in c), None)
        oversent_col = next((c for c in df.columns if "OVERSENT" in c), None)

        if part_col:
            df[part_col] = df[part_col].astype(str).str.strip().str.upper()

        if odf_col:
            df[odf_col] = df[odf_col].astype(str).str.strip().str.upper()

        df = df.reset_index(drop=True)

        frs_dict[file.name] = {
            "df": df,
            "part_col": part_col,
            "odf_col": odf_col,
            "oversent_col": oversent_col
        }

    results = []

    # =========================
    # LOOP
    # =========================
    for i, row in edited_df.iterrows():

        pn = str(row["PART NO"]).strip().upper()
        qty_needed = row["QTY NEEDED"]
        qty_sent = row["QTY SENT"]
        oversent_reply = row["OVERSENT REPLY"]
        file_name = file_selection[i]

        if pn == "":
            continue

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

        same_pn = df[df[part_col] == pn]

        if same_pn.empty:
            result["STATUS"] = "PN NOT FOUND"
            results.append(result)
            continue

        matches_odf = same_pn[same_pn[odf_col] == odf]

        if matches_odf.empty:
            result["STATUS"] = "ODF NOT FOUND"
            results.append(result)
            continue

        current_idx = matches_odf.index[0]

        previous_rows = same_pn[same_pn.index < current_idx]

        if not previous_rows.empty:
            last_row = previous_rows.iloc[-1]
            last_oversent = last_row[oversent_col]
        else:
            last_oversent = 0

        result["LAST OVERSENT"] = last_oversent

        calc = (last_oversent - qty_needed) + qty_sent
        result["CALCULATED OVERSENT"] = calc

        result["STATUS"] = "OK" if calc == oversent_reply else "NON CONFORME"

        results.append(result)

    # =========================
    # RESULT
    # =========================
    df_result = pd.DataFrame(results)

    st.success("✅ Calculation Done")

    # 🎨 COLOR
    def color_status(val):
        if val == "OK":
            return "background-color: #c6f7c6"
        elif val == "NON CONFORME":
            return "background-color: #f7c6c6"
        return ""

    styled_df = df_result.style.map(color_status, subset=["STATUS"])

    st.dataframe(styled_df, use_container_width=True)

    # DOWNLOAD
    output = BytesIO()
    df_result.to_excel(output, index=False)
    output.seek(0)

    st.download_button(
        "📥 Download Result",
        data=output,
        file_name="Oversent_Result.xlsx"
    )
