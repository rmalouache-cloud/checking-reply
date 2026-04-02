import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Verification Tool (Fast Mode)")

# =========================
# UPLOAD FILES
# =========================
files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

file_names = [f.name for f in files] if files else []

# =========================
# GLOBAL INPUT
# =========================
model = st.text_input("📌 Model")
odf = st.text_input("📌 ODF").strip().upper()

selected_file = st.selectbox("📂 Select FRS File", file_names)

# =========================
# TABLE SIZE
# =========================
st.subheader("🧾 Create Your Table")

num_rows = st.number_input("Number of articles", min_value=1, value=5)

# =========================
# TABLE
# =========================
df_input = pd.DataFrame({
    "PART NO": [""] * num_rows,
    "QTY NEEDED": [0] * num_rows,
    "QTY SENT": [0] * num_rows,
    "OVERSENT REPLY": [0] * num_rows
})

edited_df = st.data_editor(df_input, use_container_width=True)

# =========================
# PROCESS
# =========================
if st.button("▶️ Calculate"):

    if not files:
        st.error("❌ Upload file first")
        st.stop()

    if selected_file == "":
        st.error("❌ Select a file")
        st.stop()

    # =========================
    # LOAD FILE
    # =========================
    file_obj = next(f for f in files if f.name == selected_file)

    df = pd.read_excel(file_obj)
    df.columns = df.columns.str.strip().str.upper()

    # AUTO DETECT
    part_col = next((c for c in df.columns if "PART" in c), None)
    odf_col = next((c for c in df.columns if "ODF" in c), None)

    # ✅ FIX: force correct column (OVERSENT QTY)
    oversent_col = next((c for c in df.columns if "OVERSENT QTY" in c), None)

    if oversent_col is None:
        st.error("❌ Column 'OVERSENT QTY' not found in the file")
        st.stop()

    # CLEAN
    df[part_col] = df[part_col].astype(str).str.strip().str.upper()
    df[odf_col] = df[odf_col].astype(str).str.strip().str.upper()

    df = df.reset_index(drop=True)

    results = []

    # =========================
    # LOOP
    # =========================
    for _, row in edited_df.iterrows():

        pn = str(row["PART NO"]).strip().upper()
        qty_needed = row["QTY NEEDED"]
        qty_sent = row["QTY SENT"]
        oversent_reply = row["OVERSENT REPLY"]

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

        # 🔥 Take previous rows (before selected ODF)
        previous_rows = same_pn[same_pn.index < current_idx]

        if not previous_rows.empty:
            last_row = previous_rows.iloc[-1]

            # ✅ USE CORRECT COLUMN (OVERSENT QTY - column K)
            last_oversent = last_row[oversent_col]
        else:
            last_oversent = 0

        result["LAST OVERSENT"] = last_oversent

        # CALCULATION
        calc = (last_oversent - qty_needed) + qty_sent
        result["CALCULATED OVERSENT"] = calc

        # CHECK
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
