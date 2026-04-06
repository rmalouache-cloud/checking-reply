import streamlit as st
import pandas as pd

# =========================
# 🧠 SESSION STATE INIT
# =========================

if "packing_type" not in st.session_state:
    st.session_state.packing_type = "Panel"

if "model" not in st.session_state:
    st.session_state.model = ""

if "odf" not in st.session_state:
    st.session_state.odf = ""

# =========================
# 📦 DASHBOARD TITLE
# =========================

st.markdown("### 📦 Container Filling Industrial Dashboard")

# =========================
# 🎯 INPUT FIELDS
# =========================

packing_type = st.selectbox(
    "Type of Packing List",
    ["Panel", "SP", "SP/MainBoard", "OC"],
    key="packing_type"
)

model = st.text_input(
    "Model (ex: Mini LED)",
    key="model"
)

odf = st.text_input(
    "ODF (ex: IDL2500)",
    key="odf"
)

# =========================
# 🏷️ DYNAMIC TITLE
# =========================

if st.session_state.model and st.session_state.odf:
    title = f"Container Filling Industrial Dashboard of {st.session_state.packing_type} of {st.session_state.model}__{st.session_state.odf}"
else:
    title = "Container Filling Industrial Dashboard"

st.title(title)

st.markdown("---")

# =========================
# 📂 UPLOAD EXCEL FILE
# =========================

uploaded_file = st.file_uploader(
    "Upload Packing Excel file",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        st.success("File uploaded successfully ✅")

        st.dataframe(df)

        # 👉 PLACE FOR YOUR MAIN LOGIC
        st.info("Main logic will be executed here")

    except Exception as e:
        st.error(f"Error reading file: {e}")

# =========================
# 🔧 YOUR LOGIC HERE
# =========================
# Tu peux ajouter ton code existant ici sans modification majeure
