import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Oversent Verification Tool (Bulk Mode)")

# =========================
# UPLOAD FILES
# =========================
files = st.file_uploader("📂 Upload FRS Files", type=["xlsx"], accept_multiple_files=True)

# =========================
# INPUT GLOBAL
# =========================
model = st.text_input("📌 Model")
odf = st.text_input("📌 ODF").strip().upper()

# =========================
# BULK INPUT
# =========================
st.subheader("✍️ Paste your data (one line per PN)")

st.markdown("""
Format (séparé par ; ou tab) :

PN ; QTY NEEDED ; QTY SENT ; OVERSENT REPLY ; FILE NAME  

Exemple :
