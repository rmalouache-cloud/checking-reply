import streamlit as st
import pandas as pd
import io

# Configuration
st.set_page_config(
    page_title="Vérification Fournisseur",
    page_icon="✓",
    layout="wide"
)

# ==================== CSS NEUTRE ====================
st.markdown("""
<style>
    /* Police */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    /* Fond neutre */
    .stApp {
        background: #ffffff;
    }
    
    /* En-tête */
    .header {
        padding: 2rem 0 1rem 0;
        margin-bottom: 2rem;
        text-align: center;
        border-bottom: 1px solid #e5e7eb;
    }
    
    .header h1 {
        color: #111827;
        margin: 0;
        font-size: 1.8rem;
        font-weight: 500;
    }
    
    .header p {
        color: #6b7280;
        margin: 0.5rem 0 0 0;
        font-size: 0.85rem;
    }
    
    /* Cartes */
    .card {
        background: #ffffff;
        border-radius: 8px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e5e7eb;
    }
    
    .card-title {
        font-size: 0.9rem;
        font-weight: 600;
        margin-bottom: 1rem;
        color: #374151;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Zone upload */
    .upload-area {
        border: 1px dashed #d1d5db;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        background: #fafafa;
    }
    
    /* Métriques */
    .metric-wrapper {
        background: #f9fafb;
        border-radius: 8px;
        padding: 1rem;
        text-align: center;
        border: 1px solid #e5e7eb;
    }
    
    .metric-number {
        font-size: 1.8rem;
        font-weight: 600;
        color: #111827;
    }
    
    .metric-label {
        font-size: 0.7rem;
        color: #6b7280;
        margin-top: 0.4rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    
    /* Info box */
    .info-box {
        background: #f9fafb;
        padding: 0.6rem 1rem;
        border-radius: 6px;
        font-size: 0.8rem;
        color: #4b5563;
        margin-bottom: 1rem;
        border-left: 2px solid #9ca3af;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1.5rem;
        margin-top: 2rem;
        color: #9ca3af;
        font-size: 0.7rem;
        border-top: 1px solid #e5e7eb;
    }
    
    /* Divider */
    .divider {
        height: 1px;
        background: #e5e7eb;
        margin: 1.5rem 0;
    }
    
    /* Inputs */
    .stTextInput > div > div > input {
        border-radius: 6px;
        border: 1px solid #e5e7eb;
        padding: 0.4rem 0.8rem;
        font-size: 0.85rem;
    }
    
    /* Bouton */
    .stButton > button {
        background: #374151;
        color: white;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
        border-radius: 6px;
        font-size: 0.85rem;
    }
    
    .stButton > button:hover {
        background: #1f2937;
    }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
st.markdown("""
<div class="header">
    <h1>Vérification des réponses fournisseur</h1>
    <p>Contrôle des quantités Oversent</p>
</div>
""", unsafe_allow_html=True)

# ==================== FONCTIONS ====================

def charger_feuilles_reply(uploaded_file):
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        return {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in xlsx.sheet_names}
    except Exception as e:
        st.error(f"Erreur: {e}")
        return None

def charger_stocks(uploaded_files):
    stocks = {}
    for f in uploaded_files:
        try:
            stocks[f.name] = pd.read_excel(f)
        except Exception as e:
            st.error(f"Erreur {f.name}: {e}")
    return stocks

def extraire_colonnes_reply(df):
    if len(df.columns) < 9:
        return None
    df = df.copy()
    df['Part_N'] = df.iloc[:, 0].astype(str).str.strip()
    df['Description'] = df.iloc[:, 1].astype(str).str.strip()
    df['Packing_qty'] = pd.to_numeric(df.iloc[:, 3], errors='coerce').fillna(0)
    df['Qty_for'] = pd.to_numeric(df.iloc[:, 4], errors='coerce').fillna(0)
    df['Remarks'] = df.iloc[:, 6].astype(str).str.strip()
    df['Moka_file'] = df.iloc[:, 7].astype(str).str.strip().replace('.xlsx', '', regex=True).replace('.xls', '', regex=True)
    df['Oversent_FRS'] = pd.to_numeric(df.iloc[:, 8], errors='coerce').fillna(0)
    return df

def get_oversent_stock(df_stock, part_n, idl):
    if len(df_stock.columns) < 11:
        raise ValueError("Colonnes insuffisantes")
    
    col_part = df_stock.columns[3]
    col_idl = df_stock.columns[0]
    
    mask = df_stock[col_part].astype(str).str.strip() == str(part_n).strip()
    df_filtered = df_stock[mask].reset_index(drop=True)
    
    if df_filtered.empty:
        raise ValueError("Part N non trouvé")
    
    mask_idl = df_filtered[col_idl].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask_idl].index
    
    if len(idx) == 0:
        raise ValueError("IDL non trouvé")
    
    pos = idx[0]
    if pos == 0:
        raise ValueError("IDL à la première ligne")
    
    val = df_filtered.iloc[pos - 1, 10]
    return float(val) if pd.notna(val) else 0.0

# ==================== INTERFACE ====================

# Upload
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card">
        <div class="card-title">Fichier reply</div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card">
        <div class="card-title">Fichiers stock</div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    stock_files = st.file_uploader("Fichiers stock", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

if reply_file and stock_files:
    
    with st.spinner("Chargement..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        st.markdown(f"""
        <div class="info-box">
            {len(dict_reply)} feuille(s) : {', '.join(dict_reply.keys())}
        </div>
        """, unsafe_allow_html=True)
        
        # IDL
        st.markdown("""
        <div class="card">
            <div class="card-title">IDL par modèle</div>
        """, unsafe_allow_html=True)
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                idl = st.text_input(f"{modele}", key=f"idl_{modele}", placeholder="IDL")
                if idl:
                    idl_par_modele[modele] = idl
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bouton
        col_btn1, col_btn2, col_btn3 = st.columns([1,2,1])
        with col_btn2:
            verifier = st.button("Vérifier", use_container_width=True)
        
        if verifier:
            if not idl_par_modele:
                st.warning("Saisissez au moins un IDL")
            else:
                resultats = []
                erreurs = []
                
                with st.spinner("Vérification en cours..."):
                    for modele, df_feuille in dict_reply.items():
                        if modele not in idl_par_modele:
                            continue
                        
                        df_std = extraire_colonnes_reply(df_feuille)
                        if df_std is None:
                            continue
                        
                        df_filtre = df_std[df_std['Remarks'].isin(['Missing', 'shortage'])]
                        
                        if df_filtre.empty:
                            continue
                        
                        idl = idl_par_modele[modele]
                        
                        for _, row in df_filtre.iterrows():
                            part_n = row['Part_N']
                            desc = row['Description'][:50]
                            qty_for = row['Qty_for']
                            packing_qty = row['Packing_qty']
                            oversent_frs = row['Oversent_FRS']
                            moka_file = row['Moka_file']
                            remarks = row['Remarks']
                            
                            stock_file = None
                            for fname in dict_stocks.keys():
                                if moka_file in fname.replace('.xlsx', '').replace('.xls', ''):
                                    stock_file = fname
                                    break
                            
                            if not stock_file:
                                erreurs.append(f"{part_n}: Fichier {moka_file} non trouvé")
                                resultats.append({
                                    'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                    'Remarks': remarks, 'Qty for': qty_for, 'Packing Qty': packing_qty,
                                    'Oversent FRS': oversent_frs, 'Oversent Calculé': None,
                                    'Écart': None, 'Status': '❌'
                                })
                                continue
                            
                            try:
                                oversent_stock = get_oversent_stock(dict_stocks[stock_file], part_n, idl)
                                oversent_calc = oversent_stock + packing_qty - qty_for
                                ecart = oversent_calc - oversent_frs
                                correct = abs(ecart) < 0.01
                                
                                resultats.append({
                                    'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                    'Remarks': remarks, 'IDL': idl,
                                    'Qty for': qty_for, 'Packing Qty': packing_qty,
                                    'Oversent Stock': oversent_stock,
                                    'Oversent FRS': oversent_frs,
                                    'Oversent Calculé': round(oversent_calc, 1),
                                    'Écart': round(ecart, 1),
                                    'Status': '✓' if correct else '✗'
                                })
                                
                            except Exception as e:
                                erreurs.append(f"{part_n}: {str(e)}")
                                resultats.append({
                                    'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                    'Remarks': remarks, 'Status': '!'
                                })
                
                if resultats:
                    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✓'])
                    incorrects = len(df_res[df_res['Status'] == '✗'])
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number">{total}</div>
                            <div class="metric-label">Total</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number">{corrects}</div>
                            <div class="metric-label">Corrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number">{incorrects}</div>
                            <div class="metric-label">Incorrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.dataframe(df_res, use_container_width=True, hide_index=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, sheet_name='Résultats', index=False)
                        if erreurs:
                            pd.DataFrame({'Erreurs': erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    st.download_button("📥 Télécharger", output.getvalue(), "verification.xlsx", use_container_width=True)
                    
                    if incorrects == 0:
                        st.balloons()
                        st.success("Toutes les vérifications sont correctes")

else:
    st.markdown("""
    <div class="info-box" style="text-align: center;">
        Chargez les fichiers pour commencer
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div class="footer">
    Formule : Oversent réel = Oversent stock (ligne précédente) + Packing list qty - Qty for
</div>
""", unsafe_allow_html=True)
