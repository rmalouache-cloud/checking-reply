import streamlit as st
import pandas as pd
import io

# Configuration
st.set_page_config(
    page_title="Vérification Fournisseur",
    page_icon="✨",
    layout="wide"
)

# ==================== CSS ÉLÉGANT ====================
st.markdown("""
<style>
    /* Police élégante */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Inter', sans-serif;
    }
    
    /* Fond subtil */
    .stApp {
        background: #fafbfc;
    }
    
    /* En-tête raffiné */
    .header {
        background: linear-gradient(120deg, #ffffff 0%, #f8f9fc 100%);
        padding: 2rem;
        border-radius: 24px;
        margin-bottom: 2rem;
        text-align: center;
        border: 1px solid rgba(0,0,0,0.05);
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }
    
    .header h1 {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0;
        font-size: 2.2rem;
        font-weight: 600;
    }
    
    .header p {
        color: #6c757d;
        margin: 0.5rem 0 0 0;
        font-size: 0.95rem;
    }
    
    /* Cartes élégantes */
    .card {
        background: white;
        border-radius: 20px;
        padding: 1.8rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03), 0 1px 2px rgba(0,0,0,0.05);
        border: 1px solid rgba(0,0,0,0.05);
        transition: all 0.2s ease;
    }
    
    .card:hover {
        box-shadow: 0 8px 24px rgba(0,0,0,0.08);
        border-color: rgba(0,0,0,0.08);
    }
    
    .card-title {
        font-size: 1.1rem;
        font-weight: 600;
        margin-bottom: 1.2rem;
        color: #1a1a2e;
        letter-spacing: -0.3px;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Zones d'upload */
    .upload-area {
        border: 2px dashed #e0e0e0;
        border-radius: 16px;
        padding: 1.5rem;
        text-align: center;
        background: #ffffff;
        transition: all 0.2s ease;
    }
    
    .upload-area:hover {
        border-color: #667eea;
        background: #f8f9ff;
    }
    
    /* Métriques élégantes */
    .metric-wrapper {
        background: white;
        border-radius: 16px;
        padding: 1.2rem;
        text-align: center;
        border: 1px solid rgba(0,0,0,0.05);
        transition: all 0.2s ease;
    }
    
    .metric-number {
        font-size: 2.2rem;
        font-weight: 700;
        line-height: 1;
    }
    
    .metric-label {
        font-size: 0.8rem;
        color: #6c757d;
        margin-top: 0.5rem;
        font-weight: 500;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }
    
    /* Inputs stylisés */
    .stTextInput > div > div > input {
        border-radius: 12px;
        border: 1.5px solid #e9ecef;
        padding: 0.6rem 1rem;
        font-size: 0.9rem;
        transition: all 0.2s;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102,126,234,0.1);
    }
    
    /* Bouton principal */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.7rem 2rem;
        font-weight: 500;
        border-radius: 40px;
        transition: all 0.3s ease;
        font-size: 0.95rem;
        letter-spacing: 0.3px;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(102,126,234,0.3);
    }
    
    /* Divider élégant */
    .divider {
        background: linear-gradient(90deg, transparent, #e0e0e0, transparent);
        height: 1px;
        margin: 2rem 0;
    }
    
    /* Badges de statut */
    .badge-success {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 500;
        display: inline-block;
    }
    
    /* Info box */
    .info-box {
        background: #f8f9fc;
        border-left: 3px solid #667eea;
        padding: 0.8rem 1rem;
        border-radius: 10px;
        font-size: 0.85rem;
        color: #495057;
        margin-bottom: 1rem;
    }
    
    /* Footer élégant */
    .footer {
        text-align: center;
        padding: 1.5rem;
        margin-top: 2rem;
        color: #868e96;
        font-size: 0.8rem;
        border-top: 1px solid #e9ecef;
    }
    
    /* Tableau stylisé */
    .stDataFrame {
        border-radius: 16px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
st.markdown("""
<div class="header">
    <h1>✨ Vérification Fournisseur</h1>
    <p>Contrôle intelligent des réponses et des stocks</p>
</div>
""", unsafe_allow_html=True)

# ==================== FONCTIONS ====================

def charger_feuilles_reply(uploaded_file):
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        return {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in xlsx.sheet_names}
    except Exception as e:
        st.error(f"❌ {e}")
        return None

def charger_stocks(uploaded_files):
    stocks = {}
    for f in uploaded_files:
        try:
            stocks[f.name] = pd.read_excel(f)
        except Exception as e:
            st.error(f"❌ {f.name}: {e}")
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
        <div class="card-title">
            <span>📄</span> Fichier Reply
        </div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span>🗂️</span> Fichiers Stock
        </div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    stock_files = st.file_uploader("Fichiers stock", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

if reply_file and stock_files:
    
    with st.spinner("Chargement en cours..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        # Info
        st.markdown(f"""
        <div class="info-box">
            📁 {len(dict_reply)} feuille(s) détectée(s) : <strong>{', '.join(dict_reply.keys())}</strong>
        </div>
        """, unsafe_allow_html=True)
        
        # IDL
        st.markdown("""
        <div class="card">
            <div class="card-title">
                <span>🔐</span> Configuration des IDL
            </div>
        """, unsafe_allow_html=True)
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                idl = st.text_input(f"📱 {modele}", key=f"idl_{modele}", placeholder="IDL")
                if idl:
                    idl_par_modele[modele] = idl
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bouton
        col_btn1, col_btn2, col_btn3 = st.columns([1,2,1])
        with col_btn2:
            verifier = st.button("▶️ LANCER LA VÉRIFICATION", use_container_width=True)
        
        if verifier:
            if not idl_par_modele:
                st.warning("⚠️ Veuillez saisir au moins un IDL")
            else:
                
                resultats = []
                erreurs = []
                
                with st.spinner("Analyse en cours..."):
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
                                    'Status': '✅' if correct else '❌'
                                })
                                
                            except Exception as e:
                                erreurs.append(f"{part_n}: {str(e)}")
                                resultats.append({
                                    'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                    'Remarks': remarks, 'Status': '⚠️'
                                })
                
                # Résultats
                if resultats:
                    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                    st.markdown("""
                    <div class="card">
                        <div class="card-title">
                            <span>📊</span> Résultats de l'analyse
                        </div>
                    """, unsafe_allow_html=True)
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✅'])
                    incorrects = len(df_res[df_res['Status'] == '❌'])
                    warnings = len(df_res[df_res['Status'] == '⚠️'])
                    
                    # Métriques
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number" style="color: #3b82f6;">{total}</div>
                            <div class="metric-label">Total vérifié</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number" style="color: #10b981;">{corrects}</div>
                            <div class="metric-label">✅ Corrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number" style="color: #ef4444;">{incorrects}</div>
                            <div class="metric-label">❌ Incorrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        taux = f"{corrects/total*100:.0f}" if total > 0 else "0"
                        st.markdown(f"""
                        <div class="metric-wrapper">
                            <div class="metric-number" style="color: #8b5cf6;">{taux}%</div>
                            <div class="metric-label">Précision</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Tableau
                    st.markdown("<br>", unsafe_allow_html=True)
                    st.dataframe(df_res, use_container_width=True, hide_index=True)
                    
                    # Export
                    st.markdown("<br>", unsafe_allow_html=True)
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, sheet_name='Résultats', index=False)
                        if erreurs:
                            pd.DataFrame({'Erreurs': erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    st.download_button("📥 Télécharger le rapport", output.getvalue(), "verification.xlsx", use_container_width=True)
                    
                    if incorrects == 0:
                        st.balloons()
                        st.success("✨ Félicitations ! Toutes les vérifications sont correctes.")
                    
                    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="info-box" style="text-align: center;">
        ✨ Commencez par charger les fichiers dans les zones ci-dessus
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div class="footer">
    <span>📐 Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for</span>
</div>
""", unsafe_allow_html=True)
