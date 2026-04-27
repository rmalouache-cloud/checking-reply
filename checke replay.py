import streamlit as st
import pandas as pd
import io

# Configuration
st.set_page_config(
    page_title="Vérification Fournisseur",
    page_icon="✅",
    layout="wide"
)

# ==================== CSS SIMPLE ====================
st.markdown("""
<style>
    /* En-tête */
    .header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .header h1 {
        color: white;
        margin: 0;
        font-size: 2rem;
    }
    
    .header p {
        color: #e0e0e0;
        margin: 0.5rem 0 0 0;
    }
    
    /* Cartes normales */
    .card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e0e0e0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    
    .card-title {
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 1rem;
        color: #1e3c72;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    /* Zone upload */
    .upload-area {
        border: 2px dashed #2a5298;
        border-radius: 10px;
        padding: 1.5rem;
        text-align: center;
        background: #fafafa;
    }
    
    /* Métriques */
    .metric-total {
        background: #3b82f6;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        color: white;
    }
    
    .metric-correct {
        background: #10b981;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        color: white;
    }
    
    .metric-incorrect {
        background: #ef4444;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        color: white;
    }
    
    .metric-taux {
        background: #f59e0b;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        color: white;
    }
    
    .metric-number {
        font-size: 2rem;
        font-weight: bold;
    }
    
    .metric-label {
        font-size: 0.8rem;
        margin-top: 0.3rem;
    }
    
    /* Info box - texte visible */
    .info-box {
        background: #e0e7ff;
        padding: 0.8rem;
        border-radius: 8px;
        margin-bottom: 1rem;
        color: #1e3c72;
        font-weight: 500;
    }
    
    .info-box strong {
        color: #1e40af;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1.5rem;
        margin-top: 2rem;
        border-top: 1px solid #e0e0e0;
        color: #666;
    }
    
    /* Bouton centré */
    .button-container {
        display: flex;
        justify-content: center;
        margin: 1rem 0;
    }
    
    .stButton {
        display: flex;
        justify-content: center;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border: none;
        padding: 0.7rem 3rem;
        font-weight: 600;
        border-radius: 8px;
        font-size: 1rem;
        cursor: pointer;
        min-width: 250px;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
    }
    
    /* Input */
    .stTextInput > div > div > input {
        border-radius: 8px;
        border: 1px solid #d0d0d0;
        padding: 0.5rem;
    }
    
    /* Style du tableau */
    .dataframe {
        border: 1px solid #ddd !important;
        border-radius: 8px !important;
        overflow: hidden !important;
    }
    
    .dataframe th {
        background-color: #1e3c72 !important;
        color: white !important;
        font-weight: 600 !important;
        padding: 10px !important;
        border: 1px solid #2a5298 !important;
    }
    
    .dataframe td {
        border: 1px solid #ddd !important;
        padding: 8px !important;
    }
    
    /* Style pour le dataframe Streamlit */
    .stDataFrame {
        border: 1px solid #ddd;
        border-radius: 8px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
st.markdown("""
<div class="header">
    <h1>✅ Vérification des Réponses Fournisseur</h1>
    <p>Contrôle automatique des quantités Oversent vs Stock</p>
</div>
""", unsafe_allow_html=True)

# ==================== FONCTIONS ====================

def charger_feuilles_reply(uploaded_file):
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        return {sheet: pd.read_excel(uploaded_file, sheet_name=sheet) for sheet in xlsx.sheet_names}
    except Exception as e:
        st.error(f"❌ Erreur: {e}")
        return None

def charger_stocks(uploaded_files):
    stocks = {}
    for f in uploaded_files:
        try:
            stocks[f.name] = pd.read_excel(f)
        except Exception as e:
            st.error(f"❌ Erreur {f.name}: {e}")
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
        raise ValueError(f"Part N non trouvé")
    
    mask_idl = df_filtered[col_idl].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask_idl].index
    
    if len(idx) == 0:
        raise ValueError(f"IDL non trouvé")
    
    pos = idx[0]
    if pos == 0:
        raise ValueError(f"IDL à la première ligne")
    
    val = df_filtered.iloc[pos - 1, 10]
    return float(val) if pd.notna(val) else 0.0

# ==================== INTERFACE ====================

# Upload
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span>📊</span> Fichier Reply
        </div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="card">
        <div class="card-title">
            <span>📦</span> Fichiers Stock
        </div>
        <div class="upload-area">
    """, unsafe_allow_html=True)
    stock_files = st.file_uploader("Fichiers stock", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div></div>', unsafe_allow_html=True)

if reply_file and stock_files:
    
    with st.spinner("📥 Chargement..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        # Info box avec texte visible
        st.markdown(f"""
        <div class="info-box">
            📁 <strong>{len(dict_reply)} feuille(s) trouvée(s) : {', '.join(dict_reply.keys())}</strong>
        </div>
        """, unsafe_allow_html=True)
        
        # IDL par modèle
        st.markdown("""
        <div class="card">
            <div class="card-title">
                <span>🔑</span> IDL par modèle
            </div>
        """, unsafe_allow_html=True)
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                st.markdown(f"**📱 {modele}**")
                idl = st.text_input("", key=f"idl_{modele}", placeholder="Entrez l'IDL", label_visibility="collapsed")
                if idl:
                    idl_par_modele[modele] = idl
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bouton centré
        st.markdown('<div class="button-container">', unsafe_allow_html=True)
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            verifier = st.button("🚀 LANCER LA VÉRIFICATION", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        if verifier:
            if not idl_par_modele:
                st.warning("⚠️ Veuillez saisir au moins un IDL")
            else:
                resultats = []
                erreurs = []
                
                with st.spinner("⏳ Vérification en cours..."):
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
                                    'Remarks': remarks, 'Status': '❌'
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
                
                if resultats:
                    st.markdown("---")
                    st.markdown("""
                    <div class="card">
                        <div class="card-title">
                            <span>📊</span> Résultats de la vérification
                        </div>
                    """, unsafe_allow_html=True)
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✅'])
                    incorrects = len(df_res[df_res['Status'] == '❌'])
                    taux = f"{corrects/total*100:.0f}" if total > 0 else "0"
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-total">
                            <div class="metric-number">{total}</div>
                            <div class="metric-label">📊 TOTAL</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-correct">
                            <div class="metric-number">{corrects}</div>
                            <div class="metric-label">✅ CORRECTS</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-incorrect">
                            <div class="metric-number">{incorrects}</div>
                            <div class="metric-label">❌ INCORRECTS</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown(f"""
                        <div class="metric-taux">
                            <div class="metric-number">{taux}%</div>
                            <div class="metric-label">🎯 TAUX</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # Fonction de coloration pour le tableau
                    def color_status(val):
                        if val == '✅':
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold; text-align: center'
                        elif val == '❌':
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold; text-align: center'
                        elif val == '⚠️':
                            return 'background-color: #fed7aa; color: #92400e; font-weight: bold; text-align: center'
                        return ''
                    
                    # Appliquer le style avec bordures
                    styled_df = df_res.style.map(color_status, subset=['Status']).set_properties(**{
                        'border': '1px solid #ddd',
                        'padding': '8px'
                    }).set_table_styles([
                        {'selector': 'thead th', 'props': [('background-color', '#1e3c72'), ('color', 'white'), ('padding', '10px'), ('border', '1px solid #2a5298')]},
                        {'selector': 'tbody td', 'props': [('border', '1px solid #ddd')]}
                    ])
                    
                    st.dataframe(styled_df, use_container_width=True, hide_index=True)
                    
                    # Export
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, sheet_name='Résultats', index=False)
                        if erreurs:
                            pd.DataFrame({'Erreurs': erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    st.download_button("📥 Télécharger Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                    
                    if incorrects == 0:
                        st.balloons()
                        st.success("🎉 Félicitations ! Toutes les vérifications sont correctes !")
                    
                    st.markdown('</div>', unsafe_allow_html=True)

else:
    st.info("👈 Veuillez charger les fichiers pour commencer")

# Footer
st.markdown("""
<div class="footer">
    📐 Formule : Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for
</div>
""", unsafe_allow_html=True)
