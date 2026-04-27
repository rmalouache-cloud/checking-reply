import streamlit as st
import pandas as pd
import io

# Configuration
st.set_page_config(
    page_title="Vérification Fournisseur",
    page_icon="✅",
    layout="wide"
)

# ==================== CSS ====================
st.markdown("""
<style>
    /* En-tête */
    .header {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    
    .header h1 {
        color: white;
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }
    
    .header p {
        color: #e0e0e0;
        margin: 0.3rem 0 0 0;
        font-size: 0.9rem;
    }
    
    /* Zone upload simplifiée */
    .upload-area {
        border: 2px dashed #2a5298;
        border-radius: 12px;
        padding: 1rem;
        text-align: center;
        background: #f8f9fa;
        margin-bottom: 1rem;
    }
    
    /* Métriques compactes */
    .metric-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        border-radius: 12px;
        padding: 0.8rem;
        text-align: center;
        color: white;
    }
    
    .metric-number {
        font-size: 1.8rem;
        font-weight: bold;
    }
    
    .metric-label {
        font-size: 0.8rem;
        opacity: 0.9;
    }
    
    /* Info box compacte */
    .info-box {
        background: #e8f0fe;
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-size: 0.85rem;
        margin-bottom: 1rem;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1rem;
        margin-top: 1.5rem;
        border-top: 1px solid #ddd;
        color: #666;
        font-size: 0.8rem;
    }
    
    /* Bouton */
    .stButton > button {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        font-weight: 600;
        border-radius: 8px;
    }
    
    /* Réduire les marges */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
    }
    
    hr {
        margin: 1rem 0;
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

# ==================== INTERFACE COMPACTE ====================

# Upload en ligne
col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-area">', unsafe_allow_html=True)
    reply_file = st.file_uploader("📊 Fichier Reply", type=['xlsx', 'xls'], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-area">', unsafe_allow_html=True)
    stock_files = st.file_uploader("📦 Fichiers Stock", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

if reply_file and stock_files:
    
    with st.spinner("📥 Chargement..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        # Info feuilles
        st.markdown(f"""
        <div class="info-box">
            📁 {len(dict_reply)} feuille(s) : {', '.join(dict_reply.keys())}
        </div>
        """, unsafe_allow_html=True)
        
        # IDL par modèle
        st.markdown("**🔑 IDL par modèle**")
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                idl = st.text_input(f"📱 {modele}", key=f"idl_{modele}", placeholder="IDL")
                if idl:
                    idl_par_modele[modele] = idl
        
        # Bouton
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            verifier = st.button("🚀 LANCER LA VÉRIFICATION", use_container_width=True)
        
        if verifier:
            if not idl_par_modele:
                st.warning("⚠️ Saisissez au moins un IDL")
            else:
                resultats = []
                erreurs = []
                
                with st.spinner("⏳ Vérification..."):
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
                    st.markdown("<hr>", unsafe_allow_html=True)
                    st.markdown("**📊 Résultats**")
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✅'])
                    incorrects = len(df_res[df_res['Status'] == '❌'])
                    taux = f"{corrects/total*100:.0f}" if total > 0 else "0"
                    
                    # Métriques compactes
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.markdown(f"""
                        <div class="metric-box">
                            <div class="metric-number">{total}</div>
                            <div class="metric-label">TOTAL</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown(f"""
                        <div class="metric-box" style="background: linear-gradient(135deg, #10b981 0%, #059669 100%);">
                            <div class="metric-number">{corrects}</div>
                            <div class="metric-label">✅ CORRECTS</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown(f"""
                        <div class="metric-box" style="background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%);">
                            <div class="metric-number">{incorrects}</div>
                            <div class="metric-label">❌ INCORRECTS</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with col4:
                        st.markdown(f"""
                        <div class="metric-box" style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);">
                            <div class="metric-number">{taux}%</div>
                            <div class="metric-label">TAUX</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Tableau
                    def color_status(val):
                        if val == '✅':
                            return 'background-color: #d1fae5; color: #065f46; font-weight: bold'
                        elif val == '❌':
                            return 'background-color: #fee2e2; color: #991b1b; font-weight: bold'
                        elif val == '⚠️':
                            return 'background-color: #fed7aa; color: #92400e; font-weight: bold'
                        return ''
                    
                    styled_df = df_res.style.map(color_status, subset=['Status'])
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
                        st.success("🎉 Toutes les vérifications sont correctes !")

else:
    st.info("👈 Chargez les fichiers pour commencer")

# Footer
st.markdown("""
<div class="footer">
    📐 Formule : Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for
</div>
""", unsafe_allow_html=True)
