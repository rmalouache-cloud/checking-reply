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
        background: #1e3c72;
        padding: 1.5rem;
        border-radius: 10px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .header h1 {
        color: white;
        margin: 0;
        font-size: 2rem;
    }
    
    .header p {
        color: #ccc;
        margin: 0.5rem 0 0 0;
    }
    
    /* Cartes */
    .card {
        background: white;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        border: 1px solid #e0e0e0;
    }
    
    .card-title {
        font-size: 1.2rem;
        font-weight: 600;
        margin-bottom: 1rem;
        color: #1e3c72;
        border-left: 3px solid #1e3c72;
        padding-left: 0.8rem;
    }
    
    /* Métriques */
    .metric-green {
        background: #d1fae5;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        border-left: 4px solid #10b981;
    }
    
    .metric-red {
        background: #fee2e2;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        border-left: 4px solid #ef4444;
    }
    
    .metric-blue {
        background: #e0e7ff;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        border-left: 4px solid #3b82f6;
    }
    
    .metric-number {
        font-size: 2rem;
        font-weight: bold;
    }
    
    .metric-label {
        font-size: 0.9rem;
        color: #666;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1rem;
        margin-top: 2rem;
        border-top: 1px solid #e0e0e0;
        color: #666;
    }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
st.markdown("""
<div class="header">
    <h1>✅ Vérification Réponses Fournisseur</h1>
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
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">📊 Fichier Reply</div>', unsafe_allow_html=True)
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'], label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">📦 Fichiers Stock</div>', unsafe_allow_html=True)
    stock_files = st.file_uploader("Fichiers stock", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

if reply_file and stock_files:
    
    with st.spinner("Chargement..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        # Info feuilles
        st.caption(f"📁 {len(dict_reply)} feuille(s) : {', '.join(dict_reply.keys())}")
        
        # IDL par modèle
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="card-title">🔑 IDL par modèle</div>', unsafe_allow_html=True)
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                idl = st.text_input(f"{modele}", key=f"idl_{modele}", placeholder="IDL")
                if idl:
                    idl_par_modele[modele] = idl
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        # Bouton
        if st.button("▶️ VÉRIFIER", type="primary", use_container_width=True):
            
            if not idl_par_modele:
                st.warning("Saisissez au moins un IDL")
            else:
                
                resultats = []
                erreurs = []
                
                with st.spinner("Vérification..."):
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
                    st.markdown("---")
                    st.subheader("📊 Résultats")
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✅'])
                    incorrects = len(df_res[df_res['Status'] == '❌'])
                    
                    # Métriques
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.markdown(f"""
                        <div class="metric-blue">
                            <div class="metric-number">{total}</div>
                            <div class="metric-label">Total</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with c2:
                        st.markdown(f"""
                        <div class="metric-green">
                            <div class="metric-number">{corrects}</div>
                            <div class="metric-label">✅ Corrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    with c3:
                        st.markdown(f"""
                        <div class="metric-red">
                            <div class="metric-number">{incorrects}</div>
                            <div class="metric-label">❌ Incorrects</div>
                        </div>
                        """, unsafe_allow_html=True)
                    
                    # Tableau
                    st.dataframe(df_res, use_container_width=True, hide_index=True)
                    
                    # Export
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, sheet_name='Résultats', index=False)
                        if erreurs:
                            pd.DataFrame({'Erreurs': erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    st.download_button("📥 Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                    
                    if incorrects == 0:
                        st.balloons()
                        st.success("✅ Tout est correct !")

else:
    st.info("👈 Chargez les fichiers")

# Footer
st.markdown("""
<div class="footer">
    📐 Formule: Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for
</div>
""", unsafe_allow_html=True)
