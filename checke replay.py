import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(
    page_title="Vérification Réponses Fournisseur",
    page_icon="✅",
    layout="wide"
)

# ==================== CSS PERSONNALISÉ ====================
st.markdown("""
<style>
    /* Style général */
    .main {
        padding: 0rem 1rem;
    }
    
    /* En-tête */
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 15px;
        margin-bottom: 2rem;
        color: white;
        text-align: center;
    }
    
    .main-header h1 {
        margin: 0;
        font-size: 2.5rem;
        font-weight: 700;
    }
    
    .main-header p {
        margin: 0.5rem 0 0 0;
        opacity: 0.9;
    }
    
    /* Cartes */
    .card {
        background: white;
        border-radius: 15px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        border: 1px solid #e0e0e0;
    }
    
    .card-title {
        font-size: 1.3rem;
        font-weight: 600;
        margin-bottom: 1rem;
        color: #333;
        border-left: 4px solid #667eea;
        padding-left: 1rem;
    }
    
    /* Zone de upload */
    .upload-area {
        border: 2px dashed #667eea;
        border-radius: 10px;
        padding: 1.5rem;
        text-align: center;
        background: #f8f9ff;
    }
    
    /* Métriques */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
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
        font-size: 0.9rem;
        opacity: 0.9;
    }
    
    /* Boutons */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.5rem 2rem;
        font-weight: 600;
        border-radius: 10px;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 5px 15px rgba(102,126,234,0.4);
    }
    
    /* Tableau */
    .dataframe {
        font-size: 0.9rem;
    }
    
    /* Badges */
    .badge-success {
        background: #10b981;
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        font-size: 0.8rem;
        display: inline-block;
    }
    
    .badge-error {
        background: #ef4444;
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        font-size: 0.8rem;
        display: inline-block;
    }
    
    .badge-warning {
        background: #f59e0b;
        color: white;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        font-size: 0.8rem;
        display: inline-block;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem;
        margin-top: 2rem;
        border-top: 1px solid #e0e0e0;
        color: #666;
    }
    
    /* Expandeur stylisé */
    .streamlit-expanderHeader {
        font-weight: 600;
        color: #667eea;
    }
    
    hr {
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
st.markdown("""
<div class="main-header">
    <h1>✅ Vérification des Réponses Fournisseur</h1>
    <p>Comparez les quantités Oversent avec les fichiers de stock</p>
</div>
""", unsafe_allow_html=True)

# ==================== FONCTIONS ====================

def charger_toutes_les_feuilles_reply(uploaded_file):
    """Charge toutes les feuilles du fichier reply.xlsx"""
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        feuilles = {}
        for sheet_name in xlsx.sheet_names:
            feuilles[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return feuilles
    except Exception as e:
        st.error(f"❌ Erreur chargement reply.xlsx: {str(e)}")
        return None

def charger_stocks_depuis_upload(uploaded_files):
    """Charge tous les fichiers stocks"""
    stocks = {}
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                stocks[uploaded_file.name] = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"❌ Erreur chargement {uploaded_file.name}: {str(e)}")
    return stocks

def extraire_colonnes_reply(df):
    """Extrait les colonnes du fichier reply selon les positions"""
    df_copy = df.copy()
    
    if len(df_copy.columns) >= 9:
        df_copy['Part_N'] = df_copy.iloc[:, 0].astype(str).str.strip()
        df_copy['Description'] = df_copy.iloc[:, 1].astype(str).str.strip() if len(df_copy.columns) > 1 else ""
        df_copy['Packing_list_qty'] = pd.to_numeric(df_copy.iloc[:, 3], errors='coerce').fillna(0)
        df_copy['Qty_for'] = pd.to_numeric(df_copy.iloc[:, 4], errors='coerce').fillna(0)
        df_copy['Remarks'] = df_copy.iloc[:, 6].astype(str).str.strip()
        df_copy['Moka_reply'] = df_copy.iloc[:, 7].astype(str).str.strip()
        df_copy['Oversent_FRS'] = pd.to_numeric(df_copy.iloc[:, 8], errors='coerce').fillna(0)
        return df_copy
    else:
        return None

def get_oversent_from_stock(df_stock, part_n, idl):
    """Récupère l'oversent du stock (ligne précédente du même Part N)"""
    if len(df_stock.columns) < 11:
        raise ValueError(f"Fichier stock a seulement {len(df_stock.columns)} colonnes")
    
    part_n_clean = str(part_n).strip()
    col_part_no = df_stock.columns[3]
    
    # Filtrer par Part N
    mask_part = df_stock[col_part_no].astype(str).str.strip() == part_n_clean
    df_filtered_by_part = df_stock[mask_part].copy()
    
    if df_filtered_by_part.empty:
        raise ValueError(f"Part N '{part_n_clean}' non trouvé")
    
    df_filtered_by_part = df_filtered_by_part.reset_index(drop=True)
    
    # Chercher l'IDL
    col_odf = df_stock.columns[0]
    mask_idl = df_filtered_by_part[col_odf].astype(str).str.strip() == str(idl).strip()
    idx_in_filtered = df_filtered_by_part[mask_idl].index
    
    if len(idx_in_filtered) == 0:
        raise ValueError(f"IDL '{idl}' non trouvé pour {part_n_clean}")
    
    position_idl = idx_in_filtered[0]
    
    if position_idl == 0:
        raise ValueError(f"IDL {idl} est la première occurrence")
    
    position_prec = position_idl - 1
    ligne_prec = df_filtered_by_part.iloc[position_prec]
    
    # Extraire la colonne K
    oversent_value = ligne_prec.iloc[10]
    try:
        oversent_stock = float(oversent_value) if pd.notna(oversent_value) else 0.0
    except (ValueError, TypeError):
        oversent_stock = 0.0
    
    return oversent_stock

# ==================== INTERFACE PRINCIPALE ====================

# Zone de dépôt des fichiers
with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">📁 1. Chargement des fichiers</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown('<div class="upload-area">', unsafe_allow_html=True)
        reply_file = st.file_uploader(
            "📊 Fichier reply.xlsx",
            type=['xlsx', 'xls'],
            help="Fichier de réponse du fournisseur",
            label_visibility="collapsed"
        )
        if reply_file:
            st.success(f"✅ {reply_file.name}")
        st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="upload-area">', unsafe_allow_html=True)
        stock_files = st.file_uploader(
            "📦 Fichiers stocks",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Sélectionnez tous les fichiers stock",
            label_visibility="collapsed"
        )
        if stock_files:
            st.success(f"✅ {len(stock_files)} fichier(s) chargé(s)")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# Si les fichiers sont chargés
if reply_file and stock_files:
    
    # Chargement des données
    with st.spinner("📥 Chargement des fichiers en cours..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply is None:
            st.stop()
        dict_stocks = charger_stocks_depuis_upload(stock_files)
    
    # Aperçu des feuilles
    with st.expander("📋 Aperçu des feuilles du fichier reply"):
        cols = st.columns(len(dict_reply))
        for i, (nom_feuille, df) in enumerate(dict_reply.items()):
            with cols[i]:
                st.markdown(f"**📱 {nom_feuille}**")
                st.caption(f"{len(df)} lignes, {len(df.columns)} colonnes")
                st.dataframe(df.head(2), use_container_width=True, hide_index=True)
    
    # Saisie des IDL par modèle
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title">🔑 2. Saisie des IDL par modèle</div>', unsafe_allow_html=True)
    st.caption("Entrez l'IDL correspondant à chaque modèle (feuille)")
    
    idl_par_modele = {}
    nb_modeles = len(dict_reply)
    cols = st.columns(min(4, nb_modeles))
    
    for i, nom_modele in enumerate(dict_reply.keys()):
        with cols[i % len(cols)]:
            st.markdown(f"**📱 {nom_modele}**")
            idl = st.text_input(
                "IDL",
                key=f"idl_{nom_modele}",
                placeholder="Ex: IDL12345",
                label_visibility="collapsed"
            )
            if idl:
                idl_par_modele[nom_modele] = idl
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Bouton de vérification
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        verifier = st.button("🚀 LANCER LA VÉRIFICATION", use_container_width=True)
    
    if verifier:
        if not idl_par_modele:
            st.warning("⚠️ Veuillez saisir au moins un IDL")
        else:
            tous_resultats = []
            toutes_erreurs = []
            
            # Traitement par modèle
            for nom_modele, df_feuille in dict_reply.items():
                if nom_modele in idl_par_modele:
                    st.markdown(f"### 📌 Modèle: {nom_modele}")
                    
                    idl = idl_par_modele[nom_modele]
                    df_standard = extraire_colonnes_reply(df_feuille)
                    
                    if df_standard is not None:
                        # Filtrer sur Missing et shortage
                        mask = df_standard['Remarks'].isin(['Missing', 'shortage'])
                        df_filtre = df_standard[mask]
                        
                        if not df_filtre.empty:
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            for idx, (_, ligne) in enumerate(df_filtre.iterrows()):
                                progress_bar.progress((idx + 1) / len(df_filtre))
                                status_text.info(f"Traitement: {ligne['Part_N'][:50]}...")
                                
                                part_n = str(ligne['Part_N'])
                                description = str(ligne['Description'])[:50]
                                qty_for = float(ligne['Qty_for'])
                                packing_qty = float(ligne['Packing_list_qty'])
                                oversent_frs = float(ligne['Oversent_FRS'])
                                nom_fichier_stock = str(ligne['Moka_reply']).replace('.xlsx', '').replace('.xls', '')
                                remarks = str(ligne['Remarks'])
                                
                                # Chercher le fichier stock
                                fichier_trouve = None
                                for fn in dict_stocks.keys():
                                    if nom_fichier_stock in fn.replace('.xlsx', '').replace('.xls', ''):
                                        fichier_trouve = fn
                                        break
                                
                                if fichier_trouve:
                                    try:
                                        oversent_stock = get_oversent_from_stock(dict_stocks[fichier_trouve], part_n, idl)
                                        oversent_reel = oversent_stock + packing_qty - qty_for
                                        ecart = oversent_reel - oversent_frs
                                        est_correct = abs(ecart) < 0.01
                                        
                                        if est_correct:
                                            st.success(f"✅ {part_n[:40]}")
                                        else:
                                            st.error(f"❌ {part_n[:40]} | FRS={oversent_frs} | Calc={oversent_reel:.2f} | Écart={ecart:.2f}")
                                        
                                        tous_resultats.append({
                                            'Modèle': nom_modele, 'Part N': part_n, 'Description': description,
                                            'Remarks': remarks, 'IDL': idl,
                                            'Qty for': qty_for, 'Packing Qty': packing_qty,
                                            'Oversent Stock': oversent_stock, 'Oversent FRS': oversent_frs,
                                            'Oversent Calculé': round(oversent_reel, 2), 'Écart': round(ecart, 2),
                                            'Status': '✅ Correct' if est_correct else '❌ Incorrect'
                                        })
                                    except Exception as e:
                                        st.error(f"❌ {part_n[:40]}: {str(e)[:80]}")
                                        toutes_erreurs.append(f"{part_n}: {str(e)}")
                                        tous_resultats.append({
                                            'Modèle': nom_modele, 'Part N': part_n, 'Description': description,
                                            'Remarks': remarks, 'Status': f'❌ Erreur'
                                        })
                                else:
                                    st.warning(f"⚠️ {part_n[:40]}: Fichier '{nom_fichier_stock}' non trouvé")
                                    toutes_erreurs.append(f"Fichier '{nom_fichier_stock}' non trouvé")
                            
                            progress_bar.empty()
                            status_text.empty()
                        else:
                            st.info(f"Aucune ligne 'Missing' ou 'shortage' dans {nom_modele}")
            
            # Résultats finaux
            if tous_resultats:
                st.markdown("---")
                st.markdown('<div class="card">', unsafe_allow_html=True)
                st.markdown('<div class="card-title">📊 3. Résultats de la vérification</div>', unsafe_allow_html=True)
                
                df_resultats = pd.DataFrame(tous_resultats)
                
                # Métriques
                col1, col2, col3, col4 = st.columns(4)
                total = len(df_resultats)
                corrects = len(df_resultats[df_resultats['Status'] == '✅ Correct'])
                incorrects = len(df_resultats[df_resultats['Status'] == '❌ Incorrect'])
                
                with col1:
                    st.markdown(f'<div class="metric-card"><div class="metric-number">{total}</div><div class="metric-label">Total vérifié</div></div>', unsafe_allow_html=True)
                with col2:
                    st.markdown(f'<div class="metric-card"><div class="metric-number">{corrects}</div><div class="metric-label">✅ Corrects</div></div>', unsafe_allow_html=True)
                with col3:
                    st.markdown(f'<div class="metric-card"><div class="metric-number">{incorrects}</div><div class="metric-label">❌ Incorrects</div></div>', unsafe_allow_html=True)
                with col4:
                    taux = f"{corrects/total*100:.1f}%" if total > 0 else "0%"
                    st.markdown(f'<div class="metric-card"><div class="metric-number">{taux}</div><div class="metric-label">Taux de réussite</div></div>', unsafe_allow_html=True)
                
                # Tableau des résultats
                st.dataframe(df_resultats, use_container_width=True, hide_index=True)
                
                # Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                    if toutes_erreurs:
                        pd.DataFrame({'Erreurs': toutes_erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                
                st.download_button(
                    label="📥 Télécharger le rapport Excel",
                    data=output.getvalue(),
                    file_name="verification_oversent.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.markdown('</div>', unsafe_allow_html=True)
                
                if incorrects == 0:
                    st.balloons()
                    st.success("🎉 FÉLICITATIONS ! Toutes les vérifications sont correctes !")
            else:
                st.warning("Aucun résultat à afficher")

# Footer
st.markdown("""
<div class="footer">
    <p>📐 Formule de calcul: <strong>Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for</strong></p>
    <p style="font-size: 0.8rem;">© 2024 - Vérification Réponses Fournisseur</p>
</div>
""", unsafe_allow_html=True)
