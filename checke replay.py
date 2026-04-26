import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(
    page_title="Vérification Réponses Fournisseur",
    page_icon="✅",
    layout="wide"
)

st.title("✅ Vérification des Réponses Fournisseur")
st.markdown("---")

# ==================== FONCTIONS PRINCIPALES ====================

def charger_toutes_les_feuilles_reply(uploaded_file):
    """Charge toutes les feuilles du fichier reply.xlsx"""
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        feuilles = {}
        for sheet_name in xlsx.sheet_names:
            feuilles[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return feuilles
    except Exception as e:
        st.error(f"Erreur chargement reply.xlsx: {str(e)}")
        return None

def charger_stocks_depuis_upload(uploaded_files):
    """Charge tous les fichiers stocks"""
    stocks = {}
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                stocks[uploaded_file.name] = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Erreur chargement {uploaded_file.name}: {str(e)}")
    return stocks

def extraire_colonnes_reply(df):
    """
    Extrait les colonnes du fichier reply selon les positions:
    A(0) -> Part N
    B(1) -> Description
    D(3) -> Packing list qty
    E(4) -> Qty for
    G(6) -> Remarks
    H(7) -> Moka reply
    I(8) -> Oversent qty
    """
    df_copy = df.copy()
    
    if len(df_copy.columns) >= 9:
        # Extraction par position
        df_copy['Part_N'] = df_copy.iloc[:, 0].astype(str).str.strip()  # Colonne A
        df_copy['Description'] = df_copy.iloc[:, 1].astype(str).str.strip() if len(df_copy.columns) > 1 else ""  # Colonne B
        df_copy['Packing_list_qty'] = pd.to_numeric(df_copy.iloc[:, 3], errors='coerce').fillna(0)  # Colonne D
        df_copy['Qty_for'] = pd.to_numeric(df_copy.iloc[:, 4], errors='coerce').fillna(0)  # Colonne E
        df_copy['Remarks'] = df_copy.iloc[:, 6].astype(str).str.strip()  # Colonne G
        df_copy['Moka_reply'] = df_copy.iloc[:, 7].astype(str).str.strip()  # Colonne H
        df_copy['Oversent_FRS'] = pd.to_numeric(df_copy.iloc[:, 8], errors='coerce').fillna(0)  # Colonne I
        
        return df_copy
    else:
        st.error(f"Fichier reply a seulement {len(df_copy.columns)} colonnes, besoin de 9 minimum")
        return None

def get_oversent_from_stock(df_stock, part_n, idl):
    """
    Dans le fichier stock:
    - Colonne A = ODF (IDL)
    - Colonne D = PART NO.
    - Colonne K = OVERSENT QTY
    """
    # S'assurer que les colonnes existent
    if len(df_stock.columns) < 11:
        raise ValueError(f"Fichier stock a seulement {len(df_stock.columns)} colonnes, besoin de 11 minimum")
    
    # Nettoyer Part N pour comparaison
    part_n_clean = str(part_n).strip()
    
    # Extraire la colonne PART NO. (colonne D, index 3)
    col_part_no = df_stock.columns[3]  # Colonne D
    
    # Filtrer par PART NO.
    mask_part = df_stock[col_part_no].astype(str).str.strip() == part_n_clean
    df_filtered = df_stock[mask_part]
    
    if df_filtered.empty:
        # Afficher les premiers PART NO. disponibles pour déboguer
        part_no_disponibles = df_stock[col_part_no].astype(str).str.strip().head(10).tolist()
        raise ValueError(f"Part N '{part_n_clean}' non trouvé. Exemples disponibles: {part_no_disponibles}")
    
    # Chercher l'IDL dans colonne A (ODF)
    col_odf = df_stock.columns[0]  # Colonne A
    
    mask_idl = df_filtered[col_odf].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask_idl].index
    
    if len(idx) == 0:
        # Afficher les IDL disponibles pour ce Part N
        idl_disponibles = df_filtered[col_odf].astype(str).str.strip().tolist()
        raise ValueError(f"IDL '{idl}' non trouvé pour {part_n_clean}. IDL disponibles: {idl_disponibles}")
    
    ligne_idl = idx[0]
    
    if ligne_idl == 0:
        raise ValueError(f"L'IDL {idl} est à la première ligne, pas de ligne précédente")
    
    # Vérifier que la ligne précédente a le même PART NO.
    ligne_prec = ligne_idl - 1
    part_n_prec = str(df_stock.iloc[ligne_prec][col_part_no]).strip()
    
    if part_n_prec != part_n_clean:
        raise ValueError(f"Ligne précédente a un PART NO. différent: '{part_n_prec}'")
    
    # Extraire la colonne K (index 10) = OVERSENT QTY
    oversent_stock = pd.to_numeric(df_stock.iloc[ligne_prec, 10], errors='coerce').fillna(0)
    
    return oversent_stock

def verifier_modele(df_feuille, nom_modele, idl, dict_stocks):
    """Vérifie toutes les lignes d'une feuille (modèle)"""
    resultats = []
    erreurs = []
    
    # Standardiser les colonnes du reply
    df_standard = extraire_colonnes_reply(df_feuille)
    if df_standard is None:
        return resultats, ["Format de fichier reply incorrect"]
    
    # Filtrer sur Missing et shortage
    mask_missing = df_standard['Remarks'] == 'Missing'
    mask_shortage = df_standard['Remarks'] == 'shortage'
    df_filtre = df_standard[mask_missing | mask_shortage]
    
    if df_filtre.empty:
        st.info(f"📌 {nom_modele}: Aucune ligne 'Missing' ou 'shortage'")
        return resultats, erreurs
    
    st.write(f"**{nom_modele}** - {len(df_filtre)} lignes à vérifier (IDL: {idl})")
    
    # Barre de progression
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, (index_ligne, ligne) in enumerate(df_filtre.iterrows()):
        # Mise à jour progression
        progress_bar.progress((idx + 1) / len(df_filtre))
        status_text.text(f"Traitement: {ligne['Part_N']}")
        
        # Extraire les données
        part_n = ligne['Part_N']
        description = ligne['Description'] if pd.notna(ligne['Description']) else ""
        qty_for = ligne['Qty_for']
        packing_qty = ligne['Packing_list_qty']
        oversent_frs = ligne['Oversent_FRS']
        nom_fichier_stock = ligne['Moka_reply']
        remarks = ligne['Remarks']
        
        # Nettoyer le nom du fichier stock (enlever l'extension si présente)
        nom_fichier_stock = nom_fichier_stock.replace('.xlsx', '').replace('.xls', '')
        
        # Chercher le fichier stock correspondant
        fichier_stock_trouve = None
        for nom_fichier in dict_stocks.keys():
            nom_fichier_clean = nom_fichier.replace('.xlsx', '').replace('.xls', '')
            if nom_fichier_stock in nom_fichier_clean or nom_fichier_clean in nom_fichier_stock:
                fichier_stock_trouve = nom_fichier
                break
        
        if fichier_stock_trouve is None:
            erreurs.append(f"Fichier '{nom_fichier_stock}' non trouvé pour {part_n}")
            resultats.append({
                'Modèle': nom_modele,
                'Part N': part_n,
                'Description': description,
                'Remarks': remarks,
                'Status': '❌ Fichier stock manquant',
                'Fichier requis': nom_fichier_stock
            })
            st.warning(f"  ⚠️ {part_n}: Fichier '{nom_fichier_stock}' non trouvé")
            continue
        
        df_stock = dict_stocks[fichier_stock_trouve]
        
        try:
            # Récupérer l'oversent du stock (ligne précédente)
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            
            # Calculer l'oversent réel
            # Formule: oversent_reel = oversent_stock + packing_qty - qty_for
            oversent_reel = oversent_stock + packing_qty - qty_for
            
            # Calculer l'écart
            ecart = oversent_reel - oversent_frs
            
            # Vérifier si correct (tolérance 0.01)
            est_correct = abs(ecart) < 0.01
            
            # Afficher le résultat
            if est_correct:
                status = "✅ CORRECT"
                st.success(f"  ✅ {part_n}: FRS={oversent_frs} | Calculé={oversent_reel:.2f}")
            else:
                status = "❌ INCORRECT"
                st.error(f"  ❌ {part_n}: FRS={oversent_frs} | Calculé={oversent_reel:.2f} | Écart={ecart:.2f}")
            
            # Ajouter au rapport
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Remarks': remarks,
                'Qty for (col E)': qty_for,
                'Packing list qty (col D)': packing_qty,
                'Oversent stock (ligne N-1)': oversent_stock,
                'Oversent FRS (col I)': oversent_frs,
                'Oversent calculé': round(oversent_reel, 2),
                'Écart': round(ecart, 2),
                'Status': status,
                'Fichier stock utilisé': fichier_stock_trouve
            })
            
        except Exception as e:
            erreurs.append(f"{part_n} ({nom_modele}): {str(e)}")
            resultats.append({
                'Modèle': nom_modele,
                'Part N': part_n,
                'Description': description,
                'Remarks': remarks,
                'Status': f'❌ Erreur: {str(e)[:100]}',
            })
            st.error(f"  ❌ {part_n}: {str(e)}")
    
    progress_bar.empty()
    status_text.empty()
    return resultats, erreurs

# ==================== INTERFACE STREAMLIT ====================

# Sidebar
with st.sidebar:
    st.header("📂 1. Chargement des fichiers")
    
    reply_file = st.file_uploader(
        "Fichier reply.xlsx",
        type=['xlsx', 'xls'],
        help="Fichier de réponse du fournisseur"
    )
    
    stock_files = st.file_uploader(
        "Fichiers stocks",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Sélectionnez tous les fichiers stock"
    )
    
    st.markdown("---")
    st.markdown("### 📐 Structure des fichiers")
    
    with st.expander("📁 Structure reply.xlsx"):
        st.markdown("""
        | Colonne | Contenu |
        |---------|---------|
        | A | Part N |
        | B | Description |
        | D | Packing list qty |
        | E | Qty for |
        | G | Remarks |
        | H | Moka reply |
        | I | Oversent qty |
        """)
    
    with st.expander("📁 Structure fichiers stock"):
        st.markdown("""
        | Colonne | Contenu |
        |---------|---------|
        | A | ODF (IDL) |
        | D | PART NO. |
        | K | OVERSENT QTY |
        """)
    
    st.markdown("### 📐 Formule de calcul")
    st.latex(r'''
    \text{Oversent réel} = \text{Oversent stock (K, N-1)} + \text{Packing Qty (col D)} - \text{Qty for (col E)}
    ''')

# Zone principale
if reply_file and stock_files:
    st.success(f"✅ Fichier reply: {reply_file.name}")
    st.success(f"✅ {len(stock_files)} fichiers stocks chargés")
    
    # Chargement
    with st.spinner("Chargement des fichiers..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply is None:
            st.stop()
        dict_stocks = charger_stocks_depuis_upload(stock_files)
    
    # Aperçu
    with st.expander("📋 Aperçu des feuilles"):
        for nom_feuille, df in dict_reply.items():
            st.write(f"**{nom_feuille}** - {len(df)} lignes")
            # Afficher un aperçu des premières lignes
            df_preview = df.iloc[:, [0, 1, 3, 4, 6, 7, 8]] if len(df.columns) >= 9 else df
            st.dataframe(df_preview.head(3), use_container_width=True)
    
    # Saisie IDL
    st.markdown("---")
    st.header("🔑 2. Saisie des IDL par modèle")
    
    idl_par_modele = {}
    cols = st.columns(min(3, len(dict_reply)))
    
    for i, nom_modele in enumerate(dict_reply.keys()):
        with cols[i % len(cols)]:
            st.subheader(f"📱 {nom_modele}")
            idl = st.text_input(f"IDL", key=f"idl_{nom_modele}", placeholder="Ex: IDL12345")
            if idl:
                idl_par_modele[nom_modele] = idl
    
    # Bouton vérification
    st.markdown("---")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        verifier = st.button("🚀 LANCER LA VÉRIFICATION", type="primary", use_container_width=True)
    
    if verifier:
        if not idl_par_modele:
            st.error("❌ Veuillez saisir au moins un IDL")
        else:
            tous_resultats = []
            toutes_erreurs = []
            
            st.markdown("---")
            st.header("📊 3. Résultats de la vérification")
            
            for nom_modele, df_feuille in dict_reply.items():
                if nom_modele in idl_par_modele:
                    st.markdown(f"### 📌 Modèle: {nom_modele}")
                    resultats, erreurs = verifier_modele(
                        df_feuille, nom_modele, idl_par_modele[nom_modele], dict_stocks
                    )
                    tous_resultats.extend(resultats)
                    toutes_erreurs.extend(erreurs)
                else:
                    st.warning(f"⚠️ Aucun IDL saisi pour {nom_modele}")
            
            # Rapport final
            st.markdown("---")
            st.header("📈 Synthèse finale")
            
            if tous_resultats:
                df_resultats = pd.DataFrame(tous_resultats)
                
                # Statistiques
                col1, col2, col3, col4 = st.columns(4)
                total = len(df_resultats)
                corrects = len(df_resultats[df_resultats['Status'] == '✅ CORRECT']) if 'Status' in df_resultats.columns else 0
                incorrects = len(df_resultats[df_resultats['Status'] == '❌ INCORRECT']) if 'Status' in df_resultats.columns else 0
                erreurs_count = total - corrects - incorrects
                
                col1.metric("📋 Total", total)
                col2.metric("✅ Corrects", corrects, delta=f"{corrects/total*100:.0f}%" if total > 0 else "0%")
                col3.metric("❌ Incorrects", incorrects)
                col4.metric("⚠️ Erreurs", erreurs_count)
                
                # Tableau
                st.subheader("📋 Détail des vérifications")
                colonnes_affichage = ['Modèle', 'Part N', 'Description', 'Remarks', 'Qty for (col E)', 
                                      'Packing list qty (col D)', 'Oversent FRS (col I)', 
                                      'Oversent calculé', 'Écart', 'Status']
                colonnes_disponibles = [col for col in colonnes_affichage if col in df_resultats.columns]
                st.dataframe(df_resultats[colonnes_disponibles], use_container_width=True, hide_index=True)
                
                # Filtres
                st.subheader("🔍 Filtres")
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    if 'Status' in df_resultats.columns:
                        status_filter = st.multiselect("Par status", df_resultats['Status'].unique(), default=[])
                with col_f2:
                    if 'Modèle' in df_resultats.columns:
                        modele_filter = st.multiselect("Par modèle", df_resultats['Modèle'].unique(), default=[])
                
                df_filtre_aff = df_resultats.copy()
                if status_filter:
                    df_filtre_aff = df_filtre_aff[df_filtre_aff['Status'].isin(status_filter)]
                if modele_filter:
                    df_filtre_aff = df_filtre_aff[df_filtre_aff['Modèle'].isin(modele_filter)]
                
                if not df_filtre_aff.empty:
                    st.dataframe(df_filtre_aff[colonnes_disponibles], use_container_width=True, hide_index=True)
                
                # Erreurs
                if toutes_erreurs:
                    st.subheader("⚠️ Liste des erreurs")
                    for erreur in toutes_erreurs[:20]:
                        st.error(erreur)
                
                # Export
                st.subheader("💾 Export")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                    if toutes_erreurs:
                        pd.DataFrame({'Erreurs': toutes_erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                
                st.download_button("📥 Télécharger Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                
                if incorrects == 0 and erreurs_count == 0:
                    st.balloons()
                    st.success("🎉 TOUTES LES VÉRIFICATIONS SONT CORRECTES !")
                elif incorrects > 0:
                    st.warning(f"⚠️ {incorrects} incohérence(s) détectée(s)")
            else:
                st.error("❌ Aucun résultat")

elif reply_file and not stock_files:
    st.warning("⚠️ Ajoutez les fichiers stocks")
elif not reply_file and stock_files:
    st.warning("⚠️ Ajoutez le fichier reply.xlsx")
else:
    st.info("👈 Chargez les fichiers dans la barre latérale")
