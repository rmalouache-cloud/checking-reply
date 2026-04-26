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
        df_copy['Part_N'] = df_copy.iloc[:, 0].astype(str).str.strip()
        df_copy['Description'] = df_copy.iloc[:, 1].astype(str).str.strip() if len(df_copy.columns) > 1 else ""
        df_copy['Packing_list_qty'] = pd.to_numeric(df_copy.iloc[:, 3], errors='coerce').fillna(0)
        df_copy['Qty_for'] = pd.to_numeric(df_copy.iloc[:, 4], errors='coerce').fillna(0)
        df_copy['Remarks'] = df_copy.iloc[:, 6].astype(str).str.strip()
        df_copy['Moka_reply'] = df_copy.iloc[:, 7].astype(str).str.strip()
        df_copy['Oversent_FRS'] = pd.to_numeric(df_copy.iloc[:, 8], errors='coerce').fillna(0)
        
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
    
    IMPORTANT: On filtre d'abord par Part N, puis on cherche l'IDL dans ce sous-ensemble
    """
    # S'assurer que les colonnes existent
    if len(df_stock.columns) < 11:
        raise ValueError(f"Fichier stock a seulement {len(df_stock.columns)} colonnes, besoin de 11 minimum")
    
    # Nettoyer Part N
    part_n_clean = str(part_n).strip()
    
    # Colonne PART NO. (colonne D, index 3)
    col_part_no = df_stock.columns[3]
    
    # === ÉTAPE 1: Filtrer UNIQUEMENT les lignes avec le Part N recherché ===
    mask_part = df_stock[col_part_no].astype(str).str.strip() == part_n_clean
    df_filtered_by_part = df_stock[mask_part].copy()
    
    if df_filtered_by_part.empty:
        # Afficher les premiers PART NO. disponibles
        part_no_disponibles = df_stock[col_part_no].astype(str).str.strip().head(10).tolist()
        raise ValueError(f"Part N '{part_n_clean}' non trouvé. Exemples: {part_no_disponibles}")
    
    # Réinitialiser l'index pour travailler sur ce sous-ensemble
    df_filtered_by_part = df_filtered_by_part.reset_index(drop=True)
    
    # === ÉTAPE 2: Dans ce sous-ensemble, chercher l'IDL ===
    col_odf = df_stock.columns[0]  # Colonne A
    
    # Trouver la position de l'IDL dans le sous-ensemble
    mask_idl = df_filtered_by_part[col_odf].astype(str).str.strip() == str(idl).strip()
    idx_in_filtered = df_filtered_by_part[mask_idl].index
    
    if len(idx_in_filtered) == 0:
        # Afficher les IDL disponibles pour ce Part N
        idl_disponibles = df_filtered_by_part[col_odf].astype(str).str.strip().tolist()
        raise ValueError(f"IDL '{idl}' non trouvé pour {part_n_clean}. IDL disponibles: {idl_disponibles}")
    
    position_idl = idx_in_filtered[0]
    
    # === ÉTAPE 3: Vérifier qu'on n'est pas à la première ligne du sous-ensemble ===
    if position_idl == 0:
        raise ValueError(f"L'IDL {idl} est la première occurrence de {part_n_clean}, pas de ligne précédente")
    
    # === ÉTAPE 4: Prendre la ligne précédente DANS LE MÊME PART N ===
    position_prec = position_idl - 1
    ligne_prec = df_filtered_by_part.iloc[position_prec]
    
    # Vérifier que c'est bien le même Part N (normalement oui car on est dans le filtre)
    part_n_prec = str(ligne_prec[col_part_no]).strip()
    
    if part_n_prec != part_n_clean:
        raise ValueError(f"ERREUR INTERNE: Part N précédent différent: '{part_n_prec}'")
    
    # Extraire la colonne K (index 10) = OVERSENT QTY
    oversent_stock = pd.to_numeric(ligne_prec.iloc[10], errors='coerce').fillna(0)
    
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
        status_text.text(f"Traitement: {ligne['Part_N'][:50]}...")
        
        # Extraire les données
        part_n = ligne['Part_N']
        description = ligne['Description'] if pd.notna(ligne['Description']) else ""
        qty_for = ligne['Qty_for']
        packing_qty = ligne['Packing_list_qty']
        oversent_frs = ligne['Oversent_FRS']
        nom_fichier_stock = ligne['Moka_reply']
        remarks = ligne['Remarks']
        
        # Nettoyer le nom du fichier stock
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
            st.warning(f"  ⚠️ Fichier '{nom_fichier_stock}' non trouvé")
            continue
        
        df_stock = dict_stocks[fichier_stock_trouve]
        
        try:
            # Récupérer l'oversent du stock (ligne précédente du même Part N)
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            
            # Calculer l'oversent réel
            oversent_reel = oversent_stock + packing_qty - qty_for
            
            # Calculer l'écart
            ecart = oversent_reel - oversent_frs
            
            # Vérifier si correct
            est_correct = abs(ecart) < 0.01
            
            # Afficher le résultat
            if est_correct:
                status = "✅ CORRECT"
                st.success(f"  ✅ {part_n[:40]}...")
            else:
                status = "❌ INCORRECT"
                st.error(f"  ❌ {part_n[:40]}... | Écart={ecart:.2f}")
            
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
                'Fichier stock': fichier_stock_trouve
            })
            
        except Exception as e:
            erreurs.append(f"{part_n}: {str(e)}")
            resultats.append({
                'Modèle': nom_modele,
                'Part N': part_n,
                'Description': description,
                'Remarks': remarks,
                'Status': f'❌ {str(e)[:80]}',
            })
            st.error(f"  ❌ {part_n[:40]}...: {str(e)[:100]}")
    
    progress_bar.empty()
    status_text.empty()
    return resultats, erreurs

# ==================== INTERFACE STREAMLIT ====================

with st.sidebar:
    st.header("📂 1. Chargement")
    
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'])
    stock_files = st.file_uploader("Fichiers stocks", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    st.markdown("---")
    st.markdown("### 📐 Logique de calcul")
    st.markdown("""
    1. **Filtrer** le fichier stock par Part N
    2. **Chercher** l'IDL dans ce sous-ensemble
    3. **Prendre** la ligne précédente (même Part N)
    4. **Extraire** la valeur colonne K
    5. **Calculer**: `Oversent_stock + Packing_qty - Qty_for`
    """)

# Zone principale
if reply_file and stock_files:
    st.success(f"✅ Reply: {reply_file.name}")
    st.success(f"✅ Stocks: {len(stock_files)} fichiers")
    
    with st.spinner("Chargement..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply is None:
            st.stop()
        dict_stocks = charger_stocks_depuis_upload(stock_files)
    
    # Aperçu
    with st.expander("📋 Feuilles trouvées"):
        for nom_feuille, df in dict_reply.items():
            st.write(f"**{nom_feuille}** - {len(df)} lignes")
    
    # Saisie IDL
    st.markdown("---")
    st.header("🔑 2. IDL par modèle")
    
    idl_par_modele = {}
    cols = st.columns(min(3, len(dict_reply)))
    
    for i, nom_modele in enumerate(dict_reply.keys()):
        with cols[i % len(cols)]:
            st.subheader(f"📱 {nom_modele}")
            idl = st.text_input(f"IDL", key=f"idl_{nom_modele}")
            if idl:
                idl_par_modele[nom_modele] = idl
    
    # Vérification
    st.markdown("---")
    if st.button("🚀 VÉRIFIER", type="primary", use_container_width=True):
        if not idl_par_modele:
            st.error("❌ Saisissez au moins un IDL")
        else:
            tous_resultats = []
            toutes_erreurs = []
            
            for nom_modele, df_feuille in dict_reply.items():
                if nom_modele in idl_par_modele:
                    st.markdown(f"### 📌 {nom_modele}")
                    resultats, erreurs = verifier_modele(
                        df_feuille, nom_modele, idl_par_modele[nom_modele], dict_stocks
                    )
                    tous_resultats.extend(resultats)
                    toutes_erreurs.extend(erreurs)
            
            # Résultats
            st.markdown("---")
            st.header("📊 Résultats")
            
            if tous_resultats:
                df_resultats = pd.DataFrame(tous_resultats)
                
                # Stats
                total = len(df_resultats)
                corrects = len(df_resultats[df_resultats['Status'] == '✅ CORRECT'])
                incorrects = len(df_resultats[df_resultats['Status'] == '❌ INCORRECT'])
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Total", total)
                col2.metric("✅ Corrects", corrects)
                col3.metric("❌ Incorrects", incorrects)
                
                # Tableau
                st.dataframe(df_resultats, use_container_width=True, hide_index=True)
                
                # Export
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                    if toutes_erreurs:
                        pd.DataFrame({'Erreurs': toutes_erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                
                st.download_button("📥 Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                
                if incorrects == 0:
                    st.balloons()
                    st.success("🎉 TOUT EST CORRECT !")
            else:
                st.error("Aucun résultat")

else:
    st.info("👈 Chargez les fichiers")
