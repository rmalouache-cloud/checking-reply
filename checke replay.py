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

def extraire_colonnes_position(df):
    """
    Extrait les colonnes par position (E, F, G, H)
    et crée un DataFrame standardisé
    """
    df_copy = df.copy()
    
    # Déterminer les colonnes par position
    if len(df_copy.columns) >= 8:
        # Colonne E (index 4) = Qty for
        # Colonne F (index 5) = Packing list qty
        # Colonne G (index 6) = Remarks
        # Colonne H (index 7) = Moka reply
        
        df_copy['Qty_for'] = df_copy.iloc[:, 4]
        df_copy['Packing_list_qty'] = df_copy.iloc[:, 5]
        df_copy['Remarks'] = df_copy.iloc[:, 6].astype(str).str.strip()
        df_copy['Moka_reply'] = df_copy.iloc[:, 7].astype(str).str.strip()
        
        # Colonne B (index 1) = Part N (généralement)
        df_copy['Part_N'] = df_copy.iloc[:, 1].astype(str).str.strip()
        
        # Colonne C (index 2) = Description (généralement)
        df_copy['Description'] = df_copy.iloc[:, 2].astype(str).str.strip() if len(df_copy.columns) > 2 else ""
        
        # Colonne pour Oversent FRS (chercher dans les colonnes)
        col_oversent = None
        for col in df_copy.columns:
            if 'oversent' in str(col).lower():
                col_oversent = col
                break
        if col_oversent is None:
            col_oversent = df_copy.columns[8] if len(df_copy.columns) > 8 else df_copy.columns[0]
        
        df_copy['Oversent_FRS'] = df_copy[col_oversent]
        
        return df_copy
    else:
        st.error(f"Fichier a seulement {len(df_copy.columns)} colonnes, besoin de 8 minimum")
        return None

def get_oversent_from_stock(df_stock, part_n, idl):
    """
    Dans le fichier stock:
    - Filtre par Part N
    - Trouve l'IDL dans colonne A
    - Retourne la valeur colonne K de la ligne précédente
    """
    # Trouver la colonne Part N dans le stock
    col_part_n = None
    for col in df_stock.columns:
        if 'part' in str(col).lower() or 'pn' in str(col).lower():
            col_part_n = col
            break
    
    if col_part_n is None:
        # Si non trouvé, utiliser colonne B (index 1)
        col_part_n = df_stock.columns[1] if len(df_stock.columns) > 1 else df_stock.columns[0]
    
    # Filtrer par Part N
    mask_part = df_stock[col_part_n].astype(str).str.strip() == str(part_n).strip()
    df_filtered = df_stock[mask_part]
    
    if df_filtered.empty:
        raise ValueError(f"Part N '{part_n}' non trouvé dans le stock")
    
    # Chercher l'IDL dans colonne A (ODF)
    col_odf = df_stock.columns[0]  # Colonne A
    
    mask_idl = df_filtered[col_odf].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask_idl].index
    
    if len(idx) == 0:
        # Afficher les IDL disponibles pour déboguer
        idl_disponibles = df_filtered[col_odf].astype(str).str.strip().tolist()
        raise ValueError(f"IDL '{idl}' non trouvé. IDL disponibles: {idl_disponibles[:5]}")
    
    ligne_idl = idx[0]
    
    if ligne_idl == 0:
        raise ValueError(f"L'IDL {idl} est à la première ligne, pas de ligne précédente")
    
    # Vérifier que la ligne précédente a le même Part N
    ligne_prec = ligne_idl - 1
    part_n_prec = str(df_stock.iloc[ligne_prec][col_part_n]).strip()
    
    if part_n_prec != str(part_n).strip():
        raise ValueError(f"Ligne précédente a un Part N différent: {part_n_prec}")
    
    # Extraire la colonne K (index 10) = OVERSENT QTY
    if df_stock.shape[1] > 10:
        oversent_stock = df_stock.iloc[ligne_prec, 10]
    else:
        # Si pas de colonne K, utiliser la dernière colonne
        oversent_stock = df_stock.iloc[ligne_prec, -1]
        st.warning(f"Colonne K (index 10) non trouvée, utilisation colonne {df_stock.columns[-1]}")
    
    # Convertir en numérique
    try:
        oversent_stock = float(oversent_stock)
    except:
        oversent_stock = 0
    
    return oversent_stock

def calculer_oversent_reel(oversent_stock, packing_qty, qty_for):
    """
    Calcul du véritable oversent:
    oversent_reel = oversent_stock + packing_qty - qty_for
    """
    return oversent_stock + packing_qty - qty_for

def verifier_modele(df_feuille, nom_modele, idl, dict_stocks):
    """Vérifie toutes les lignes d'une feuille (modèle)"""
    resultats = []
    erreurs = []
    
    # Standardiser les colonnes
    df_standard = extraire_colonnes_position(df_feuille)
    if df_standard is None:
        return resultats, ["Format de fichier incorrect"]
    
    # Filtrer sur Missing et shortage
    mask_missing = df_standard['Remarks'] == 'Missing'
    mask_shortage = df_standard['Remarks'] == 'shortage'
    df_filtre = df_standard[mask_missing | mask_shortage]
    
    if df_filtre.empty:
        st.info(f"📌 {nom_modele}: Aucune ligne 'Missing' ou 'shortage'")
        return resultats, erreurs
    
    st.write(f"**{nom_modele}** - {len(df_filtre)} lignes à vérifier (IDL: {idl})")
    
    # Barre de progression pour ce modèle
    progress_bar = st.progress(0)
    
    for idx, (index_ligne, ligne) in enumerate(df_filtre.iterrows()):
        # Mise à jour progression
        progress_bar.progress((idx + 1) / len(df_filtre))
        
        # Extraire les données
        part_n = ligne['Part_N']
        description = ligne['Description'] if pd.notna(ligne['Description']) else ""
        qty_for = ligne['Qty_for']
        packing_qty = ligne['Packing_list_qty']
        oversent_frs = ligne['Oversent_FRS']
        nom_fichier_stock = ligne['Moka_reply']
        
        # Nettoyer le nom du fichier
        nom_fichier_stock = nom_fichier_stock.replace('.xlsx', '').replace('.xls', '')
        
        # Chercher le fichier stock correspondant
        fichier_stock_trouve = None
        for nom_fichier in dict_stocks.keys():
            if nom_fichier_stock in nom_fichier or nom_fichier.startswith(nom_fichier_stock):
                fichier_stock_trouve = nom_fichier
                break
        
        if fichier_stock_trouve is None:
            erreurs.append(f"Fichier '{nom_fichier_stock}' non trouvé pour {part_n}")
            resultats.append({
                'Modèle': nom_modele,
                'Part N': part_n,
                'Description': description,
                'Status': '❌ Fichier stock manquant',
                'Fichier requis': nom_fichier_stock
            })
            continue
        
        df_stock = dict_stocks[fichier_stock_trouve]
        
        try:
            # Récupérer l'oversent du stock (ligne précédente)
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            
            # Calculer l'oversent réel
            oversent_reel = calculer_oversent_reel(oversent_stock, packing_qty, qty_for)
            
            # Calculer l'écart
            ecart = oversent_reel - oversent_frs
            
            # Vérifier si correct (tolérance 0.01)
            est_correct = abs(ecart) < 0.01
            
            # Afficher le résultat
            if est_correct:
                status = "✅ CORRECT"
                st.success(f"  ✅ {part_n}: {oversent_frs} = {oversent_reel:.2f}")
            else:
                status = "❌ INCORRECT"
                st.error(f"  ❌ {part_n}: FRS={oversent_frs} | Calculé={oversent_reel:.2f} | Écart={ecart:.2f}")
            
            # Ajouter au rapport
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Remarks': ligne['Remarks'],
                'Qty for (col E)': qty_for,
                'Packing list qty (col F)': packing_qty,
                'Oversent stock (ligne N-1)': oversent_stock,
                'Oversent FRS': oversent_frs,
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
                'Remarks': ligne['Remarks'],
                'Status': f'❌ Erreur: {str(e)[:100]}',
            })
            st.error(f"  ❌ {part_n}: Erreur - {str(e)}")
    
    progress_bar.empty()
    return resultats, erreurs

# ==================== INTERFACE STREAMLIT ====================

# Sidebar pour le chargement des fichiers
with st.sidebar:
    st.header("📂 1. Chargement des fichiers")
    
    reply_file = st.file_uploader(
        "Fichier reply.xlsx",
        type=['xlsx', 'xls'],
        help="Fichier de réponse du fournisseur avec plusieurs feuilles"
    )
    
    stock_files = st.file_uploader(
        "Fichiers stocks",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Sélectionnez tous les fichiers stock du fournisseur"
    )
    
    st.markdown("---")
    st.markdown("### 📐 Formule de calcul")
    st.latex(r'''
    \text{Oversent réel} = \text{Oversent stock (N-1)} + \text{Packing Qty} - \text{Qty for}
    ''')
    
    st.markdown("### 📌 Structure attendue")
    st.markdown("""
    **Fichier reply (positions):**
    - Colonne B = Part N
    - Colonne C = Description
    - Colonne E = Qty for
    - Colonne F = Packing list qty
    - Colonne G = Remarks ('Missing'/'shortage')
    - Colonne H = Moka reply (nom fichier stock)
    
    **Fichier stock:**
    - Colonne A = ODF (IDL)
    - Colonne K = OVERSENT QTY
    """)

# Zone principale
if reply_file and stock_files:
    st.success(f"✅ Fichier reply: {reply_file.name}")
    st.success(f"✅ {len(stock_files)} fichiers stocks chargés")
    
    # Chargement des données
    with st.spinner("Chargement des fichiers..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply is None:
            st.stop()
        dict_stocks = charger_stocks_depuis_upload(stock_files)
    
    # Aperçu des feuilles
    with st.expander("📋 Aperçu des feuilles trouvées"):
        for nom_feuille, df in dict_reply.items():
            st.write(f"**{nom_feuille}** - {len(df)} lignes, {len(df.columns)} colonnes")
    
    # Saisie des IDL par modèle
    st.markdown("---")
    st.header("🔑 2. Saisie des IDL par modèle")
    st.caption("Entrez l'IDL correspondant à chaque modèle (feuille)")
    
    idl_par_modele = {}
    
    # Créer des colonnes pour la saisie
    cols = st.columns(min(3, len(dict_reply)))
    
    for i, (nom_modele, df) in enumerate(dict_reply.items()):
        with cols[i % len(cols)]:
            st.subheader(f"📱 {nom_modele}")
            idl = st.text_input(
                f"IDL",
                key=f"idl_{nom_modele}",
                placeholder="Ex: IDL12345"
            )
            if idl:
                idl_par_modele[nom_modele] = idl
    
    # Bouton de vérification
    st.markdown("---")
    col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
    with col_btn2:
        verifier = st.button("🚀 LANCER LA VÉRIFICATION", type="primary", use_container_width=True)
    
    if verifier:
        if not idl_par_modele:
            st.error("❌ Veuillez saisir au moins un IDL")
        else:
            # Résultats
            tous_resultats = []
            toutes_erreurs = []
            
            st.markdown("---")
            st.header("📊 3. Résultats de la vérification")
            
            # Traitement par modèle
            for nom_modele, df_feuille in dict_reply.items():
                if nom_modele in idl_par_modele:
                    st.markdown(f"### 📌 Modèle: {nom_modele}")
                    
                    resultats, erreurs = verifier_modele(
                        df_feuille,
                        nom_modele,
                        idl_par_modele[nom_modele],
                        dict_stocks
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
                erreurs_count = len(df_resultats[df_resultats['Status'].str.startswith('❌', na=False)]) - incorrects if 'Status' in df_resultats.columns else 0
                
                col1.metric("📋 Total vérifié", total)
                col2.metric("✅ Corrects", corrects, delta=f"{corrects/total*100:.0f}%" if total > 0 else "0%")
                col3.metric("❌ Incorrects", incorrects, delta=f"-{incorrects/total*100:.0f}%" if total > 0 else "0%", delta_color="inverse")
                col4.metric("⚠️ Erreurs", erreurs_count)
                
                # Tableau détaillé
                st.subheader("📋 Détail des vérifications")
                
                # Colonnes à afficher
                colonnes_affichage = ['Modèle', 'Part N', 'Description', 'Remarks', 'Qty for (col E)', 
                                      'Packing list qty (col F)', 'Oversent FRS', 'Oversent calculé', 'Écart', 'Status']
                colonnes_disponibles = [col for col in colonnes_affichage if col in df_resultats.columns]
                
                st.dataframe(
                    df_resultats[colonnes_disponibles],
                    use_container_width=True,
                    hide_index=True
                )
                
                # Filtres
                st.subheader("🔍 Filtres")
                col_f1, col_f2 = st.columns(2)
                
                with col_f1:
                    if 'Status' in df_resultats.columns:
                        status_filter = st.multiselect(
                            "Filtrer par status",
                            options=df_resultats['Status'].unique(),
                            default=[]
                        )
                    else:
                        status_filter = []
                
                with col_f2:
                    if 'Modèle' in df_resultats.columns:
                        modele_filter = st.multiselect(
                            "Filtrer par modèle",
                            options=df_resultats['Modèle'].unique(),
                            default=[]
                        )
                    else:
                        modele_filter = []
                
                # Application des filtres
                df_filtre_aff = df_resultats.copy()
                if status_filter:
                    df_filtre_aff = df_filtre_aff[df_filtre_aff['Status'].isin(status_filter)]
                if modele_filter:
                    df_filtre_aff = df_filtre_aff[df_filtre_aff['Modèle'].isin(modele_filter)]
                
                if not df_filtre_aff.empty:
                    st.dataframe(
                        df_filtre_aff[colonnes_disponibles],
                        use_container_width=True,
                        hide_index=True
                    )
                
                # Affichage des erreurs
                if toutes_erreurs:
                    st.subheader("⚠️ Liste des erreurs")
                    for erreur in toutes_erreurs[:20]:
                        st.error(erreur)
                    if len(toutes_erreurs) > 20:
                        st.warning(f"... et {len(toutes_erreurs) - 20} autres erreurs")
                
                # Export Excel
                st.subheader("💾 Export des résultats")
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Feuille des résultats
                    df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                    
                    # Feuille des erreurs
                    if toutes_erreurs:
                        df_erreurs = pd.DataFrame({'Erreurs': toutes_erreurs})
                        df_erreurs.to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    # Feuille de statistiques
                    stats_data = {
                        'Statistique': ['Total vérifié', 'Corrects', 'Incorrects', 'Erreurs', 'Taux de réussite'],
                        'Valeur': [total, corrects, incorrects, erreurs_count, f"{corrects/total*100:.1f}%" if total > 0 else "0%"]
                    }
                    pd.DataFrame(stats_data).to_excel(writer, sheet_name='Statistiques', index=False)
                    
                    # Feuille des incorrects uniquement
                    if incorrects > 0:
                        df_incorrects = df_resultats[df_resultats['Status'] == '❌ INCORRECT']
                        df_incorrects.to_excel(writer, sheet_name='Incorrects', index=False)
                
                st.download_button(
                    label="📥 Télécharger le rapport Excel",
                    data=output.getvalue(),
                    file_name="verification_oversent.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Message final
                st.markdown("---")
                if incorrects == 0 and erreurs_count == 0:
                    st.balloons()
                    st.success("🎉 FÉLICITATIONS ! Toutes les vérifications sont correctes !")
                elif incorrects > 0:
                    st.warning(f"⚠️ Attention : {incorrects} incohérence(s) détectée(s) à corriger avec le fournisseur")
                else:
                    st.info("ℹ️ Quelques erreurs techniques, vérifiez les fichiers")
            
            else:
                st.error("❌ Aucun résultat généré")

elif reply_file and not stock_files:
    st.warning("⚠️ Veuillez ajouter les fichiers stocks")

elif not reply_file and stock_files:
    st.warning("⚠️ Veuillez ajouter le fichier reply.xlsx")

else:
    st.info("👈 Commencez par charger les fichiers dans la barre latérale")

# Footer
st.markdown("---")
st.markdown("""
### 📞 Support

**Formule de calcul :** `Oversent réel = Oversent stock (ligne N-1) + Packing list qty - Qty for`

En cas de problème, vérifiez que:
1. Les fichiers sont au format Excel (.xlsx)
2. Les colonnes sont aux bonnes positions (E, F, G, H dans reply)
3. Les IDL existent dans les fichiers stock
""")
