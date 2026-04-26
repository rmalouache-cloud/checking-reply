import streamlit as st
import pandas as pd
import io
from pathlib import Path

# Configuration de la page
st.set_page_config(
    page_title="Vérification Réponses Fournisseur",
    page_icon="✅",
    layout="wide"
)

# Titre
st.title("✅ Vérification des Réponses Fournisseur")
st.markdown("---")

# Fonctions
def charger_toutes_les_feuilles_reply(uploaded_file):
    """Charge toutes les feuilles du fichier reply.xlsx uploadé"""
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        feuilles = {}
        for sheet_name in xlsx.sheet_names:
            feuilles[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return feuilles
    except Exception as e:
        st.error(f"Erreur lors du chargement de reply.xlsx: {str(e)}")
        return None

def charger_stocks_depuis_upload(uploaded_files):
    """Charge les fichiers stocks uploadés"""
    stocks = {}
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                stocks[uploaded_file.name] = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Erreur lors du chargement de {uploaded_file.name}: {str(e)}")
    return stocks

def get_oversent_from_stock(df_stock, part_n, idl):
    """
    Dans le fichier stock filtré par Part N,
    trouve l'IDL dans colonne A et retourne la valeur colonne K de la ligne précédente
    """
    # Chercher la colonne Part N
    col_part_n = None
    for col in df_stock.columns:
        if 'part' in col.lower() or 'pn' in col.lower() or 'part number' in col.lower():
            col_part_n = col
            break
    
    if col_part_n is None:
        col_part_n = df_stock.columns[0]
        st.warning(f"Colonne Part N non trouvée, utilisation de '{col_part_n}'")
    
    # Filtrer par Part N
    try:
        df_filtered = df_stock[df_stock[col_part_n].astype(str).str.strip() == str(part_n).strip()]
    except Exception as e:
        raise ValueError(f"Erreur lors du filtrage du Part N {part_n}: {str(e)}")
    
    if df_filtered.empty:
        raise ValueError(f"Part N {part_n} non trouvé dans le fichier stock")
    
    # Chercher la ligne avec l'IDL dans colonne A (ODF)
    col_odf = df_stock.columns[0]
    
    # Convertir en string pour comparaison
    try:
        mask = df_filtered[col_odf].astype(str).str.strip() == str(idl).strip()
        idx = df_filtered[mask].index
    except Exception as e:
        raise ValueError(f"Erreur lors de la recherche de l'IDL {idl}: {str(e)}")
    
    if len(idx) == 0:
        raise ValueError(f"IDL {idl} non trouvé pour le Part N {part_n}")
    
    ligne_idl = idx[0]
    
    if ligne_idl == 0:
        raise ValueError(f"L'IDL {idl} est à la première ligne, pas de ligne précédente")
    
    # Vérifier que la ligne précédente a le même Part N
    ligne_prec = ligne_idl - 1
    part_n_prec = df_stock.iloc[ligne_prec][col_part_n]
    
    if str(part_n_prec).strip() != str(part_n).strip():
        raise ValueError(f"La ligne précédente n'a pas le même Part N ({part_n_prec} != {part_n})")
    
    # Extraire la colonne K (OVERSENT QTY) - index 10
    if df_stock.shape[1] > 10:
        oversent_precedent = df_stock.iloc[ligne_prec, 10]
    else:
        oversent_precedent = df_stock.iloc[ligne_prec, -1]
        st.warning(f"Colonne K non trouvée, utilisation de la dernière colonne")
    
    return oversent_precedent

def verifier_modele(feuille_df, nom_modele, idl, dict_stocks, progress_bar=None, status_text=None):
    """Vérifie toutes les lignes d'une feuille (modèle)"""
    resultats = []
    erreurs = []
    
    # Vérifier les colonnes nécessaires
    colonnes_requises = ['Remarks', 'Part N', 'Qty for', 'Packing list qty', 'Oversent qty', 'Moka reply']
    colonnes_manquantes = [col for col in colonnes_requises if col not in feuille_df.columns]
    
    if colonnes_manquantes:
        erreurs.append(f"Colonnes manquantes dans la feuille {nom_modele}: {', '.join(colonnes_manquantes)}")
        return resultats, erreurs
    
    # Filtrer sur Missing et shortage
    try:
        df_filtre = feuille_df[feuille_df['Remarks'].astype(str).str.strip().isin(['Missing', 'shortage'])]
    except Exception as e:
        erreurs.append(f"Erreur lors du filtrage des remarques pour {nom_modele}: {str(e)}")
        return resultats, erreurs
    
    if df_filtre.empty:
        return resultats, erreurs
    
    total_lignes = len(df_filtre)
    
    for idx, (index_ligne, ligne) in enumerate(df_filtre.iterrows()):
        # Mise à jour de la progression
        if progress_bar and total_lignes > 0:
            progress = (idx + 1) / total_lignes
            progress_bar.progress(progress, text=f"Traitement {nom_modele}: {idx+1}/{total_lignes}")
        if status_text:
            status_text.text(f"Traitement: {nom_modele} - Part N {ligne['Part N']}")
        
        part_n = str(ligne['Part N']).strip()
        description = ligne.get('Description', '')
        qty_for = ligne['Qty for']
        packing_qty = ligne['Packing list qty']
        oversent_frs = ligne['Oversent qty']
        nom_fichier_stock = str(ligne['Moka reply']).strip()
        
        if nom_fichier_stock not in dict_stocks:
            erreurs.append(f"Fichier stock '{nom_fichier_stock}' non trouvé pour Part N {part_n} (modèle {nom_modele})")
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Qty for': qty_for,
                'Packing list qty': packing_qty,
                'Oversent FRS': oversent_frs,
                'Oversent calculé': None,
                'Écart': None,
                'Status': '❌ Fichier stock manquant',
                'Fichier stock requis': nom_fichier_stock
            })
            continue
        
        df_stock = dict_stocks[nom_fichier_stock]
        
        try:
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            oversent_reel = oversent_stock - qty_for + packing_qty
            
            # Comparaison
            est_correct = abs(oversent_reel - oversent_frs) < 0.01
            
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Qty for': qty_for,
                'Packing list qty': packing_qty,
                'Oversent FRS': oversent_frs,
                'Oversent calculé': round(oversent_reel, 2),
                'Écart': round(oversent_reel - oversent_frs, 2),
                'Status': '✅ Correct' if est_correct else '❌ Incorrect',
                'Fichier stock utilisé': nom_fichier_stock
            })
            
        except Exception as e:
            erreurs.append(f"Erreur pour {part_n} (modèle {nom_modele}): {str(e)}")
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Qty for': qty_for,
                'Packing list qty': packing_qty,
                'Oversent FRS': oversent_frs,
                'Oversent calculé': None,
                'Écart': None,
                'Status': f'❌ Erreur: {str(e)[:50]}',
                'Fichier stock utilisé': nom_fichier_stock
            })
    
    return resultats, erreurs

# Interface Streamlit
with st.sidebar:
    st.header("📂 Chargement des fichiers")
    
    # Upload du fichier reply
    reply_file = st.file_uploader(
        "1. Téléchargez le fichier reply.xlsx",
        type=['xlsx', 'xls'],
        help="Fichier de réponse du fournisseur"
    )
    
    # Upload des fichiers stocks
    stock_files = st.file_uploader(
        "2. Téléchargez les fichiers stocks",
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        help="Sélectionnez tous les fichiers stocks du fournisseur"
    )
    
    st.markdown("---")
    st.markdown("### 📋 Instructions")
    st.markdown("""
    1. Téléchargez le fichier `reply.xlsx`
    2. Téléchargez tous les fichiers stocks
    3. Pour chaque modèle, entrez l'IDL correspondant
    4. Lancez la vérification
    """)

# Zone principale
if reply_file and stock_files:
    st.success(f"✅ Fichier reply chargé : {reply_file.name}")
    st.success(f"✅ {len(stock_files)} fichiers stocks chargés")
    
    # Chargement des données
    with st.spinner("Chargement des fichiers..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply is None:
            st.stop()
        dict_stocks = charger_stocks_depuis_upload(stock_files)
    
    st.info(f"📊 {len(dict_reply)} feuilles trouvées dans reply.xlsx : {', '.join(dict_reply.keys())}")
    
    # Saisie des IDL par modèle
    st.markdown("---")
    st.header("🔑 Saisie des IDL par modèle")
    
    idl_par_modele = {}
    
    # Créer des colonnes dynamiques
    cols = st.columns(min(3, len(dict_reply)))
    
    for i, (nom_modele, df) in enumerate(dict_reply.items()):
        col_idx = i % len(cols) if cols else 0
        with cols[col_idx] if cols else st:
            st.subheader(f"📱 {nom_modele}")
            idl = st.text_input(f"IDL pour {nom_modele}", key=f"idl_{nom_modele}")
            if idl:
                idl_par_modele[nom_modele] = idl
    
    # Bouton de vérification
    st.markdown("---")
    if st.button("🚀 LANCER LA VÉRIFICATION", type="primary", use_container_width=True):
        if not idl_par_modele:
            st.error("❌ Veuillez saisir au moins un IDL")
        else:
            # Initialisation
            tous_resultats = []
            toutes_erreurs = []
            
            # Barre de progression globale
            progress_bar = st.progress(0, text="Démarrage...")
            status_text = st.empty()
            
            # Traitement par modèle
            modeles_traites = 0
            for nom_modele, df_feuille in dict_reply.items():
                if nom_modele in idl_par_modele:
                    idl = idl_par_modele[nom_modele]
                    status_text.info(f"📋 Traitement du modèle : {nom_modele} (IDL: {idl})")
                    
                    resultats, erreurs = verifier_modele(
                        df_feuille, nom_modele, idl, dict_stocks,
                        progress_bar, status_text
                    )
                    
                    tous_resultats.extend(resultats)
                    toutes_erreurs.extend(erreurs)
                    modeles_traites += 1
                else:
                    st.warning(f"⚠️ Aucun IDL saisi pour le modèle {nom_modele}")
            
            # Nettoyer la progression
            progress_bar.empty()
            status_text.empty()
            
            # Affichage des résultats
            st.markdown("---")
            st.header("📊 Résultats de la vérification")
            
            if tous_resultats:
                df_resultats = pd.DataFrame(tous_resultats)
                
                # Statistiques
                col1, col2, col3, col4 = st.columns(4)
                
                total = len(df_resultats)
                corrects = len(df_resultats[df_resultats['Status'] == '✅ Correct']) if 'Status' in df_resultats.columns else 0
                incorrects = len(df_resultats[df_resultats['Status'] == '❌ Incorrect']) if 'Status' in df_resultats.columns else 0
                erreurs_count = len(df_resultats[df_resultats['Status'].str.contains('Fichier stock manquant|Erreur', na=False)]) if 'Status' in df_resultats.columns else 0
                
                col1.metric("Total vérifiés", total)
                col2.metric("✅ Corrects", corrects)
                col3.metric("❌ Incorrects", incorrects)
                col4.metric("⚠️ Erreurs", erreurs_count)
                
                # Affichage du tableau des résultats
                st.subheader("📋 Détail des vérifications")
                
                # Colonnes à afficher
                colonnes_affichage = ['Modèle', 'IDL utilisé', 'Part N', 'Description', 'Qty for', 'Packing list qty', 'Oversent FRS', 'Oversent calculé', 'Écart', 'Status']
                colonnes_disponibles = [col for col in colonnes_affichage if col in df_resultats.columns]
                
                st.dataframe(
                    df_resultats[colonnes_disponibles],
                    use_container_width=True,
                    hide_index=True
                )
                
                # Filtres interactifs
                st.subheader("🔍 Filtres")
                col_filtre1, col_filtre2 = st.columns(2)
                
                with col_filtre1:
                    if 'Status' in df_resultats.columns:
                        status_filter = st.multiselect(
                            "Filtrer par status",
                            options=df_resultats['Status'].unique(),
                            default=[]
                        )
                    else:
                        status_filter = []
                
                with col_filtre2:
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
                
                st.dataframe(
                    df_filtre_aff[colonnes_disponibles],
                    use_container_width=True,
                    hide_index=True
                )
                
                # Affichage des erreurs
                if toutes_erreurs:
                    st.subheader("⚠️ Liste des erreurs rencontrées")
                    for erreur in toutes_erreurs[:10]:  # Limiter à 10 erreurs affichées
                        st.error(erreur)
                    if len(toutes_erreurs) > 10:
                        st.warning(f"... et {len(toutes_erreurs) - 10} autres erreurs")
                
                # Bouton de téléchargement
                st.subheader("💾 Export des résultats")
                
                # Conversion en Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                    if toutes_erreurs:
                        df_erreurs = pd.DataFrame({'Erreurs': toutes_erreurs})
                        df_erreurs.to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    # Ajouter une feuille de statistiques
                    stats_data = {
                        'Statistique': ['Total vérifiés', 'Corrects', 'Incorrects', 'Erreurs'],
                        'Valeur': [total, corrects, incorrects, erreurs_count]
                    }
                    df_stats = pd.DataFrame(stats_data)
                    df_stats.to_excel(writer, sheet_name='Statistiques', index=False)
                
                st.download_button(
                    label="📥 Télécharger le rapport Excel",
                    data=output.getvalue(),
                    file_name="verification_oversent.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Message final
                if incorrects > 0 or erreurs_count > 0:
                    st.warning(f"⚠️ Attention : {incorrects + erreurs_count} anomalies détectées !")
                else:
                    st.success("✅ Toutes les vérifications sont correctes !")
            
            else:
                st.warning("⚠️ Aucun résultat à afficher")

elif reply_file and not stock_files:
    st.warning("⚠️ Veuillez télécharger les fichiers stocks")
elif not reply_file and stock_files:
    st.warning("⚠️ Veuillez télécharger le fichier reply.xlsx")
else:
    st.info("👈 Veuillez charger les fichiers dans la barre latérale")

# Footer
st.markdown("---")
st.markdown("### 📌 Notes importantes")
st.markdown("""
- Les fichiers doivent être au format Excel (.xlsx ou .xls)
- Le fichier reply doit contenir les colonnes : `Remarks`, `Part N`, `Qty for`, `Packing list qty`, `Oversent qty`, `Moka reply`
- Les fichiers stocks doivent contenir une colonne `ODF` (colonne A) et une colonne `OVERSENT QTY` (colonne K)
- L'IDL est unique par modèle (chaque feuille du fichier reply)
""")
