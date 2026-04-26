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

def charger_toutes_les_feuilles_reply(uploaded_file):
    """Charge toutes les feuilles du fichier reply.xlsx uploadé"""
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        feuilles = {}
        for sheet_name in xlsx.sheet_names:
            feuilles[sheet_name] = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return feuilles
    except Exception as e:
        st.error(f"Erreur: {str(e)}")
        return None

def charger_stocks_depuis_upload(uploaded_files):
    """Charge les fichiers stocks uploadés"""
    stocks = {}
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                stocks[uploaded_file.name] = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Erreur chargement {uploaded_file.name}: {str(e)}")
    return stocks

def get_oversent_from_stock(df_stock, part_n, idl):
    """Trouve l'IDL et retourne la valeur de la ligne précédente"""
    # Chercher la colonne Part N
    col_part_n = None
    for col in df_stock.columns:
        if 'part' in col.lower() or 'pn' in col.lower() or 'part number' in col.lower():
            col_part_n = col
            break
    
    if col_part_n is None:
        col_part_n = df_stock.columns[0]
    
    # Filtrer par Part N
    df_filtered = df_stock[df_stock[col_part_n].astype(str).str.strip() == str(part_n).strip()]
    
    if df_filtered.empty:
        raise ValueError(f"Part N {part_n} non trouvé")
    
    # Chercher l'IDL dans colonne A (ODF)
    col_odf = df_stock.columns[0]
    mask = df_filtered[col_odf].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask].index
    
    if len(idx) == 0:
        raise ValueError(f"IDL {idl} non trouvé")
    
    ligne_idl = idx[0]
    
    if ligne_idl == 0:
        raise ValueError(f"IDL à la première ligne")
    
    ligne_prec = ligne_idl - 1
    
    # Extraire colonne K (index 10)
    if df_stock.shape[1] > 10:
        oversent_precedent = df_stock.iloc[ligne_prec, 10]
    else:
        oversent_precedent = df_stock.iloc[ligne_prec, -1]
    
    return oversent_precedent

def extraire_remarks_colonne_g(df):
    """
    Extrait la colonne G (index 6) quelle que soit son en-tête
    et crée une colonne 'Remarks' standardisée
    """
    df_copy = df.copy()
    
    # La colonne G est l'index 6 (7ème colonne)
    if len(df_copy.columns) > 6:
        # Créer une colonne Remarks à partir de la colonne G
        df_copy['Remarks'] = df_copy.iloc[:, 6].astype(str).str.strip()
        return df_copy
    else:
        st.error(f"Le fichier n'a que {len(df_copy.columns)} colonnes, pas de colonne G (index 6)")
        return None

def verifier_modele(feuille_df, nom_modele, idl, dict_stocks, progress_bar=None, status_text=None):
    """Vérifie toutes les lignes d'une feuille (modèle)"""
    resultats = []
    erreurs = []
    
    # Extraire la colonne G comme Remarks
    feuille_df = extraire_remarks_colonne_g(feuille_df)
    if feuille_df is None:
        return resultats, ["Impossible d'extraire la colonne G"]
    
    # Définir les noms de colonnes attendus
    # Part N peut être dans différentes colonnes
    col_part_n = None
    for col in feuille_df.columns:
        if 'part' in col.lower() or 'pn' in col.lower():
            col_part_n = col
            break
    
    if col_part_n is None:
        # Si on trouve pas, utiliser colonne B (index 1) souvent c'est là
        col_part_n = feuille_df.columns[1] if len(feuille_df.columns) > 1 else feuille_df.columns[0]
    
    # Description souvent colonne C (index 2)
    col_description = feuille_df.columns[2] if len(feuille_df.columns) > 2 else feuille_df.columns[0]
    
    # Qty for est colonne E (index 4)
    col_qty_for = feuille_df.columns[4] if len(feuille_df.columns) > 4 else feuille_df.columns[0]
    
    # Packing list qty souvent colonne F (index 5)
    col_packing = feuille_df.columns[5] if len(feuille_df.columns) > 5 else feuille_df.columns[0]
    
    # Oversent qty souvent colonne ? (à adapter)
    col_oversent = None
    for col in feuille_df.columns:
        if 'oversent' in col.lower() or 'over sent' in col.lower():
            col_oversent = col
            break
    if col_oversent is None:
        col_oversent = feuille_df.columns[7] if len(feuille_df.columns) > 7 else feuille_df.columns[0]
    
    # Moka reply est colonne H (index 7)
    col_moka = feuille_df.columns[7] if len(feuille_df.columns) > 7 else feuille_df.columns[0]
    
    st.write(f"**Structure détectée pour {nom_modele}:**")
    st.write(f"- Part N: '{col_part_n}'")
    st.write(f"- Description: '{col_description}'")
    st.write(f"- Qty for: '{col_qty_for}' (colonne E)")
    st.write(f"- Packing list qty: '{col_packing}' (colonne F)")
    st.write(f"- Remarks: colonne G (valeurs trouvées)")
    st.write(f"- Oversent qty: '{col_oversent}'")
    st.write(f"- Moka reply: '{col_moka}' (colonne H)")
    
    # Filtrer sur Missing et shortage dans la colonne Remarks
    try:
        mask_missing = feuille_df['Remarks'] == 'Missing'
        mask_shortage = feuille_df['Remarks'] == 'shortage'
        df_filtre = feuille_df[mask_missing | mask_shortage]
        
        st.write(f"  - 'Missing' trouvés: {mask_missing.sum()}")
        st.write(f"  - 'shortage' trouvés: {mask_shortage.sum()}")
        st.write(f"  - Total à vérifier: {len(df_filtre)}")
        
    except Exception as e:
        erreurs.append(f"Erreur filtrage: {str(e)}")
        return resultats, erreurs
    
    if df_filtre.empty:
        st.warning(f"Aucune ligne avec 'Missing'/'shortage' dans {nom_modele}")
        return resultats, erreurs
    
    total_lignes = len(df_filtre)
    
    for idx, (index_ligne, ligne) in enumerate(df_filtre.iterrows()):
        if progress_bar and total_lignes > 0:
            progress = (idx + 1) / total_lignes
            progress_bar.progress(progress, text=f"{nom_modele}: {idx+1}/{total_lignes}")
        if status_text:
            status_text.text(f"Traitement: {nom_modele}")
        
        part_n = str(ligne[col_part_n]).strip()
        description = str(ligne[col_description]) if pd.notna(ligne[col_description]) else ""
        qty_for = ligne[col_qty_for]
        packing_qty = ligne[col_packing]
        oversent_frs = ligne[col_oversent]
        nom_fichier_stock = str(ligne[col_moka]).strip()
        
        st.write(f"  → Vérification: {part_n} - {nom_fichier_stock}")
        
        # Chercher le fichier stock
        if nom_fichier_stock not in dict_stocks:
            # Recherche approximative
            fichier_trouve = None
            for nom_fichier in dict_stocks.keys():
                if nom_fichier_stock.replace('.xlsx', '').replace('.xls', '') in nom_fichier:
                    fichier_trouve = nom_fichier
                    break
            if fichier_trouve:
                nom_fichier_stock = fichier_trouve
                st.info(f"  Fichier trouvé: {fichier_trouve}")
            else:
                erreurs.append(f"Fichier '{nom_fichier_stock}' non trouvé")
                resultats.append({
                    'Modèle': nom_modele,
                    'IDL': idl,
                    'Part N': part_n,
                    'Description': description,
                    'Status': '❌ Fichier stock manquant',
                    'Fichier requis': nom_fichier_stock
                })
                continue
        
        df_stock = dict_stocks[nom_fichier_stock]
        
        try:
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            oversent_reel = oversent_stock - qty_for + packing_qty
            
            est_correct = abs(oversent_reel - oversent_frs) < 0.01
            
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Qty for (col E)': qty_for,
                'Packing Qty (col F)': packing_qty,
                'Oversent FRS': oversent_frs,
                'Oversent calculé': round(oversent_reel, 2),
                'Écart': round(oversent_reel - oversent_frs, 2),
                'Status': '✅ Correct' if est_correct else '❌ Incorrect',
                'Fichier stock': nom_fichier_stock,
                'Remarks': ligne['Remarks']
            })
            
            if est_correct:
                st.success(f"    ✅ Correct")
            else:
                st.error(f"    ❌ Incorrect (FRS:{oversent_frs} vs Calculé:{round(oversent_reel,2)})")
            
        except Exception as e:
            st.error(f"    ❌ Erreur: {str(e)}")
            erreurs.append(f"Erreur {part_n}: {str(e)}")
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Status': f'❌ {str(e)[:50]}',
            })
    
    return resultats, erreurs

# Interface Sidebar
with st.sidebar:
    st.header("📂 Fichiers")
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'])
    stock_files = st.file_uploader("Fichiers stocks", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    st.markdown("---")
    st.markdown("### 📌 Structure attendue")
    st.markdown("""
    **Fichier reply:**
    - Colonne G = Remarks ('Missing'/'shortage')
    - Colonne H = Moka reply (nom fichier stock)
    - Colonne E = Qty for
    - Colonne F = Packing list qty
    
    **Fichiers stock:**
    - Colonne A = ODF (IDL)
    - Colonne K = OVERSENT QTY
    """)

# Main
if reply_file and stock_files:
    st.success(f"✅ Reply: {reply_file.name}")
    st.success(f"✅ Stocks: {len(stock_files)} fichiers")
    
    with st.spinner("Chargement..."):
        dict_reply = charger_toutes_les_feuilles_reply(reply_file)
        if dict_reply:
            dict_stocks = charger_stocks_depuis_upload(stock_files)
            
            st.markdown("---")
            st.header("🔑 IDL par modèle")
            
            idl_par_modele = {}
            cols = st.columns(min(3, len(dict_reply)))
            
            for i, nom_modele in enumerate(dict_reply.keys()):
                with cols[i % len(cols)] if cols else st:
                    idl = st.text_input(f"IDL pour {nom_modele}", key=f"idl_{nom_modele}")
                    if idl:
                        idl_par_modele[nom_modele] = idl
            
            if st.button("🚀 VÉRIFIER", type="primary", use_container_width=True):
                if not idl_par_modele:
                    st.error("Saisissez au moins un IDL")
                else:
                    tous_resultats = []
                    toutes_erreurs = []
                    
                    for nom_modele, df_feuille in dict_reply.items():
                        if nom_modele in idl_par_modele:
                            st.subheader(f"📋 {nom_modele}")
                            
                            resultats, erreurs = verifier_modele(
                                df_feuille, nom_modele, idl_par_modele[nom_modele],
                                dict_stocks
                            )
                            tous_resultats.extend(resultats)
                            toutes_erreurs.extend(erreurs)
                    
                    st.markdown("---")
                    st.header("📊 Résultats")
                    
                    if tous_resultats:
                        df_resultats = pd.DataFrame(tous_resultats)
                        
                        col1, col2, col3 = st.columns(3)
                        total = len(df_resultats)
                        corrects = len(df_resultats[df_resultats['Status'] == '✅ Correct'])
                        incorrects = len(df_resultats[df_resultats['Status'] == '❌ Incorrect'])
                        
                        col1.metric("Total", total)
                        col2.metric("✅ Corrects", corrects)
                        col3.metric("❌ Incorrects", incorrects)
                        
                        st.dataframe(df_resultats, use_container_width=True, hide_index=True)
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_resultats.to_excel(writer, sheet_name='Résultats', index=False)
                            if toutes_erreurs:
                                pd.DataFrame({'Erreurs': toutes_erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                        
                        st.download_button("📥 Télécharger Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                    else:
                        st.error("Aucun résultat")
            
elif reply_file and not stock_files:
    st.warning("Ajoutez les fichiers stocks")
elif not reply_file and stock_files:
    st.warning("Ajoutez le fichier reply.xlsx")
else:
    st.info("👈 Chargez les fichiers")
