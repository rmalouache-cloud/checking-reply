import streamlit as st
import pandas as pd
import io

# Configuration
st.set_page_config(
    page_title="Vérification Fournisseur",
    page_icon="✅",
    layout="wide"
)

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
    
    # Filtrer par Part N
    mask = df_stock[col_part].astype(str).str.strip() == str(part_n).strip()
    df_filtered = df_stock[mask].reset_index(drop=True)
    
    if df_filtered.empty:
        raise ValueError(f"Part N non trouvé")
    
    # Chercher IDL
    mask_idl = df_filtered[col_idl].astype(str).str.strip() == str(idl).strip()
    idx = df_filtered[mask_idl].index
    
    if len(idx) == 0:
        raise ValueError(f"IDL non trouvé")
    
    pos = idx[0]
    if pos == 0:
        raise ValueError(f"IDL à la première ligne")
    
    # Ligne précédente
    val = df_filtered.iloc[pos - 1, 10]
    return float(val) if pd.notna(val) else 0.0

# ==================== INTERFACE ====================

st.title("✅ Vérification Réponses Fournisseur")
st.markdown("---")

# Upload
col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 Fichier Reply")
    reply_file = st.file_uploader("reply.xlsx", type=['xlsx', 'xls'])

with col2:
    st.subheader("📦 Fichiers Stock")
    stock_files = st.file_uploader("Fichiers stock", type=['xlsx', 'xls'], accept_multiple_files=True)

st.markdown("---")

if reply_file and stock_files:
    
    # Chargement
    with st.spinner("Chargement..."):
        dict_reply = charger_feuilles_reply(reply_file)
        dict_stocks = charger_stocks(stock_files)
    
    if dict_reply and dict_stocks:
        
        # Aperçu rapide
        st.caption(f"📁 {len(dict_reply)} feuille(s) trouvée(s) : {', '.join(dict_reply.keys())}")
        
        # IDL par modèle
        st.subheader("🔑 IDL par modèle")
        
        idl_par_modele = {}
        cols = st.columns(min(3, len(dict_reply)))
        
        for i, modele in enumerate(dict_reply.keys()):
            with cols[i % len(cols)]:
                idl = st.text_input(f"{modele}", key=f"idl_{modele}", placeholder="IDL")
                if idl:
                    idl_par_modele[modele] = idl
        
        st.markdown("---")
        
        # Bouton vérification
        if st.button("▶️ VÉRIFIER", type="primary", use_container_width=True):
            
            if not idl_par_modele:
                st.warning("Saisissez au moins un IDL")
            else:
                
                resultats = []
                erreurs = []
                
                # Traitement
                for modele, df_feuille in dict_reply.items():
                    if modele not in idl_par_modele:
                        continue
                    
                    st.markdown(f"### {modele}")
                    
                    df_std = extraire_colonnes_reply(df_feuille)
                    if df_std is None:
                        st.error(f"Format incorrect")
                        continue
                    
                    # Filtrer Missing/Shortage
                    df_filtre = df_std[df_std['Remarks'].isin(['Missing', 'shortage'])]
                    
                    if df_filtre.empty:
                        st.info("Aucune ligne Missing/shortage")
                        continue
                    
                    progress = st.progress(0)
                    
                    for i, (_, row) in enumerate(df_filtre.iterrows()):
                        progress.progress((i + 1) / len(df_filtre))
                        
                        part_n = row['Part_N']
                        desc = row['Description'][:40]
                        qty_for = row['Qty_for']
                        packing_qty = row['Packing_qty']
                        oversent_frs = row['Oversent_FRS']
                        moka_file = row['Moka_file']
                        remarks = row['Remarks']
                        idl = idl_par_modele[modele]
                        
                        # Chercher fichier stock
                        stock_file = None
                        for fname in dict_stocks.keys():
                            if moka_file in fname.replace('.xlsx', '').replace('.xls', ''):
                                stock_file = fname
                                break
                        
                        if not stock_file:
                            erreurs.append(f"{part_n}: Fichier {moka_file} non trouvé")
                            resultats.append({
                                'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                'Remarks': remarks, 'Status': '❌ Fichier manquant'
                            })
                            continue
                        
                        try:
                            oversent_stock = get_oversent_stock(dict_stocks[stock_file], part_n, idl)
                            oversent_calc = oversent_stock + packing_qty - qty_for
                            ecart = oversent_calc - oversent_frs
                            correct = abs(ecart) < 0.01
                            
                            if correct:
                                st.success(f"✅ {part_n[:45]}")
                            else:
                                st.error(f"❌ {part_n[:45]} | FRS:{oversent_frs} | Calc:{oversent_calc:.1f} | Écart:{ecart:.1f}")
                            
                            resultats.append({
                                'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                'Remarks': remarks, 'IDL': idl,
                                'Qty for': qty_for, 'Packing Qty': packing_qty,
                                'Oversent Stock': oversent_stock,
                                'Oversent FRS': oversent_frs,
                                'Oversent Calculé': round(oversent_calc, 1),
                                'Écart': round(ecart, 1),
                                'Status': '✅ Correct' if correct else '❌ Incorrect'
                            })
                            
                        except Exception as e:
                            erreurs.append(f"{part_n}: {str(e)}")
                            resultats.append({
                                'Modèle': modele, 'Part N': part_n, 'Description': desc,
                                'Remarks': remarks, 'Status': '❌ Erreur'
                            })
                            st.error(f"❌ {part_n[:45]}: {str(e)[:60]}")
                    
                    progress.empty()
                    st.markdown("---")
                
                # Résumé
                if resultats:
                    st.subheader("📊 Résumé")
                    
                    df_res = pd.DataFrame(resultats)
                    total = len(df_res)
                    corrects = len(df_res[df_res['Status'] == '✅ Correct'])
                    incorrects = len(df_res[df_res['Status'] == '❌ Incorrect'])
                    
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Total", total)
                    c2.metric("✅ Corrects", corrects)
                    c3.metric("❌ Incorrects", incorrects)
                    c4.metric("Taux", f"{corrects/total*100:.0f}%" if total > 0 else "0%")
                    
                    # Tableau
                    st.dataframe(df_res, use_container_width=True, hide_index=True)
                    
                    # Export
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_res.to_excel(writer, sheet_name='Résultats', index=False)
                        if erreurs:
                            pd.DataFrame({'Erreurs': erreurs}).to_excel(writer, sheet_name='Erreurs', index=False)
                    
                    st.download_button("📥 Télécharger Excel", output.getvalue(), "verification.xlsx", use_container_width=True)
                    
                    if incorrects == 0:
                        st.success("✅ Tout est correct !")
                    else:
                        st.warning(f"⚠️ {incorrects} incohérence(s) détectée(s)")

else:
    st.info("👈 Chargez les fichiers pour commencer")

# Formule en bas
st.markdown("---")
st.caption("📐 Formule: Oversent réel = Oversent stock (ligne précédente) + Packing list qty - Qty for")
