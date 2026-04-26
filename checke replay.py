import pandas as pd
import os

def charger_toutes_les_feuilles_reply(chemin_reply):
    """Charge toutes les feuilles du fichier reply.xlsx dans un dictionnaire"""
    xlsx = pd.ExcelFile(chemin_reply)
    feuilles = {}
    for sheet_name in xlsx.sheet_names:
        feuilles[sheet_name] = pd.read_excel(chemin_reply, sheet_name=sheet_name)
    return feuilles

def charger_tous_stocks(dossier_stocks):
    """Charge tous les fichiers Excel du dossier stocks"""
    stocks = {}
    for fichier in os.listdir(dossier_stocks):
        if fichier.endswith(('.xlsx', '.xls')):
            chemin = os.path.join(dossier_stocks, fichier)
            stocks[fichier] = pd.read_excel(chemin)
    return stocks

def get_oversent_from_stock(df_stock, part_n, idl):
    """
    Dans le fichier stock filtré par Part N,
    trouve l'IDL dans colonne A et retourne la valeur colonne K de la ligne précédente
    """
    # Filtrer le stock par Part N (chercher dans tout le dataframe)
    # Supposons que la colonne du Part N s'appelle 'Part N' ou similaire
    # À adapter selon le nom exact de la colonne dans vos fichiers stock
    col_part_n = None
    for col in df_stock.columns:
        if 'part' in col.lower() or 'pn' in col.lower():
            col_part_n = col
            break
    
    if col_part_n is None:
        raise ValueError("Colonne Part N non trouvée dans le fichier stock")
    
    df_filtered = df_stock[df_stock[col_part_n].astype(str).str.strip() == str(part_n).strip()]
    
    if df_filtered.empty:
        raise ValueError(f"Part N {part_n} non trouvé dans le fichier stock")
    
    # Chercher la ligne avec l'IDL dans colonne A (ODF)
    idx = df_filtered[df_filtered.iloc[:, 0].astype(str).str.strip() == str(idl).strip()].index
    
    if len(idx) == 0:
        raise ValueError(f"IDL {idl} non trouvé pour le Part N {part_n}")
    
    # Trouver la position dans le dataframe original
    ligne_idl = idx[0]
    
    if ligne_idl == 0:
        raise ValueError(f"L'IDL {idl} est à la première ligne du fichier, pas de ligne précédente")
    
    # Vérifier que la ligne précédente a le même Part N
    ligne_prec = ligne_idl - 1
    part_n_prec = df_stock.iloc[ligne_prec][col_part_n]
    
    if str(part_n_prec).strip() != str(part_n).strip():
        raise ValueError(f"La ligne précédente n'a pas le même Part N ({part_n_prec} != {part_n})")
    
    oversent_precedent = df_stock.iloc[ligne_prec, 10]  # colonne K (index 10)
    return oversent_precedent

def verifier_modele(feuille_df, nom_modele, idl, dict_stocks):
    """Vérifie toutes les lignes d'une feuille (modèle)"""
    resultats = []
    
    # Filtrer sur Missing et shortage
    df_filtre = feuille_df[feuille_df['Remarks'].astype(str).str.strip().isin(['Missing', 'shortage'])]
    
    if df_filtre.empty:
        print(f"  Aucune ligne avec Missing/shortage pour {nom_modele}")
        return resultats
    
    print(f"\n  Traitement du modèle {nom_modele} (IDL: {idl}) - {len(df_filtre)} lignes à vérifier")
    
    for idx, ligne in df_filtre.iterrows():
        part_n = str(ligne['Part N']).strip()
        description = ligne['Description']
        qty_for = ligne['Qty for']
        packing_qty = ligne['Packing list qty']
        oversent_frs = ligne['Oversent qty']
        nom_fichier_stock = str(ligne['Moka reply']).strip()
        
        print(f"    Vérification Part N: {part_n}")

        if nom_fichier_stock not in dict_stocks:
            print(f"      ❌ Fichier stock {nom_fichier_stock} non trouvé")
            continue
        
        df_stock = dict_stocks[nom_fichier_stock]
        
        try:
            oversent_stock = get_oversent_from_stock(df_stock, part_n, idl)
            oversent_reel = oversent_stock - qty_for + packing_qty
            
            # Comparaison avec tolerance
            est_correct = abs(oversent_reel - oversent_frs) < 0.01
            
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Qty for': qty_for,
                'Packing list qty': packing_qty,
                'Oversent FRS': oversent_frs,
                'Oversent calculé': oversent_reel,
                'Écart': oversent_reel - oversent_frs,
                'Correct': 'OUI' if est_correct else 'NON',
                'Fichier stock': nom_fichier_stock
            })
            
            status = "✅" if est_correct else "❌"
            print(f"      {status} FRS={oversent_frs} | Calculé={oversent_reel:.2f} | Écart={oversent_reel - oversent_frs:.2f}")
            
        except Exception as e:
            print(f"      ❌ Erreur: {e}")
            resultats.append({
                'Modèle': nom_modele,
                'IDL utilisé': idl,
                'Part N': part_n,
                'Description': description,
                'Erreur': str(e)
            })
    
    return resultats

def main():
    # Configuration
    chemin_reply = "reply.xlsx"
    dossier_stocks = "stocks_frs"
    
    print("=" * 60)
    print("VÉRIFICATION DES RÉPONSES FOURNISSEUR")
    print("=" * 60)
    
    # 1. Chargement des fichiers
    print("\n📂 Chargement des fichiers...")
    dict_reply = charger_toutes_les_feuilles_reply(chemin_reply)
    print(f"  - {len(dict_reply)} feuilles trouvées dans reply.xlsx : {list(dict_reply.keys())}")
    
    dict_stocks = charger_tous_stocks(dossier_stocks)
    print(f"  - {len(dict_stocks)} fichiers stocks chargés")
    
    # 2. Traitement par modèle (feuille)
    tous_resultats = []
    
    for nom_modele, df_feuille in dict_reply.items():
        print(f"\n{'='*60}")
        print(f"📋 MODÈLE : {nom_modele}")
        print(f"{'='*60}")
        
        # Demander l'IDL pour ce modèle
        idl = input(f"  Entrez l'IDL pour le modèle {nom_modele} : ").strip()
        
        # Vérifier le modèle
        resultats = verifier_modele(df_feuille, nom_modele, idl, dict_stocks)
        tous_resultats.extend(resultats)
    
    # 3. Affichage du rapport final
    print("\n" + "=" * 60)
    print("📊 RAPPORT FINAL")
    print("=" * 60)
    
    if tous_resultats:
        df_resultats = pd.DataFrame(tous_resultats)
        
        # Compter les erreurs
        if 'Correct' in df_resultats.columns:
            nb_ok = len(df_resultats[df_resultats['Correct'] == 'OUI'])
            nb_ko = len(df_resultats[df_resultats['Correct'] == 'NON'])
            print(f"\n✅ Corrects : {nb_ok}")
            print(f"❌ Incorrects : {nb_ko}")
        
        # Sauvegarder
        df_resultats.to_excel("verification_complete.xlsx", index=False)
        print(f"\n💾 Rapport complet sauvegardé dans 'verification_complete.xlsx'")
        
        # Afficher les erreurs
        if 'Correct' in df_resultats.columns:
            erreurs = df_resultats[df_resultats['Correct'] == 'NON']
            if not erreurs.empty:
                print("\n⚠️  DÉTAIL DES ERREURS :")
                print(erreurs[['Modèle', 'Part N', 'Oversent FRS', 'Oversent calculé', 'Écart']].to_string(index=False))
    else:
        print("\n⚠️ Aucun résultat à afficher")

if __name__ == "__main__":
    main()
