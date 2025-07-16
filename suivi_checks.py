import os
import glob
from datetime import datetime
import shutil

def executer(mois, annee):
    print(f"\n✅ Vous avez choisi : Check imputations ({mois}/{annee})\n")

    # 🔁 Chemin de base utilisateur
    base_path = os.path.expanduser("~")

    # 📁 Dossier du mois, ex : "25-06" pour juin 2025
    sous_dossier = f"{annee[-2:]}-{mois}"
    dossier = os.path.join(
        base_path,
        "Wavestone",
        "WO - CTO - CDM - Clôture",
        sous_dossier
    )

    print(f"📁 Chemin de recherche :\n{dossier}\n")

    # ❗ Vérifie si le dossier existe
    if not os.path.exists(dossier):
        print(f"❌ Dossier introuvable : {dossier}")
        return

    # 🔍 Recherche des fichiers "Check imputations"
    pattern = os.path.join(dossier, "Check imputations*.xls*")
    print(f"🔍 Recherche avec le motif :\n{pattern}\n")

    fichiers = glob.glob(pattern)

    if not fichiers:
        print("❌ Aucun fichier 'Check imputations' trouvé.")
        return

    print(f"🔢 {len(fichiers)} fichier(s) trouvé(s).\n")

    # 📋 Trie par date de modification
    fichiers_trie = sorted(fichiers, key=os.path.getmtime, reverse=True)
    fichier_plus_recent = fichiers_trie[0]

    print(f"📄 Fichier sélectionné :\n{os.path.basename(fichier_plus_recent)}")
    print(f"📍 Chemin complet :\n{fichier_plus_recent}\n")

    # 📁 Prépare le dossier local 'extract'
    projet_racine = os.path.dirname(os.path.abspath(__file__))
    dossier_extract = os.path.join(projet_racine, "extract")
    os.makedirs(dossier_extract, exist_ok=True)

    # 📤 Copie du fichier
    nom_fichier = os.path.basename(fichier_plus_recent)
    chemin_copie = os.path.join(dossier_extract, nom_fichier)
    shutil.copy2(fichier_plus_recent, chemin_copie)

    print(f"✅ Copie locale enregistrée dans :\n{chemin_copie}\n")
    
    return chemin_copie