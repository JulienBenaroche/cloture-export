import os
import glob
from datetime import datetime
import shutil

def executer(mois, annee):
    print(f"✅ Vous avez choisi : Suivi des réestimations non soumises ({mois}/{annee})")

    # 🔁 Récupère le chemin de base de l'utilisateur
    base_path = os.path.expanduser("~")

    # 📁 Construit le chemin vers le bon dossier partagé OneDrive
    sous_dossier = f"{annee[-2:]}-{mois}"  # Ex : "25-06" pour juin 2025
    dossier = os.path.join(
        base_path,
        "Wavestone",
        "WO - CTO - CDM - Clôture",
        sous_dossier,
        "Fichiers suivi"
    )

    print(f"📁 Chemin de recherche : {dossier}")

    # ❗ Vérifie si le dossier existe
    if not os.path.exists(dossier):
        print(f"❌ Dossier introuvable : {dossier}")
        return

    # 🔍 Motif des fichiers à trouver
    pattern = os.path.join(
        dossier,
        f"{annee}{mois} - Suivi des réestimations non soumises - extract *.xlsx"
    )
    print(f"🔍 Recherche avec motif : {pattern}")

    # 🔎 Cherche tous les fichiers correspondant au motif
    fichiers = glob.glob(pattern)

    if not fichiers:
        print("❌ Aucun fichier trouvé avec ce motif.")
        return

    # 🕒 Fonction pour extraire la date/heure depuis le nom du fichier
    def extraire_datetime(fichier):
        try:
            nom = os.path.basename(fichier)
            partie_date = nom.split("extract ")[1].split("_")[0] + nom.split("extract ")[1].split("_")[1]
            return datetime.strptime(partie_date, "%Y%m%d%H%M")
        except Exception as e:
            print(f"❌ Erreur parsing {fichier} : {e}")
            return datetime.min

    # 📋 Trie les fichiers du plus récent au plus ancien
    fichiers_trie = sorted(fichiers, key=extraire_datetime, reverse=True)

    # ✅ Récupère le plus récent
    fichier_plus_recent = fichiers_trie[0]

    # 🖨️ Affiche le nom et le chemin
    print(f"📄 Fichier le plus récent trouvé : {os.path.basename(fichier_plus_recent)}")
    print(f"📍 Chemin complet : {fichier_plus_recent}")

    # 📁 Prépare le dossier 'extract' dans le projet
    projet_racine = os.path.dirname(os.path.abspath(__file__))
    dossier_extract = os.path.join(projet_racine, "extract")
    os.makedirs(dossier_extract, exist_ok=True)

    # 📤 Copie le fichier
    nom_fichier = os.path.basename(fichier_plus_recent)
    chemin_copie = os.path.join(dossier_extract, nom_fichier)
    shutil.copy2(fichier_plus_recent, chemin_copie)

    print(f"✅ Copie locale enregistrée dans : {chemin_copie}")
    
    return chemin_copie