import os
import glob
from datetime import datetime
import shutil

def executer(mois, annee):
    print(f"✅ Vous avez choisi : Suivi du TACE ({mois}/{annee})")

    # 🔁 Répertoire OneDrive
    base_path = os.path.expanduser("~")
    sous_dossier = f"{annee[-2:]}-{mois}"
    dossier = os.path.join(base_path, "Wavestone", "WO - CTO - CDM - Clôture", sous_dossier)

    print(f"📁 Chemin de recherche : {dossier}")

    if not os.path.exists(dossier):
        print(f"❌ Dossier introuvable : {dossier}")
        return

    # 🔍 Chercher les fichiers
    pattern = os.path.join(dossier, "* - Analyse des imputations - extract *.xlsx")
    print(f"🔍 Recherche avec motif : {pattern}")
    fichiers = glob.glob(pattern)

    if not fichiers:
        print("❌ Aucun fichier trouvé avec ce motif.")
        return

    # 🕒 Trier par date/heure dans le nom
    def extraire_datetime(fichier):
        try:
            nom = os.path.basename(fichier)
            date_heure_str = nom.split("extract ")[1].split("_")[0] + nom.split("extract ")[1].split("_")[1]
            return datetime.strptime(date_heure_str, "%Y%m%d%H%M")
        except Exception as e:
            print(f"❌ Erreur parsing {fichier} : {e}")
            return datetime.min

    fichiers_trie = sorted(fichiers, key=extraire_datetime, reverse=True)
    fichier_plus_recent = fichiers_trie[0]
    print(f"📄 Fichier le plus récent trouvé : {os.path.basename(fichier_plus_recent)}")
    print(f"📍 Chemin complet : {fichier_plus_recent}")

    nom_fichier = os.path.basename(fichier_plus_recent)

    # 📁 Copier uniquement dans SharePoint/extract
    dossier_sharepoint_extract = os.path.join(base_path, "Wavestone", "WO - CTO - CDM - Clôture", "extract")
    os.makedirs(dossier_sharepoint_extract, exist_ok=True)
    chemin_copie_sharepoint = os.path.join(dossier_sharepoint_extract, nom_fichier)
    shutil.copy2(fichier_plus_recent, chemin_copie_sharepoint)
    print(f"✅ Copie enregistrée dans : {chemin_copie_sharepoint}")

    return chemin_copie_sharepoint

# import os
# import glob
# from datetime import datetime

# def executer(mois, annee):
#     print(f"✅ Vous avez choisi : Suivi du TACE ({mois}/{annee})")

#     # 🔁 Récupère le chemin de base de l'utilisateur (ex: C:\\Users\\TON_USER)
#     base_path = os.path.expanduser("~")

#     # 📁 Construit le chemin vers le dossier partagé OneDrive
#     sous_dossier = f"{annee[-2:]}-{mois}"  # Ex : "25-06" pour juin 2025
#     dossier = os.path.join(base_path, "Wavestone", "WO - CTO - CDM - Clôture", sous_dossier)

#     # ❗ Vérifie si le dossier existe
#     if not os.path.exists(dossier):
#         print(f"❌ Dossier introuvable : {dossier}")
#         return

#     # 🔍 Construit le motif de recherche avec glob pour trouver les bons fichiers .xlsx
#     pattern = os.path.join(dossier, f"* - Analyse des imputations - extract *.xlsx")
#     print(f"🔍 Recherche avec motif : {pattern}")

#     # 🔎 Cherche tous les fichiers correspondant au motif
#     fichiers = glob.glob(pattern)

#     # ❌ Aucun fichier ne correspond
#     if not fichiers:
#         print("❌ Aucun fichier trouvé avec ce motif.")
#         return

#     # 🕒 Fonction pour extraire la date et l'heure à partir du nom de fichier
#     def extraire_datetime(fichier):
#         try:
#             nom = os.path.basename(fichier)  # Garde uniquement le nom du fichier
#             # Extrait la chaîne de date et heure (ex: 20240628_0920) depuis le nom
#             date_heure_str = nom.split("extract ")[1].split("_")[0] + nom.split("extract ")[1].split("_")[1]
#             # Transforme cette chaîne en objet datetime
#             return datetime.strptime(date_heure_str, "%Y%m%d%H%M")
#         except Exception as e:
#             print(f"❌ Erreur parsing {fichier} : {e}")
#             return datetime.min  # Valeur très ancienne pour ne pas gêner le tri

#     # 📋 Trie les fichiers trouvés du plus récent au plus ancien
#     fichiers_trie = sorted(fichiers, key=extraire_datetime, reverse=True)

#     # ✅ Récupère le fichier le plus récent
#     fichier_plus_recent = fichiers_trie[0]

#     # 🖨️ Affiche le nom et le chemin du fichier retenu
#     print(f"📄 Fichier le plus récent trouvé : {os.path.basename(fichier_plus_recent)}")
#     print(f"📍 Chemin complet : {fichier_plus_recent}")


#     import shutil

#     # 📁 Détermine le chemin du dossier 'extract' à la racine du projet
#     projet_racine = os.path.dirname(os.path.abspath(__file__))
#     dossier_extract = os.path.join(projet_racine, "extract")
#     os.makedirs(dossier_extract, exist_ok=True)  # Crée le dossier s'il n'existe pas

#     # 📝 Nom de fichier et destination
#     nom_fichier = os.path.basename(fichier_plus_recent)
#     chemin_copie = os.path.join(dossier_extract, nom_fichier)

#     # 📤 Copie du fichier dans le dossier 'extract'
#     shutil.copy2(fichier_plus_recent, chemin_copie)

#     print(f"✅ Copie locale enregistrée dans : {chemin_copie}")

#     return chemin_copie
# # https://digiplace.sharepoint.com/:x:/r/sites/WOP-CTO-CDM/_layouts/15/Doc.aspx?sourcedoc=%7B68781610-492B-4F21-97BF-A95BB9AB7C80%7D&file=202506-%20Suivi%20des%20imputations%20non%20soumises%20-%20extract%2020252406_0900_QIL.xlsx&action=default&mobileredirect=true