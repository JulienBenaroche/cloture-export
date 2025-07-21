
import os
import glob
import math
import shutil
import re
from datetime import datetime
from openpyxl import load_workbook

def get_x_fichiers_recents(dossier, pattern="*.xlsx", limit=2):
    fichiers = glob.glob(os.path.join(dossier, pattern))
    fichiers_trie = sorted(fichiers, key=os.path.getmtime, reverse=True)
    return fichiers_trie[:limit]

def arrondir_heure_par_15min(dt):
    heure = dt.hour
    minute = dt.minute
    quart = math.ceil(minute / 15) * 15
    if quart == 60:
        heure += 1
        quart = 0
    return f"{heure:02d}{quart:02d}"

def fusionner(choix, mois, annee):
    # ✅ Nouveau chemin persistant pour le dossier 'extract'
    dossier_extract = os.path.join(
        os.path.expanduser("~"),
        "Wavestone",
        "WO - CTO - CDM - Clôture",
        "extract"
    )
    print("🔍 Utilisateur courant :", os.getlogin())
    print("📂 Home path utilisé :", os.path.expanduser("~"))
    os.makedirs(dossier_extract, exist_ok=True)
    print(f"📁 Dossier extract utilisé : {dossier_extract}")


    # if choix == "Suivi du TACE Timesheets":
    # if choix == "Suivi du TACE Overrun":
    # if choix == "Check imputations":

    if choix == "Suivi des imputations non soumises":
        fichiers = get_x_fichiers_recents(dossier_extract, "*.xlsx", limit=2)
        if len(fichiers) < 2:
            print("❌ Moins de 2 fichiers trouvés pour la fusion.")
            return

        fichier_1, fichier_2 = fichiers
        print(f"✅ Fusion de :\n- {os.path.basename(fichier_1)}\n- {os.path.basename(fichier_2)}")

        if "Suivi des imputations" in fichier_1:
            fichier_dest = fichier_1
            fichier_source = fichier_2
        else:
            fichier_dest = fichier_2
            fichier_source = fichier_1

        wb_dest = load_workbook(fichier_dest)
        ws_dest = wb_dest["Imputations non soumises"]

        wb_src = load_workbook(fichier_source)
        ws_src = wb_src["Sheet1"]

        row_dest = 2
        for row in ws_src.iter_rows(min_row=2, max_col=7):
            for col_index, cell in enumerate(row, start=1):
                ws_dest.cell(row=row_dest, column=col_index).value = cell.value
            row_dest += 1

        now = datetime.now()
        nouvelle_heure = arrondir_heure_par_15min(now)
        ancien_nom = os.path.basename(fichier_dest)
        nouveau_nom = re.sub(r'_\d{4}_', f'_{nouvelle_heure}_', ancien_nom)
        chemin_nouveau = os.path.join(dossier_extract, nouveau_nom)

        wb_dest.save(chemin_nouveau)
        wb_dest.close()
        wb_src.close()

        print(f"Ligne insérée : {row_dest - 1}")
        print(f"✅ Fusion terminée dans le fichier : {nouveau_nom}")
        print("✅ Fichier généré avec succès.")

        # 📤 Copier vers le sous-dossier de clôture
        sous_dossier = f"{annee[-2:]}-{mois}"
        dossier_sharepoint = os.path.join(
            os.path.expanduser("~"),
            "Wavestone",
            "WO - CTO - CDM - Clôture",
            sous_dossier,
            "Fichiers suivi"
        )
        os.makedirs(dossier_sharepoint, exist_ok=True)

        chemin_final = os.path.join(dossier_sharepoint, nouveau_nom)
        shutil.copy2(chemin_nouveau, chemin_final)
        print(f"📤 Copie réussie dans SharePoint :\n{chemin_final}")

        # 🧹 Supprimer les anciens fichiers "Suivi des imputations non soumises" dans le sous-dossier (sauf le nouveau)
        for fichier in os.listdir(dossier_sharepoint):
            chemin_fichier = os.path.join(dossier_sharepoint, fichier)
            if (
                fichier.startswith(f"{annee}{mois}- Suivi des imputations non soumises")
                and fichier != nouveau_nom
                and os.path.isfile(chemin_fichier)
            ):
                os.remove(chemin_fichier)
                print(f"🗑️ Ancien fichier supprimé : {fichier}")

        # 🧼 Vider le dossier "extract"
        for fichier_temp in os.listdir(dossier_extract):
            chemin_temp = os.path.join(dossier_extract, fichier_temp)
            if os.path.isfile(chemin_temp):
                os.remove(chemin_temp)
        print("🧼 Dossier 'extract' vidé.")

         # 🧹 Vider également le dossier extract du SharePoint (hors sous-dossier)
        dossier_extract_sharepoint = os.path.join(
            os.path.expanduser("~"),
            "Wavestone",
            "WO - CTO - CDM - Clôture",
            "extract"
        )
        for fichier in os.listdir(dossier_extract_sharepoint):
            chemin_fichier = os.path.join(dossier_extract_sharepoint, fichier)
            if os.path.isfile(chemin_fichier):
                os.remove(chemin_fichier)
                print(f"🗑️ Fichier supprimé dans SharePoint/extract : {fichier}")
