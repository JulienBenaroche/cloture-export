import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
import datetime
import threading
import Scraping
import fusion

# Apparence sombre et thème moderne
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")  # "dark-blue", "green", etc.

# === Fonctions générales ===
choix_possibles = [
    "Suivi des imputations non soumises",
    "Suivi du TACE",
    "Suivi des réestimations non soumises",
    "Check imputations"
]


def filtrer_options(event):
    current_text = combo_var.get()
    filtered = [opt for opt in choix_possibles if current_text.lower()
                in opt.lower()]
    combo.configure(values=filtered or choix_possibles)


def log(message):
    print(message)  # pour debug console
    log_box.configure(state="normal")
    log_box.insert("end", f"{message}\n")
    log_box.see("end")
    log_box.configure(state="disabled")


def lancer_script(choix, mois, annee):
    try:
        log("Authentification en attente...")

        Scraping.log = log
        Scraping.lancer_scraping(choix, mois, annee)

        # 💡 Appel du module de fusion si applicable
        log("🔀 Fusion des fichiers si nécessaire...")
        fusion.fusionner(choix, mois, annee)

        if choix == "Suivi des imputations non soumises":
            import suivi_imputation as module
        elif choix == "Suivi du TACE":
            import suivi_tace as module
        elif choix == "Suivi des réestimations non soumises":
            import suivi_reestimations as module
        elif choix == "Check imputations":
            import suivi_checks as module
        else:
            raise ValueError("Choix non reconnu")

        log("Traitement en cours...")
        

        chemin_fichier = module.executer(mois, annee)

        if not chemin_fichier or chemin_fichier == "non":
            log("❌ Aucun fichier généré.")
            messagebox.showerror("Échec", "❌ Aucun fichier généré.")
            return

        # ✅ Ajoute ceci ici
        fusion.fusionner(choix, mois, annee)

        # log("✅ Fichier généré avec succès.")
        messagebox.showinfo("Succès", f"✅ Fichier généré :\n{chemin_fichier}")

    except Exception as e:
        log(f"❌ Erreur : {e}")
        messagebox.showerror("Erreur", f"❌ Erreur :\n{e}")


def executer():
    choix = combo_var.get()
    mois = mois_var.get()
    annee = annee_var.get()

    if choix not in choix_possibles:
        messagebox.showwarning("Attention", "Choix invalide.")
        return

    if choix != "Check imputations" and (not mois or not annee):
        messagebox.showwarning(
            "Attention", "Veuillez renseigner mois et année.")
        return

    threading.Thread(target=lancer_script, args=(
        choix, mois, annee), daemon=True).start()


# === Fenêtre principale ===
root = ctk.CTk()
root.title("Update Clôture")
root.geometry("720x460")
root.resizable(False, False)

# === Cadre principal ===
main_frame = ctk.CTkFrame(root)
main_frame.pack(expand=True, fill="both", padx=20, pady=20)

# === Titre ===
ctk.CTkLabel(main_frame, text="Choisissez une action :",
             font=("Roboto", 18)).pack(pady=(20, 10))

# === Choix d'action ===
combo_var = ctk.StringVar()
combo = ctk.CTkComboBox(main_frame, variable=combo_var,
                        values=choix_possibles, width=300, font=("Roboto", 13))
combo.pack(pady=(0, 10))
combo.bind("<KeyRelease>", filtrer_options)

# === Sélection date ===
now = datetime.datetime.now()
mois_var = ctk.StringVar(value=f"{now.month:02d}")
annee_var = ctk.StringVar(value=str(now.year))

date_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
date_frame.pack(pady=10)

ctk.CTkLabel(date_frame, text="Mois", font=(
    "Roboto", 13)).grid(row=0, column=0, padx=10)
ctk.CTkComboBox(date_frame, width=80, variable=mois_var,
                values=[f"{i:02d}" for i in range(1, 13)],
                font=("Roboto", 13)).grid(row=0, column=1)

ctk.CTkLabel(date_frame, text="Année", font=(
    "Roboto", 13)).grid(row=0, column=2, padx=10)
ctk.CTkComboBox(date_frame, width=100, variable=annee_var,
                values=[str(now.year + i) for i in range(-2, 3)],
                font=("Roboto", 13)).grid(row=0, column=3)

# === Bouton d'exécution ===
ctk.CTkButton(main_frame, text="🚀 Lancer", width=150,
              height=40, command=executer, font=("Roboto", 14, "bold")).pack(pady=25)

# === Zone de log dynamique ===
log_box = ctk.CTkTextbox(main_frame, height=100, width=500,
                         corner_radius=10, font=("Roboto", 12))
log_box.pack(pady=(0, 10))
log_box.configure(state="disabled")

# === Lancement ===
root.mainloop()
