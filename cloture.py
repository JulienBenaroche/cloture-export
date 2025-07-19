import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk, ImageDraw
import threading
import datetime
import os
import sys
import Scraping
from customtkinter import CTkImage

# Apparence générale
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

# === PyInstaller-friendly path ===
def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# === Options disponibles ===
choix_possibles = [
    "Suivi des imputations non soumises",
    "Suivi du TACE",
    "Suivi des réestimations non soumises",
    "Check imputations"
]

# === Fonctions ===
def filtrer_options(event):
    current_text = combo_var.get()
    filtered = [option for option in choix_possibles if current_text.lower() in option.lower()]
    combo.configure(values=filtered if filtered else choix_possibles)

def lancer_script(choix, mois, annee):
    try:
        Scraping.lancer_scraping(choix, mois, annee)

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

        chemin_fichier = module.executer(mois, annee)

        if not chemin_fichier or chemin_fichier == "non":
            messagebox.showerror("Échec", "❌ Aucun fichier généré. Vérifiez les données.")
            return

        messagebox.showinfo("Succès", f"✅ Succès !\n\n📄 Fichier : {chemin_fichier}")

    except Exception as e:
        messagebox.showerror("Erreur", f"❌ Une erreur est survenue :\n{e}")

def executer():
    choix = combo_var.get()
    mois = mois_var.get()
    annee = annee_var.get()

    if choix not in choix_possibles:
        messagebox.showwarning("Attention", "Choix invalide. Veuillez sélectionner une option valide.")
        return

    if choix != "Check imputations" and (not mois or not annee):
        messagebox.showwarning("Attention", "Veuillez remplir le mois et l'année.")
        return

    threading.Thread(target=lancer_script, args=(choix, mois, annee), daemon=True).start()

# === Fenêtre principale ===
root = ctk.CTk()
root.title("Update clôture details")

# Taille adaptative
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
app_width = int(screen_width * 0.4)
app_height = int(screen_height * 0.45)
root.geometry(f"{app_width}x{app_height}")
x = (screen_width - app_width) // 2
y = (screen_height - app_height) // 2
root.geometry(f"{app_width}x{app_height}+{x}+{y}")
root.minsize(520, 370)

# === Image de fond ===
background_path = resource_path("background.jpg")
if not os.path.exists(background_path):
    messagebox.showerror("Erreur", f"Image de fond introuvable : {background_path}")
    root.destroy()
    exit()

try:
    resample = Image.Resampling.LANCZOS
except AttributeError:
    resample = Image.LANCZOS

bg_image = Image.open(background_path).resize((app_width, app_height), resample)
overlay = Image.new("RGBA", bg_image.size)
draw = ImageDraw.Draw(overlay)
draw.rounded_rectangle([(app_width*0.05, app_height*0.1), (app_width*0.95, app_height*0.9)], radius=25, fill=(245, 245, 245, 230))
combined = Image.alpha_composite(bg_image.convert("RGBA"), overlay)
bg_photo = CTkImage(light_image=combined, size=(app_width, app_height))
bg_label = ctk.CTkLabel(master=root, image=bg_photo, text="")
bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# === Conteneur principal ===
frame = ctk.CTkFrame(master=root, corner_radius=20, fg_color="#f4f4f4")
frame.place(relx=0.5, rely=0.5, anchor="center")

ctk.CTkLabel(frame, text="Choisissez une action à exécuter :", font=("Segoe UI", 16)).pack(pady=(20, 8))

combo_var = ctk.StringVar()
combo = ctk.CTkComboBox(frame, variable=combo_var, values=choix_possibles, width=app_width * 0.5)
combo.pack(pady=(0, 15))
combo.bind("<KeyRelease>", filtrer_options)

# === Choix de date ===
now = datetime.datetime.now()
mois_var = ctk.StringVar(value=f"{now.month:02d}")
annee_var = ctk.StringVar(value=str(now.year))

date_frame = ctk.CTkFrame(master=frame, fg_color="transparent")
date_frame.pack(pady=5)

ctk.CTkLabel(date_frame, text="Mois :", font=("Segoe UI", 13)).grid(row=0, column=0, padx=10)
ctk.CTkComboBox(date_frame, width=70, variable=mois_var, values=[f"{i:02d}" for i in range(1, 13)]).grid(row=0, column=1)

ctk.CTkLabel(date_frame, text="Année :", font=("Segoe UI", 13)).grid(row=0, column=2, padx=10)
ctk.CTkComboBox(date_frame, width=90, variable=annee_var, values=[str(now.year + i) for i in range(-2, 3)]).grid(row=0, column=3)

# === Bouton ===
ctk.CTkButton(frame, text="🚀 Lancer", width=140, height=38, font=("Segoe UI", 13, "bold"), command=executer).pack(pady=25)

root.mainloop()
