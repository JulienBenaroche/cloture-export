
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk, ImageDraw
import threading
import datetime
import os
import sys
import Scraping
from customtkinter import CTkImage

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

# === PyInstaller-friendly path ===


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


# === Options valides ===
choix_possibles = [
    "Suivi des imputations non soumises",
    "Suivi du TACE",
    "Suivi des réestimations non soumises",
    "Check imputations"
]

llast_value = ""


def filtrer_options(event):
    global last_value
    current_text = combo_var.get()
    if current_text == last_value:
        return
    last_value = current_text

    filtered = [
        option for option in choix_possibles if current_text.lower() in option.lower()]
    combo['values'] = filtered if filtered else choix_possibles


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
            messagebox.showerror(
                "Échec", "❌ Aucun fichier n'a été généré. Vérifiez les données.")
            return

        messagebox.showinfo(
            "Succès", f"✅ Succès !\n\n📄 Fichier : {chemin_fichier}")

    except Exception as e:
        messagebox.showerror("Erreur", f"❌ Une erreur est survenue :\n{e}")


def executer():
    choix = combo_var.get()
    mois = mois_var.get()
    annee = annee_var.get()

    if choix not in choix_possibles:
        messagebox.showwarning(
            "Attention", "Choix invalide. Veuillez sélectionner une option valide.")
        return

    if choix != "Check imputations" and (not mois or not annee):
        messagebox.showwarning(
            "Attention", "Veuillez remplir le mois et l'année.")
        return

    threading.Thread(target=lancer_script, args=(
        choix, mois, annee), daemon=True).start()


# === Fenêtre principale ===
root = ctk.CTk()
root.title("Update clôture details")
root.geometry("520x370")
root.resizable(False, False)

# Centrage universel (portable/fixe/DPI)
root.update_idletasks()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x = int((screen_width - 520) / 2)
y = int((screen_height - 370) / 2)
root.geometry(f"+{x}+{y}")


# === Image de fond avec overlay ===
w, h = 520, 370
background_path = resource_path("background.jpg")
if not os.path.exists(background_path):
    messagebox.showerror(
        "Erreur", f"Image de fond introuvable : {background_path}")
    root.destroy()
    exit()

try:
    resample = Image.Resampling.LANCZOS
except AttributeError:
    resample = Image.LANCZOS

bg_image = Image.open(background_path).resize((w, h), resample)
overlay = Image.new("RGBA", bg_image.size)
draw = ImageDraw.Draw(overlay)
draw.rounded_rectangle([(40, 40), (480, 330)],
                       radius=25, fill=(245, 245, 245, 230))
combined = Image.alpha_composite(bg_image.convert("RGBA"), overlay)
bg_photo = CTkImage(light_image=combined, size=(w, h))
bg_label = ctk.CTkLabel(master=root, image=bg_photo, text="")
bg_label.place(x=0, y=0, relwidth=1, relheight=1)

# === Conteneur principal ===
frame = ctk.CTkFrame(master=root, corner_radius=20, fg_color="#f4f4f4")
frame.place(relx=0.5, rely=0.5, anchor="center")

ctk.CTkLabel(frame, text="Choisissez une action à exécuter :",
             font=("Segoe UI", 15)).pack(pady=(15, 5))

combo_var = ctk.StringVar()
combo = ttk.Combobox(frame, textvariable=combo_var,
                     values=choix_possibles, width=45)
combo.pack(pady=(0, 15))
combo.bind("<KeyRelease>", filtrer_options)

# === Date ===
now = datetime.datetime.now()
mois_var = ctk.StringVar(value=f"{now.month:02d}")
annee_var = ctk.StringVar(value=str(now.year))

date_frame = ctk.CTkFrame(master=frame, fg_color="transparent")
date_frame.pack(pady=5)

ctk.CTkLabel(date_frame, text="Mois :", font=(
    "Segoe UI", 12)).grid(row=0, column=0, padx=10)
ctk.CTkComboBox(date_frame, width=60, variable=mois_var, values=[
                f"{i:02d}" for i in range(1, 13)]).grid(row=0, column=1)

ctk.CTkLabel(date_frame, text="Année :", font=(
    "Segoe UI", 12)).grid(row=0, column=2, padx=10)
ctk.CTkComboBox(date_frame, width=80, variable=annee_var, values=[
                str(now.year + i) for i in range(-2, 3)]).grid(row=0, column=3)

# === Bouton Lancer ===
ctk.CTkButton(frame, text="🚀 Lancer", width=120, height=36, corner_radius=10, font=(
    "Segoe UI", 12, "bold"), command=executer).pack(pady=20)

# Forcer la fenêtre à se calculer complètement
root.update_idletasks()

# Obtenir la vraie taille une fois le layout chargé
width = root.winfo_width()
height = root.winfo_height()

# Recalculer le centrage avec la taille réelle
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")

root.mainloop()
