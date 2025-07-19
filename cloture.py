import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw
import threading  
import datetime
import os
import sys
import Scraping

# === PyInstaller-friendly path ===
def resource_path(relative_path):
    """Retourne le chemin absolu vers le fichier (compatible PyInstaller)"""
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

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
                "Échec",
                "❌ Le programme n’a pas abouti.\n\nAucun fichier n’a été généré ou enregistré.\nMerci de vérifier les données et de réessayer."
            )
            return

        messagebox.showinfo(
            "Succès",
            f"✅ Le programme s’est terminé avec succès !\n\n📄 Fichier généré :\n{chemin_fichier}"
        )

    except Exception as e:
        messagebox.showerror("Erreur", f"❌ Une erreur est survenue :\n{e}")

def executer():
    choix = combo.get()
    mois = mois_var.get()
    annee = annee_var.get()

    if not choix:
        messagebox.showwarning("Attention", "Veuillez choisir une action.")
        return

    if choix != "Check imputations" and (not mois or not annee):
        messagebox.showwarning("Attention", "Veuillez remplir le mois et l’année.")
        return

    threading.Thread(target=lancer_script, args=(choix, mois, annee), daemon=True).start()

# === Interface graphique ===
root = tk.Tk()
root.title("Update clôture details")

# ➕ Centrage
w, h = 450, 300
ws, hs = root.winfo_screenwidth(), root.winfo_screenheight()
x = (ws // 2) - (w // 2)
y = (hs // 2) - (h // 2)
root.geometry(f"{w}x{h}+{x}+{y}")
root.resizable(False, False)

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

bg_image = Image.open(background_path).resize((w, h), resample)

# Crée une image avec un rectangle blanc transparent au centre
overlay = Image.new("RGBA", bg_image.size)
draw = ImageDraw.Draw(overlay)
draw.rounded_rectangle([(50, 50), (400, 250)], radius=15, fill=(255, 255, 255, 210))
combined = Image.alpha_composite(bg_image.convert("RGBA"), overlay)

bg_photo = ImageTk.PhotoImage(combined)

canvas = tk.Canvas(root, width=w, height=h, highlightthickness=0)
canvas.pack()
canvas.create_image(0, 0, anchor="nw", image=bg_photo)

# === Style
style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", font=("Segoe UI", 10), background="#ffffff")
style.configure("TCombobox", font=("Segoe UI", 10))
style.configure("RoundedButton.TButton",
                font=("Segoe UI", 10, "bold"),
                padding=6,
                relief="flat")
style.map("RoundedButton.TButton",
          background=[('active', '#dbeafe')],
          relief=[('pressed', 'flat')],
          foreground=[('disabled', '#888')])

# === Contenu sur le canvas (dans la "fenêtre")
panel = tk.Frame(canvas, bg="#ffffff")

tk.Label(panel, text="Choisissez une action à exécuter :", font=("Segoe UI", 12, "bold"), bg="#ffffff").pack(pady=(10, 5))

combo = ttk.Combobox(panel, state="readonly", width=38, font=("Segoe UI", 10))
combo['values'] = [
    "Suivi des imputations non soumises",
    "Suivi du TACE",
    "Suivi des réestimations non soumises",
    "Check imputations"
]
combo.pack()

# 📅 Date
frame_dates = tk.Frame(panel, bg="#ffffff")
frame_dates.pack(pady=15)

now = datetime.datetime.now()
mois_var = tk.StringVar(value=f"{now.month:02d}")
annee_var = tk.StringVar(value=str(now.year))

tk.Label(frame_dates, text="Mois :", bg="#ffffff").grid(row=0, column=0, padx=5)
ttk.Combobox(frame_dates, textvariable=mois_var, width=5, state="readonly",
             values=[f"{i:02d}" for i in range(1, 13)]).grid(row=0, column=1, padx=5)

tk.Label(frame_dates, text="Année :", bg="#ffffff").grid(row=0, column=2, padx=5)
ttk.Combobox(frame_dates, textvariable=annee_var, width=6, state="readonly",
             values=[str(now.year + i) for i in range(-2, 3)]).grid(row=0, column=3, padx=5)

ttk.Button(panel, text="🚀 Lancer", style="RoundedButton.TButton", command=executer).pack(pady=5)

canvas.create_window(w//2, h//2, window=panel)

root.mainloop()
