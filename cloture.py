""""
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import datetime
"""

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk, ImageDraw
import threading
import datetime
import os

def lancer_script(choix, mois, annee):
    try:
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
background_path = "background.jpg"
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

"""

def lancer_script(choix, mois, annee):
    import os, sys, time, glob, traceback, importlib
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from openpyxl import Workbook

    start_time = time.time()
    sys.stdout.reconfigure(encoding='utf-8')

    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    download_dir = os.path.join(base_path, "extract")
    os.makedirs(download_dir, exist_ok=True)

    chromedriver_path = os.path.join(base_path, "chromedriver-win64", "chromedriver.exe")
    if not os.path.exists(chromedriver_path):
        messagebox.showerror("Erreur", f"chromedriver.exe introuvable.\nChemin : {chromedriver_path}")
        return

    options = Options()
    options.add_argument("--start-maximized")
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")
        print("Ouverture Wavekeeper")

        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, "loginfmt")))
            email_field = driver.find_element(By.NAME, "loginfmt")
            email_field.send_keys("julien.benaroche@wavestone.com")
            driver.find_element(By.ID, "idSIButton9").click()
        except:
            print("Déjà connecté ?")

        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//div[text()='Timesheets']"))
        )

        driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")

        try:
            list_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.oi-view-list"))
            )
            list_btn.click()
        except:
            print("Vue liste déjà active")

        try:
            export_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "o_list_export_xlsx"))
            )
            export_btn.click()
            print("Export lancé")
        except:
            print("Erreur export")
            traceback.print_exc()
            with open("debug_export_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)

        time.sleep(10)
        downloaded_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
        if not downloaded_files:
            wb = Workbook()
            empty_file_path = os.path.join(download_dir, f"test_{mois}_{annee}.xlsx")
            wb.save(empty_file_path)
            print(f"Fichier vide créé : {empty_file_path}")
        else:
            print(f"Fichier téléchargé : {downloaded_files[0]}")

    finally:
        driver.quit()
        duration = time.time() - start_time
        print(f"Script terminé en {round(duration, 2)}s")

        # 📦 Appel du module après exécution Selenium
        try:
            if choix == "Suivi des imputations non soumises":
                import suivi_imputation as module
            elif choix == "Suivi du TACE":
                import suivi_tace as module
            elif choix == "Suivi des réestimations non soumises":
                import suivi_reestimations as module
            else:
                print("❌ Choix non reconnu")
                return
            module.executer(mois, annee)
        except Exception as e:
            print(f"❌ Erreur lors de l'exécution du module complémentaire : {e}")

# --- Interface graphique ---
def executer():
    choix = combo.get()
    mois = mois_var.get()
    annee = annee_var.get()
    if not choix or not mois or not annee:
        messagebox.showwarning("Attention", "Veuillez remplir tous les champs.")
        return
    threading.Thread(target=lancer_script, args=(choix, mois, annee), daemon=True).start()

root = tk.Tk()
root.title("pdate cloture details")
root.geometry("400x250")
root.configure(bg="#f5f5f5")

style = ttk.Style()
style.theme_use('clam')
style.configure("TLabel", font=("Segoe UI", 10), background="#f5f5f5")
style.configure("TButton", font=("Segoe UI", 10), padding=6)
style.configure("TCombobox", font=("Segoe UI", 10))

# Titre
tk.Label(root, text="Choisissez une action à exécuter :", bg="#f5f5f5", font=("Segoe UI", 11, "bold")).pack(pady=(20, 5))

combo = ttk.Combobox(root, state="readonly", values=[
    "Suivi des imputations non soumises",
    "Suivi du TACE",
    "Suivi des réestimations non soumises"
], width=35)
combo.pack()

frame_dates = tk.Frame(root, bg="#f5f5f5")
frame_dates.pack(pady=20)

now = datetime.datetime.now()

# Mois
tk.Label(frame_dates, text="Mois :", bg="#f5f5f5").grid(row=0, column=0, padx=5)
mois_var = tk.StringVar(value=f"{now.month:02d}")
mois_box = ttk.Combobox(frame_dates, textvariable=mois_var, width=5, state="readonly")
mois_box['values'] = [f"{i:02d}" for i in range(1, 13)]
mois_box.grid(row=0, column=1, padx=5)

# Année
tk.Label(frame_dates, text="Année :", bg="#f5f5f5").grid(row=0, column=2, padx=5)
annee_var = tk.StringVar(value=str(now.year))
annee_box = ttk.Combobox(frame_dates, textvariable=annee_var, width=6, state="readonly")
annee_box['values'] = [str(now.year + i) for i in range(-2, 3)]
annee_box.grid(row=0, column=3, padx=5)

# Lancer
ttk.Button(root, text="Lancer", command=executer).pack(pady=10)

root.mainloop()

"""