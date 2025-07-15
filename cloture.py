import os
import sys
import time
import glob
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook

start_time = time.time()  # Chronomètre de performance

sys.stdout.reconfigure(encoding='utf-8')

# Dossier du script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Chemin vers chromedriver
chromedriver_path = os.path.join(script_dir, "chromedriver-win64", "chromedriver.exe")
if not os.path.exists(chromedriver_path):
    print("❌ ERREUR : chromedriver.exe introuvable.")
    print(f"📁 Chemin attendu : {chromedriver_path}")
    sys.exit(1)

# 📁 Dossier de téléchargement : ./extract
download_dir = os.path.join(script_dir, "extract")
os.makedirs(download_dir, exist_ok=True)

# Configuration du navigateur
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
    # 1. Accès à la page d’accueil
    driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")
    print("🌐 Ouverture de la page d'accueil Wavekeeper...")

    # 2. Auto-remplissage email
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.NAME, "loginfmt")))
        email_field = driver.find_element(By.NAME, "loginfmt")
        email_field.send_keys("julien.benaroche@wavestone.com")
        driver.find_element(By.ID, "idSIButton9").click()
        print("📧 Email auto-rempli.")
    except:
        print("⚠️ Champ email non trouvé : peut-être déjà connecté ou chargement lent.")

    # 3. Attente de connexion (module Timesheets visible)
    print("⏳ En attente de validation MFA / chargement tableau de bord...")
    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//div[text()='Timesheets']"))
        )
        print("🔓 Connexion détectée.")
    except:
        print("❌ Échec : tableau de bord non détecté.")
        driver.quit()
        sys.exit(1)

    # 4. Redirection vers Timesheets
    driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")
    print("📄 Module Timesheets ouvert.")

    # 5. Activer vue Liste
    try:
        list_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.oi-view-list"))
        )
        list_btn.click()
        print("✅ Vue 'Liste' activée.")
    except:
        print("⚠️ Vue Liste non cliquable ou déjà active.")

    # 6. Cliquer sur Export
    try:
        export_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "o_list_export_xlsx"))
        )
        export_btn.click()
        print("✅ Export lancé.")
    except Exception as e:
        print("❌ Impossible de cliquer sur Export.")
        traceback.print_exc()
        with open("debug_export_page.html", "w", encoding="utf-8") as f:
            f.write(driver.page_source)
        print("📄 Page HTML sauvegardée (debug_export_page.html).")

    # 7. Attente du lancement du téléchargement
    try:
        WebDriverWait(driver, 30).until(lambda d: "xlsx" in d.page_source.lower())
        print("📥 Téléchargement enclenché.")
    except:
        print("ℹ️ Fin du script sans confirmation de téléchargement visible.")

        # Vérifie s'il y a un .xlsx dans le dossier
        downloaded_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
        if not downloaded_files:
            # Créer un fichier Excel vide
            wb = Workbook()
            empty_file_path = os.path.join(download_dir, "test.xlsx")
            wb.save(empty_file_path)
            print(f"📄 Aucun fichier téléchargé : fichier vide créé → {empty_file_path}")

finally:
    driver.quit()
    duration = time.time() - start_time
    print(f"⏱️ Script terminé en {round(duration, 2)} secondes.")
