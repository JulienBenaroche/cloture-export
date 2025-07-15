import os
import sys
import time
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# S'assurer que la sortie console utilise UTF-8 (Windows)
sys.stdout.reconfigure(encoding='utf-8')

# Détecte automatiquement le dossier du script actuel
script_dir = os.path.dirname(os.path.abspath(__file__))

# Chemin vers chromedriver.exe dans le sous-dossier "chromedriver-win64"
chromedriver_path = os.path.join(script_dir, "chromedriver-win64", "chromedriver.exe")

# Vérifie que le fichier chromedriver.exe existe
if not os.path.exists(chromedriver_path):
    print("❌ ERREUR : chromedriver.exe est introuvable dans le dossier 'chromedriver-win64'.")
    print(f"📁 Chemin attendu : {chromedriver_path}")
    sys.exit(1)

# Configuration des options Chrome
options = Options()
options.add_argument("--start-maximized")
prefs = {"download.prompt_for_download": False}
options.add_experimental_option("prefs", prefs)

# Lancement du navigateur
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service, options=options)

try:
    # 1. Ouverture de la page d'accueil
    driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")
    print("🌐 Ouverture de la page d'accueil Wavekeeper...")
    time.sleep(2)

    # 2. Auto-remplissage email (si champ visible)
    try:
        email_field = driver.find_element(By.NAME, "loginfmt")
        email_field.send_keys("julien.benaroche@wavestone.com")
        driver.find_element(By.ID, "idSIButton9").click()
        print("📧 Email auto-rempli.")
    except:
        print("⚠️ Email déjà saisi ou session déjà active.")

    # 3. Attente de la page d'accueil des modules (icône "Timesheets")
    print("⏳ En attente de validation MFA / chargement du tableau de bord...")
    try:
        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//div[text()='Timesheets']"))
        )
        print("🔓 Connexion détectée. Redirection vers Timesheets.")
    except:
        print("❌ Échec : tableau de bord non détecté.")
        driver.quit()
        sys.exit(1)

    # 4. Aller directement à la page Timesheets
    driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")
    print("📄 Module Timesheets ouvert.")
    time.sleep(5)

    # 5. Forcer vue Liste (si pas déjà activée)
    try:
        list_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.oi-view-list"))
        )
        list_btn.click()
        print("✅ Vue 'Liste' activée.")
        time.sleep(2)
    except Exception:
        print("⚠️ Bouton vue Liste non cliquable ou déjà actif.")

    # 6. Attente que le bouton EXPORT soit cliquable
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
        print("📄 Page HTML sauvegardée pour analyse (debug_export_page.html).")

    time.sleep(10)

finally:
    driver.quit()
