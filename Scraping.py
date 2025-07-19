# # scraping.py

# import os, sys, time, glob, traceback
# from selenium import webdriver
# from selenium.webdriver.chrome.service import Service
# from selenium.webdriver.chrome.options import Options
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from openpyxl import Workbook

# def attendre_cliquable(driver, by, value, timeout=15):
#     """Attend qu'un élément soit cliquable, sinon lève une erreur"""
#     return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

# def lancer_scraping(choix, mois, annee):
#     email="julien.benaroche@wavestone.com"
#     start_time = time.time()
#     sys.stdout.reconfigure(encoding='utf-8')

#     # Définir le chemin de base
#     if getattr(sys, 'frozen', False):
#         base_path = sys._MEIPASS
#     else:
#         base_path = os.path.dirname(os.path.abspath(__file__))

#     # Préparer dossier extract
#     download_dir = os.path.join(base_path, "extract")
#     os.makedirs(download_dir, exist_ok=True)

#     # Vérifier que le chromedriver est bien là
#     chromedriver_path = os.path.join(base_path, "chromedriver-win64", "chromedriver.exe")
#     if not os.path.exists(chromedriver_path):
#         raise FileNotFoundError(f"chromedriver.exe introuvable.\nChemin : {chromedriver_path}")

#     # Config Chrome
#     options = Options()
#     options.add_argument("--start-maximized")
#     prefs = {
#         "download.default_directory": download_dir,
#         "download.prompt_for_download": False,
#         "directory_upgrade": True,
#         "safebrowsing.enabled": True
#     }
#     options.add_experimental_option("prefs", prefs)

#     # Lancer Chrome
#     service = Service(chromedriver_path)
#     driver = webdriver.Chrome(service=service, options=options)

#     try:
#         print("🌐 Ouverture de Wavekeeper...")
#         driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")

#         # Étape de connexion
#         try:
#             email_field = attendre_cliquable(driver, By.NAME, "loginfmt", timeout=15)
#             email_field.clear()
#             email_field.send_keys(email)

#             next_btn = attendre_cliquable(driver, By.ID, "idSIButton9", timeout=10)
#             next_btn.click()
#             print("🔐 Email saisi, validation...")
#         except Exception:
#             print("🔄 Déjà connecté ou champ non affiché.")

#         # Attend la présence de l’accueil
#         WebDriverWait(driver, 300).until(
#             EC.presence_of_element_located((By.XPATH, "//div[text()='Timesheets']"))
#         )
#         print("✅ Connexion réussie")

#         # Aller à la page des timesheets
#         print("➡️ Accès à la page d'export...")
#         driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")

#         # Activer la vue liste (si possible)
#         try:
#             list_btn = attendre_cliquable(driver, By.CSS_SELECTOR, "button.oi-view-list", timeout=10)
#             list_btn.click()
#             print("👁️ Vue liste activée")
#         except:
#             print("⚠️ Vue liste déjà active ou non trouvée")

#         # Cliquer sur "Exporter"
#         try:
#             export_btn = attendre_cliquable(driver, By.CLASS_NAME, "o_list_export_xlsx", timeout=10)
#             export_btn.click()
#             print("📥 Export déclenché")
#         except:
#             print("❌ Échec export")
#             traceback.print_exc()
#             with open("debug_export_page.html", "w", encoding="utf-8") as f:
#                 f.write(driver.page_source)

#         # Vérifier le fichier téléchargé
#         WebDriverWait(driver, 15).until(lambda d: len(glob.glob(os.path.join(download_dir, "*.xlsx"))) > 0)
#         downloaded_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
#         if not downloaded_files:
#             # Crée un fichier vide si rien n’a été téléchargé
#             wb = Workbook()
#             empty_file_path = os.path.join(download_dir, f"vide_{mois}_{annee}.xlsx")
#             wb.save(empty_file_path)
#             print(f"📄 Aucun fichier téléchargé. Fichier vide généré : {empty_file_path}")
#         else:
#             print(f"📄 Fichier téléchargé : {downloaded_files[0]}")

#     finally:
#         driver.quit()
#         duration = round(time.time() - start_time, 2)
#         print(f"✅ Script terminé en {duration}s")




import os, sys, time, glob, traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook

def attendre_telechargement_termine(download_dir, timeout=60):
    print("⏳ Attente de fin de téléchargement...")
    for _ in range(timeout):
        cr_files = glob.glob(os.path.join(download_dir, "*.crdownload"))
        xlsx_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
        if not cr_files and xlsx_files:
            return xlsx_files[0]
        time.sleep(1)
    raise TimeoutError("⛔️ Temps dépassé : le fichier n'a pas été téléchargé.")

def lancer_scraping(choix, mois, annee):
    email = "julien.benaroche@wavestone.com"
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
        raise FileNotFoundError(f"chromedriver.exe introuvable : {chromedriver_path}")

    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)
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
        print("🌐 Ouverture de Wavekeeper...")
        driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")

        try:
            email_field = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.NAME, "loginfmt"))
            )
            email_field.clear()
            email_field.send_keys(email)
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            ).click()
            print("🔐 Email saisi, validation...")
        except:
            print("🔄 Déjà connecté ou champ non affiché.")

        WebDriverWait(driver, 300).until(
            EC.presence_of_element_located((By.XPATH, "//div[text()='Timesheets']"))
        )
        print("✅ Connexion réussie")

        print("➡️ Accès à la page calendrier.event...")
        driver.get("https://wavekeeper.wavestone-app.com/web#action=240&model=calendar.event&view_type=list&menu_id=170")

        print("👀 Attente du tableau événements")
        try:
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, "//th//span[text()='Attendees']"))
            )
            print("📄 Tableau chargé ✅")
        except Exception:
            print("❌ Le tableau n’a jamais été détecté")
            with open("debug_calendar_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise RuntimeError("🚨 Échec : tableau non détecté")

        print("🚩 Je suis juste avant le bloc de clic Export")

        try:
            print("🔍 Recherche du bouton Export...")
            export_btn = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "button.o_list_export_xlsx"))
            )
            print("🟢 Bouton détecté dans le DOM")

            driver.execute_script("arguments[0].scrollIntoView(true);", export_btn)
            time.sleep(1)

            ActionChains(driver).move_to_element(export_btn).click().perform()
            print("✅ CLIC EXPORT effectué avec succès")

        except Exception as e:
            print("❌ ÉCHEC DU CLIC sur Export")
            traceback.print_exc()
            with open("debug_export_page.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            raise RuntimeError("🚨 Le bouton Export n’a pas pu être cliqué")

        try:
            fichier = attendre_telechargement_termine(download_dir)
            print(f"📄 Fichier téléchargé : {fichier}")
        except:
            wb = Workbook()
            empty_file_path = os.path.join(download_dir, f"vide_{mois}_{annee}.xlsx")
            wb.save(empty_file_path)
            print(f"📄 Aucun fichier détecté. Fichier vide créé : {empty_file_path}")

        time.sleep(10)

    finally:
        print("🧪 Fenêtre Chrome laissée ouverte pour observation (driver.quit() désactivé)")
        driver.quit()
        duration = round(time.time() - start_time, 2)
        print(f"✅ Script terminé en {duration}s")