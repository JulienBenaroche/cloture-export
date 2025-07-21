import os, sys, time, glob
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
import shutil


def attendre_cliquable(driver, by, value, timeout=15):
    return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))

def attendre_telechargement_termine(download_dir, timeout=60):
    print("⏳ Attente de fin de téléchargement...")
    for _ in range(timeout):
        cr_files = glob.glob(os.path.join(download_dir, "*.crdownload"))
        xlsx_files = glob.glob(os.path.join(download_dir, "*.xlsx"))
        if not cr_files and xlsx_files:
            print("✅ Fichier téléchargé :", os.path.basename(xlsx_files[0]))
            return xlsx_files[0]
        time.sleep(1)
    raise TimeoutError("❌ Téléchargement non terminé après 60s.")

def lancer_scraping(choix, mois, annee):
    email = "julien.benaroche@wavestone.com"
    sys.stdout.reconfigure(encoding='utf-8')
    start_time = time.time()

    base_path = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    chromedriver_path = os.path.join(base_path, "chromedriver-win64", "chromedriver.exe")
    if not os.path.exists(chromedriver_path):
        raise FileNotFoundError(f"chromedriver.exe introuvable : {chromedriver_path}")

    dossier_extract = os.path.join(
    os.path.expanduser("~"),
    "Wavestone",
    "WO - CTO - CDM - Clôture",
    "extract"
    )

    os.makedirs(dossier_extract, exist_ok=True)

    prefs = {
        "download.default_directory": dossier_extract,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }

    options = Options()
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)
    options.add_experimental_option("prefs", prefs)

    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=options)

    try:
        print("🌐 Ouverture de Wavekeeper...")
        driver.get("https://wavekeeper.wavestone-app.com/web#cids=1&action=menu")

        try:
            email_field = attendre_cliquable(driver, By.NAME, "loginfmt")
            email_field.clear()
            email_field.send_keys(email)
            attendre_cliquable(driver, By.ID, "idSIButton9").click()
            print("🔐 Email saisi, validation...")
        except:
            print("🔄 Déjà connecté ou champ non affiché.")

        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(),'Timesheets')]"))
        )
        print("✅ Connexion réussie")

        print("➡️ Accès à la page Timesheets (menu_id=216)...")
        driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "nav.o_main_navbar"))
        )

        try:
            print("➡️ Ouverture du menu 'Review'...")
            review_button = attendre_cliquable(driver, By.CSS_SELECTOR, "button.dropdown-toggle[title='Review']")
            driver.execute_script("arguments[0].scrollIntoView(true);", review_button)
            review_button.click()
            print("✅ Bouton 'Review' cliqué.")
        except Exception as e:
            print(f"❌ Erreur sur 'Review' : {e}")
            return

        try:
            print("➡️ Clic sur 'Team Timesheets'...")
            team_link = attendre_cliquable(driver, By.XPATH, "//a[@href='#menu_id=1063&action=1911' and contains(text(),'Team Timesheets')]")
            team_link.click()
            print("✅ Lien 'Team Timesheets' cliqué.")
        except Exception as e:
            print(f"❌ Erreur au clic sur 'Team Timesheets' : {e}")
            return

        print("⏳ Attente de la page Team Timesheets (grille)...")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//th[contains(text(),'Day(s)')]"))
        )
        print("📊 Grille Team Timesheets chargée avec succès.")

        try:
            print("🧹 Suppression du filtre actif...")
            bouton_supprimer_filtre = attendre_cliquable(driver, By.CSS_SELECTOR, "i.o_facet_remove.oi.oi-close.btn.btn-link")
            bouton_supprimer_filtre.click()
            print("✅ Filtre supprimé.")
        except Exception as e:
            print(f"❌ Impossible de supprimer le filtre : {e}")

        try:
            print("⭐ Ouverture du menu 'Favorites'...")
            bouton_favorites = attendre_cliquable(driver, By.XPATH, "//button[.//span[text()='Favorites']]")
            bouton_favorites.click()
            print("✅ Menu 'Favorites' ouvert.")
        except Exception as e:
            print(f"❌ Erreur lors de l'ouverture du menu 'Favorites' : {e}")

        try:
            print("🎯 Sélection du filtre 'CTO non soumises'...")
            filtre_cto = attendre_cliquable(driver, By.XPATH, "//span[contains(text(), 'CTO non soumises') and not(ancestor::i)]")
            filtre_cto.click()
            print("✅ Filtre 'CTO non soumises' sélectionné.")
        except Exception as e:
            print(f"❌ Impossible de sélectionner le filtre : {e}")

        try:
            print("🖱️ Clic humain simulé avec ActionChains sur la case 'Tout sélectionner'...")
            checkbox = attendre_cliquable(driver, By.ID, "checkbox-comp-1")
            ActionChains(driver).move_to_element(checkbox).click().perform()

            print("⏳ Attente que les lignes restent cochées après le clic...")
            WebDriverWait(driver, 5).until(
                lambda d: d.execute_script("""
                    const boxes = document.querySelectorAll('input[type="checkbox"][id^="checkbox-"]:not(#checkbox-comp-1)');
                    return Array.from(boxes).length > 0 && Array.from(boxes).every(b => b.checked);
                """)
            )
            print("✅ Toutes les lignes sont restées sélectionnées ✅")
        except Exception as e:
            print(f"❌ Échec du clic ActionChains ou de la vérification : {e}")


        try:
            print("🔁 Clic sur le lien 'Select all XXX' pour tout sélectionner (même hors écran)...")
            bouton_select_all = attendre_cliquable(driver, By.CSS_SELECTOR, "a.o_list_select_domain")
            ActionChains(driver).move_to_element(bouton_select_all).pause(0.3).click().perform()
            print("✅ Tous les enregistrements correspondants sélectionnés.")
        except Exception as e:
            print(f"❌ Erreur lors du clic sur 'Select all' : {e}")


        try:
            print("⚙️ Clic naturel simulé sur le bouton 'Action'...")
            bouton_action = attendre_cliquable(driver, By.XPATH, "//button[.//span[text()='Action']]")
            
            # Scroll vers le bouton avec JS, attendre un petit délai, puis clic avec ActionChains
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", bouton_action)
            time.sleep(0.5)
            ActionChains(driver).move_to_element(bouton_action).pause(0.5).click().perform()

            print("✅ Bouton 'Action' cliqué de manière réaliste.")
        except Exception as e:
            print(f"❌ Erreur lors du clic sur le bouton Action : {e}")

        try:
            print("📦 Clic sur l'entrée 'Export' dans le menu Action...")
            export_menu_item = attendre_cliquable(driver, By.XPATH, "//span[normalize-space(text())='Export' and contains(@class,'o_menu_item')]")
            export_menu_item.click()
            print("✅ Export déclenché.")
        except Exception as e:
            print(f"❌ Erreur lors du clic sur 'Export' : {e}")


        from selenium.webdriver.support.ui import Select

        try:
            print("📋 Sélection du template 'Extract Taux de présence - NE PAS SUPPRIMER'...")
            select_element = attendre_cliquable(driver, By.CSS_SELECTOR, "select.form-select.ms-4.o_exported_lists_select")
            select = Select(select_element)
            select.select_by_visible_text("Extract Taux de présence - NE PAS SUPPRIMER")
            print("✅ Template sélectionné.")
        except Exception as e:
            print(f"❌ Erreur lors de la sélection du template : {e}")


# test
        from selenium.webdriver.support.ui import Select

        try:
            print("📋 Sélection du template 'CTO - relances' (comportement humain)...")
            select_element = attendre_cliquable(driver, By.CSS_SELECTOR, "select.o_exported_lists_select")
            select = Select(select_element)
            select.select_by_visible_text("CTO - relances")
            print("✅ Template 'CTO - relances' sélectionné.")
        except Exception as e:
            print(f"❌ Erreur lors de la sélection du template : {e}")

        try:
            print("📦 Clic humain sur le bouton 'Export'...")
            bouton_export = attendre_cliquable(driver, By.CSS_SELECTOR, "button.o_select_button.btn.btn-primary")
            actions = ActionChains(driver)
            actions.move_to_element(bouton_export).pause(0.5).click().perform()
            print("✅ Clic sur le bouton Export effectué.")

            # ⏳ Attente du téléchargement
            fichier_telecharge = attendre_telechargement_termine(dossier_extract)
            print(f"📥 Fichier téléchargé : {fichier_telecharge}")

        except Exception as e:
            print(f"❌ Erreur lors du clic sur Export ou téléchargement : {e}")


    except Exception as e:
        print(f"❌ Erreur globale : {e}")

    finally:
        duration = round(time.time() - start_time, 2)
        print(f"✅ Script terminé en {duration}s — Chrome reste ouvert.")
