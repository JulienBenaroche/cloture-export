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


        if choix == "Suivi du TACE Timesheets": 
            filtre_a_choisir = f"CTO - {mois}/{annee[-2:]}"  # ex : CTO - 06/25

            print("➡️ Accès à la page Timesheets analysis (menu_id=387)...")
            try:
                driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=387&action=1951")
                print("✅ Page Timesheets analysis chargée avec succès (structure + données visibles).")

            except Exception as e:
                import traceback
                print(f"❌ Erreur lors du chargement complet de la page Timesheets analysis : {type(e).__name__} - {e}")
                traceback.print_exc()
        

            try:
                print("⭐ Rechertes'...")
                bouton_fav = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//button[.//span[text()='Favorites']]"))
                )

                print("🧪 Tentative de clic JS sur le bouton 'Favorites'...")
                driver.execute_script("arguments[0].scrollIntoView(true);", bouton_fav)
                time.sleep(0.3)  # pour stabilité
                driver.execute_script("arguments[0].click();", bouton_fav)

                # attendre que la dropdown apparaisse
                WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "div.o-dropdown--menu.dropdown-menu.d-block"))
                )
                print("✅ Menu déroulant 'Favorites' ouvert avec succès.")
            except Exception as e:
                print(f"❌ Erreur à l'ouverture de 'Favorites' (même en JS) : {e}")


            try:
                print(f"🎯 Attente du filtre '{filtre_a_choisir}' dans la liste...")
                xpath_filtre = f"//span[contains(text(), '{filtre_a_choisir}')]"
                filtre_element = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_filtre))
                )
                filtre_element.click()
                try:
                    print("⏳ Attente que les lignes soient rechargées après application du filtre...")
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("""
                            const lignes = document.querySelectorAll("table.o_list_view tbody tr");
                            return lignes.length > 0;
                        """)
                    )
                    print("✅ Lignes rechargées après filtre.")
                except Exception as e:
                    print(f"❌ Timeout d'attente du rechargement des lignes : {e}")

                print(f"✅ Filtre '{filtre_a_choisir}' sélectionné.")
            except Exception as e:
                print(f"❌ Impossible de sélectionner le filtre '{filtre_a_choisir}' : {e}")
            
            try:
                print("🖱️ Clic humain sur la case 'Tout sélectionner'...")
                checkbox = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "checkbox-comp-1"))
                )
                ActionChains(driver).move_to_element(checkbox).pause(0.5).click().perform()
                print("✅ Case 'Tout sélectionner' cochée.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur la case 'Tout sélectionner' : {e}")

            try:
                print("🔁 Clic sur 'Select all'...")
                bouton_select_all = attendre_cliquable(driver, By.CSS_SELECTOR, "a.o_list_select_domain")
                ActionChains(driver).move_to_element(bouton_select_all).pause(0.3).click().perform()
                print("✅ Tous les enregistrements sélectionnés.")
            except Exception as e:
                print(f"❌ Erreur 'Select all' : {e}")

            try:
                print("⚙️ Clic sur le bouton 'Action'...")
                bouton_action = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Action']]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", bouton_action)
                ActionChains(driver).move_to_element(bouton_action).pause(0.5).click().perform()
                print("✅ Bouton 'Action' cliqué.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur 'Action' : {e}")

            # 📦 Clic humain sur l'entrée 'Export' dans le menu déroulant
            try:
                print("📦 Clic sur 'Export' dans le menu Action...")
                bouton_export = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Export' and contains(@class,'o_menu_item')]"))
                )
                ActionChains(driver).move_to_element(bouton_export).pause(0.4).click().perform()
                print("✅ 'Export' sélectionné.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur 'Export' : {e}")

            from selenium.webdriver.support.ui import Select

            try:
                print("📋 Sélection du template 'CTO_Imputation'...")

                # Sélecteur CSS inchangé (c’est bien la bonne balise <select>)
                select_element = attendre_cliquable(driver, By.CSS_SELECTOR, "select.o_exported_lists_select")

                # Utilisation de Select sur l'élément
                select = Select(select_element)

                # Sélection par le texte visible exact
                select.select_by_visible_text("CTO_Imputation")

                print("✅ Template 'CTO_Imputation' sélectionné.")
            except Exception as e:
                print(f"❌ Erreur lors de la sélection du template : {type(e).__name__} - {e}")

            try:
                print("📥 Clic sur le bouton Export...")
                bouton_export = attendre_cliquable(driver, By.CSS_SELECTOR, "button.o_select_button.btn.btn-primary")
                ActionChains(driver).move_to_element(bouton_export).pause(0.5).click().perform()
                print("✅ Export lancé.")

                fichier_telecharge = attendre_telechargement_termine(dossier_extract)
                print(f"📥 Fichier téléchargé : {fichier_telecharge}")
            except Exception as e:
                print(f"❌ Erreur téléchargement : {e}")

        else:
            print(f"ℹ️ Aucun scraping requis pour le choix : {choix} — Connexion uniquement.")
            return
        

        # ================================
        # 💡 Comportement conditionnel selon le choix 2 eme fichier TACE
        # ================================

        if choix == "Suivi du TACE Overrun":
            print("coucou&")
            filtre_a_choisir = f"CTO - {mois}/{annee[-2:]}"  # ex : CTO - 06/25

            try:
                print("📁 Accès à la page Kanban des projets TACE...")
                driver.get("https://wavekeeper.wavestone-app.com/web#action=505&model=project.project&view_type=kanban&cids=1&menu_id=176")
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.o_kanban_view"))
                )
                print("✅ Page Kanban chargée.")
            except Exception as e:
                print(f"❌ Erreur lors du chargement de la page Kanban : {e}")
                return

            try:
                print("📄 Navigation vers la liste TACE...")
                driver.get("https://wavekeeper.wavestone-app.com/web#action=507&model=exceed.reporting.service_line&view_type=list&cids=1&menu_id=383")

                print("⏳ Attente du tableau OU du message vide...")
                WebDriverWait(driver, 20).until(
                    EC.any_of(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "table.o_list_view")),
                        EC.presence_of_element_located((By.CSS_SELECTOR, ".o_view_nocontent"))
                    )
                )

                print("⏳ Attente que la liste TACE soit bien rendue...")
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.o_list_renderer"))
                )
                print("✅ Composant liste TACE détecté (DOM stable).")


                print("✅ Liste TACE entièrement chargée.")
            except Exception as e:
                print(f"❌ Erreur lors de la navigation vers la liste TACE : {e}")
                return

            try:
                print("⭐ Attente de visibilité du bouton 'Favorites'...")
                bouton_favorites = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Favorites']]"))
                )
                bouton_favorites.click()
                print("✅ Menu 'Favorites' ouvert.")
            except Exception as e:
                print(f"❌ Erreur à l'ouverture de 'Favorites' : {e}")

            try:
                print(f"🎯 Attente du filtre '{filtre_a_choisir}' dans la liste...")
                xpath_filtre = f"//span[contains(text(), '{filtre_a_choisir}')]"
                filtre_element = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, xpath_filtre))
                )
                filtre_element.click()
                try:
                    print("⏳ Attente que les lignes soient rechargées après application du filtre...")
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("""
                            const lignes = document.querySelectorAll("table.o_list_view tbody tr");
                            return lignes.length > 0;
                        """)
                    )
                    print("✅ Lignes rechargées après filtre.")
                except Exception as e:
                    print(f"❌ Timeout d'attente du rechargement des lignes : {e}")

                print(f"✅ Filtre '{filtre_a_choisir}' sélectionné.")
            except Exception as e:
                print(f"❌ Impossible de sélectionner le filtre '{filtre_a_choisir}' : {e}")
            
            try:
                print("🖱️ Clic humain sur la case 'Tout sélectionner'...")
                checkbox = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.ID, "checkbox-comp-1"))
                )
                ActionChains(driver).move_to_element(checkbox).pause(0.5).click().perform()
                print("✅ Case 'Tout sélectionner' cochée.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur la case 'Tout sélectionner' : {e}")


            #  action importante a discuter avec un filtre que doit creer Quentin pour voir si ca fonctionne    # 🔁 Clic humain sur "Select all XXX" 
            # try:
            #     print("🔁 Clic sur le lien 'Select all XXX' (TACE)...")
            #     bouton_select_all = attendre_cliquable(driver, By.CSS_SELECTOR, "a.o_list_select_domain")
            #     ActionChains(driver).move_to_element(bouton_select_all).pause(0.3).click().perform()
            #     print("✅ Tous les enregistrements correspondants sélectionnés.")
            # except Exception as e:
            #     print(f"❌ Erreur lors du clic sur 'Select all' (TACE) : {e}")


            # ⚙️ Clic humain sur le bouton 'Action'

            try:
                print("⚙️ Clic sur le bouton 'Action'...")
                bouton_action = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Action']]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", bouton_action)
                ActionChains(driver).move_to_element(bouton_action).pause(0.5).click().perform()
                print("✅ Bouton 'Action' cliqué.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur 'Action' : {e}")

            # 📦 Clic humain sur l'entrée 'Export' dans le menu déroulant
            try:
                print("📦 Clic sur 'Export' dans le menu Action...")
                bouton_export = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Export' and contains(@class,'o_menu_item')]"))
                )
                ActionChains(driver).move_to_element(bouton_export).pause(0.4).click().perform()
                print("✅ 'Export' sélectionné.")
            except Exception as e:
                print(f"❌ Erreur lors du clic sur 'Export' : {e}")

            from selenium.webdriver.support.ui import Select

            try:
                print("📋 Sélection du template '0-Original view'...")
                
                # Attendre que le <select> soit cliquable
                select_element = attendre_cliquable(driver, By.CSS_SELECTOR, "select.o_exported_lists_select")
                
                # Scroller vers l'élément pour s'assurer qu'il est visible
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", select_element)
                time.sleep(0.3)  # Pause pour le réalisme humain
                
                # Sélection par le texte visible
                select = Select(select_element)
                select.select_by_visible_text("0-Original view")

                print("✅ Template '0-Original view' sélectionné avec succès.")
            except Exception as e:
                print(f"❌ Erreur lors de la sélection du template : {type(e).__name__} - {e}")
      
            try:
                print("📥 Clic sur le bouton Export...")

                # Attente et clic naturel sur le bouton Export
                bouton_export = attendre_cliquable(driver, By.CSS_SELECTOR, "button.o_select_button.btn.btn-primary")
                ActionChains(driver).move_to_element(bouton_export).pause(0.5).click().perform()

                print("✅ Export lancé.")

                # Attente du téléchargement dans le dossier prévu
                fichier_telecharge = attendre_telechargement_termine(dossier_extract)
                print(f"📥 Fichier téléchargé : {fichier_telecharge}")
            except Exception as e:
                print(f"❌ Erreur téléchargement : {e}")

            return


        if choix == "Suivi des imputations non soumises":
            print("➡️ Accès à la page Timesheets (menu_id=216)...")
            driver.get("https://wavekeeper.wavestone-app.com/web#menu_id=216&action=1872")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "nav.o_main_navbar"))
            )

            # ⚙️ Le reste du scraping + export
            from selenium.webdriver.support.ui import Select

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
                print("🖱️ Clic humain sur la case 'Tout sélectionner'...")
                checkbox = attendre_cliquable(driver, By.ID, "checkbox-comp-1")
                ActionChains(driver).move_to_element(checkbox).click().perform()
                WebDriverWait(driver, 5).until(
                    lambda d: d.execute_script("""
                        const boxes = document.querySelectorAll('input[type="checkbox"][id^="checkbox-"]:not(#checkbox-comp-1)');
                        return Array.from(boxes).length > 0 && Array.from(boxes).every(b => b.checked);
                    """)
                )
                print("✅ Toutes les lignes sont restées sélectionnées ✅")
            except Exception as e:
                print(f"❌ Sélection échouée : {e}")

            try:
                print("🔁 Clic sur 'Select all'...")
                bouton_select_all = attendre_cliquable(driver, By.CSS_SELECTOR, "a.o_list_select_domain")
                ActionChains(driver).move_to_element(bouton_select_all).pause(0.3).click().perform()
                print("✅ Tous les enregistrements sélectionnés.")
            except Exception as e:
                print(f"❌ Erreur 'Select all' : {e}")

            try:
                print("⚙️ Clic sur le bouton 'Action'...")
                bouton_action = attendre_cliquable(driver, By.XPATH, "//button[.//span[text()='Action']]")
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", bouton_action)
                time.sleep(0.5)
                ActionChains(driver).move_to_element(bouton_action).pause(0.5).click().perform()
                print("✅ Bouton 'Action' cliqué.")
            except Exception as e:
                print(f"❌ Erreur Action : {e}")

            try:
                print("📦 Clic sur 'Export'...")
                export_menu_item = attendre_cliquable(driver, By.XPATH, "//span[normalize-space(text())='Export' and contains(@class,'o_menu_item')]")
                export_menu_item.click()
                print("✅ Export déclenché.")
            except Exception as e:
                print(f"❌ Erreur Export : {e}")

            try:
                print("📋 Sélection du template 'CTO - relances'...")
                select_element = attendre_cliquable(driver, By.CSS_SELECTOR, "select.o_exported_lists_select")
                select = Select(select_element)
                select.select_by_visible_text("CTO - relances")
                print("✅ Template sélectionné.")
            except Exception as e:
                print(f"❌ Erreur Template : {e}")

            try:
                print("📥 Clic sur le bouton Export...")
                bouton_export = attendre_cliquable(driver, By.CSS_SELECTOR, "button.o_select_button.btn.btn-primary")
                ActionChains(driver).move_to_element(bouton_export).pause(0.5).click().perform()
                print("✅ Export lancé.")

                fichier_telecharge = attendre_telechargement_termine(dossier_extract)
                print(f"📥 Fichier téléchargé : {fichier_telecharge}")
            except Exception as e:
                print(f"❌ Erreur téléchargement : {e}")

        else:
            print(f"ℹ️ Aucun scraping requis pour le choix : {choix} — Connexion uniquement.")
            return

    except Exception as e:
        print(f"❌ Erreur globale : {e}")
    finally:
        duration = round(time.time() - start_time, 2)
        print(f"✅ Script terminé en {duration}s — Chrome reste ouvert.")