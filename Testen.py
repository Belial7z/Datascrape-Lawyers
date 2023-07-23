import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

options = Options()
options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

driver.get("https://bravsearch.bea-brak.de/bravsearch/index.brak")
time.sleep(2)

# Dropdown-Tabelle öffnen
dropdown_button = driver.find_element(By.ID, "searchForm:txtSpecialization_label")
dropdown_button.click()

# Option auswählen
option = driver.find_element(By.XPATH, "//li[contains(text(), 'Arbeitsrecht')]")
option.click()

# Lade Postleitzahlen aus Excel-Datei
excel_data = pd.read_excel("PLZ.xlsx")
postleitzahlen = excel_data["Postleitzahlen"].tolist()

# Ersten Eintrag einfügen und Suche starten
plz_input = driver.find_element(By.ID, "searchForm:txtPostal")
plz_input.send_keys(str(postleitzahlen[0]))

suche_starten_button = driver.find_element(By.ID, "searchForm:cmdSearch")
suche_starten_button.click()
time.sleep(1)

# Ergebnisse speichern
data = []

# Anwälte durchgehen
i = 0
max_increment = 40  # Maximalinkrement für i

while i < max_increment:
    try:
        # Element "Info" klicken
        info_element = driver.find_element(By.XPATH, f"//a[@id='resultForm:dlResultList:{i}:j_idt208']")
        info_element.click()
        time.sleep(1)

        # Namen extrahieren, falls vorhanden
        try:
            name_element = driver.find_element(By.XPATH, "//div[@id='resultDetailForm:tabPersonal:j_idt306:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            Name = name_element.text.strip()
        except:
            Name = ""

        # Adresse extrahieren, falls vorhanden
        try:
            adresse_element = driver.find_element(By.XPATH,
                                                  "//div[@id='resultDetailForm:tabPersonal:j_idt352:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            Adresse = adresse_element.text.strip()
        except:
            Adresse = ""

        # Kanzlei extrahieren, falls vorhanden
        try:
            kanzlei_element = driver.find_element(By.XPATH,
                                                  "//div[@id='resultDetailForm:tabPersonal:j_idt345:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            Kanzlei = kanzlei_element.text.strip()
        except:
            Kanzlei = ""

        # E-Mail-Adresse extrahieren, falls vorhanden
        try:
            email_element = driver.find_element(By.XPATH, "//div[@id='resultDetailForm:tabPersonal:j_idt388:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            email = email_element.text.strip()
        except:
            email = ""

        try:
            anrede_element = driver.find_element(By.XPATH,
                                               "//div[@id='resultDetailForm:tabPersonal:j_idt265:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            Anrede = anrede_element.text.strip()
        except:
            Anrede = ""

        # Telefonnummer extrahieren, falls vorhanden
        try:
            telefon_element = driver.find_element(By.XPATH, "//div[@id='resultDetailForm:tabPersonal:j_idt367:textEntry']//div[@class='cssColResultDetailText cssColResultDetailTextLine']")
            Telefon = telefon_element.text.strip()
        except:
            Telefon = ""


        # Daten in Liste speichern, falls mindestens ein Wert vorhanden ist
        if Name or email or Telefon or Kanzlei or Adresse or Anrede:
            data.append({'Kanzlei': Kanzlei, 'Adresse': Adresse, 'Name': Name, 'Telefon': Telefon, 'E-Mail': email, 'Anrede':Anrede,})

        # Fenster schließen
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(1)

        i += 1

    except:
        break

# Ergebnisse in eine Excel-Datei speichern
df = pd.DataFrame(data)
df.to_excel("ergebnis.xlsx", index=False)

# Browser schließen
driver.quit()
