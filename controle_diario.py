from oauth2client.service_account import ServiceAccountCredentials
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium import webdriver
from bs4 import BeautifulSoup
from datetime import datetime
from http import client
import gspread

import logging
import time
import sys
import re
import os

start_time = time.perf_counter()

logging.basicConfig(
    level=logging.INFO,
    filename="controle_diario.log",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def get_chrome_path():
    with open("chrome-path.txt", mode="r") as f:
        return f.read()


# Set and open chrome
current_user = os.getlogin()
options = Options()
service = Service()
url = "http://trakcare.saude.df.gov.br/trakcare/"
options.binary_location = get_chrome_path()
service.path = "chromedriver.exe"

try:
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(url)
except Exception as e:
    logging.info(f"Could not load {driver}. Exception raised: {e}")
    print("Failed to login. Quitting the program...")
    driver.quit()
    sys.exit()


def find_fill_click(element1, element2=None, input_text=None):
    if input_text and element2 is not None:
        elem = driver.find_element(by=By.XPATH, value=element1)
        elem.send_keys(input_text)
        elem = driver.find_element(by=By.XPATH, value=element2)
        elem.click()
    elif element2 is None and input_text is not None:
        elem = driver.find_element(by=By.XPATH, value=element1)
        elem.send_keys(input_text)
        elem.click()
    else:
        elem = driver.find_element(by=By.XPATH, value=element1)
        elem.click()


def switch_window():
    for handle in driver.window_handles:
        driver.switch_to.window(handle)


def wait_element(element, delay):
    try:
        WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.XPATH, element))
        )
    except Exception as e:
        logging.info(
            f"Could not load {element} in {delay} seconds. Exception raised: {e}"
        )
        print("Loading took too much time!")
        print("Failed to login. Quitting the program...")
        driver.quit()
        sys.exit()


# Navigate until reach target page
try:
    find_fill_click(element1="""//*[@id="light"]/h22/i/h22/i/h22/i/h22/i/h22/h22/a""")
except Exception as e:
    logging.info(f"Could not load element cause no pop-up. Exception raised: {e}")
find_fill_click(element1="""//*[@id="ButtonLogin"]""")
switch_window()
wait_element("""//*[@id="USERNAME"]""", delay=10)
find_fill_click(element1="""//*[@id="USERNAME"]""", input_text="login")
find_fill_click(
    element1="""//*[@id="PASSWORD"]""",
    element2="""//*[@id="Logon"]""",
    input_text="password",
)
find_fill_click(element1="""//*[@id="LocListGroupz3"]""")
wait_element(element="""//*[@id="TRAK_main"]""", delay=10)

try:
    driver.execute_script(
        """document.querySelector('#TRAK_main').contentWindow.document.querySelector('area[href*="Leitos Contratados"]').click()"""
    )
except Exception as e:
    print("Failed to login. Quitting the program...")
    logging.info(f"Could not click in Leitos Contratos. Exception raised: {e}")
    driver.quit()
    sys.exit()

time.sleep(2)
driver.switch_to.frame("TRAK_main")
frame_html = driver.page_source
soup = BeautifulSoup(frame_html, "html.parser")
total_pages = int(len(soup.find_all(id=re.compile("F.Planz"))))
driver.switch_to.default_content()
logging.info(f"Number of pages to get data from: {total_pages}")
raw_contratadas = driver.execute_script(
    """return Array.from(document.querySelector("#TRAK_main").contentWindow.document.querySelectorAll("label[id^=Ward]")).map(node => node.textContent.trim())"""
)
contratadas = []

for raw_contratada in raw_contratadas:
    contratadas.append(
        raw_contratada.split("-")[0]
        .strip()
        .replace("H.", "")
        .replace(" Lago Sul", "")
        .replace("Sao", "São")
    )

tipo_uti_lst = []

for raw_contratada in raw_contratadas:
    tipo_uti_lst.append(
        raw_contratada.split("-")[1]
        .strip()
        .replace("UTI ", "")
        .replace("COVID", "Covid-19")
    )


# Get data from each page
for current_page in range(1, total_pages + 1):
    driver.execute_script(
        f"""document.querySelector('#TRAK_main').contentWindow.document.querySelector('a[id*="F.Planz{current_page}"]').click()"""
    )
    time.sleep(1)
    driver.switch_to.frame("TRAK_main")
    frame_html = driver.page_source
    soup = BeautifulSoup(frame_html, "html.parser")

    # Get contratada
    contratada = contratadas[current_page - 1]

    # Get date
    today_date = datetime.today().strftime("%d/%m/%Y")

    # Get tipo uti
    tipo_uti = tipo_uti_lst[current_page - 1]

    # Get leitos contratados
    leitos_contratados = len(soup.find_all("div", class_="Bed"))

    # Get leitos ocupados
    # pac_internados = len(
    # soup.find_all(string=re.compile("Pac Internado$"))  # Added "$" to avoid counting "Pac Internado COVID"
    pac_internados = len(soup.find_all(string=re.compile("Pac Internado")))
    leitos_ocupados = len(soup.find_all("div", class_="BedBody")) + pac_internados

    # Get leitos bloqueados
    direcionados = len(soup.find_all(string=re.compile("Direcionado")))
    bloqueados = len(soup.find_all("div", class_="BedBodyClosed"))
    leitos_bloqueados = bloqueados - direcionados - pac_internados

    # Get leitos vagos
    leitos_vagos = leitos_contratados - leitos_ocupados - leitos_bloqueados

    # Get alta medica
    bed_body = str(soup.find_all(class_="BedBody"))
    alta_medica = bed_body.count("Paciente com Alta Médica")

    # Get observacoes
    bloqueados = soup.find_all(class_="BedBodyClosed")
    obs_list = []
    to_remove_obs = [
        "Direcionado",
        "Direcionado COVID",
        "Pac Internado",
        "Pac internado COVID",  # Re-added "Pac Internado COVID"
    ]

    for not_formated_bloqueado in bloqueados:
        formated_bloqueado = not_formated_bloqueado.text.strip().split("\n")[0]
        if formated_bloqueado not in to_remove_obs:
            obs_list.append(formated_bloqueado)

    obs_counter = {i: obs_list.count(i) for i in obs_list}
    obs_counted = []

    for key in obs_counter:
        obs_counted.append(f"{obs_counter[key]} {key}")

    observacoes = " / ".join(obs_counted)

    # Result
    final_result = [
        today_date,
        contratada,
        tipo_uti,
        leitos_contratados,
        leitos_ocupados,
        leitos_vagos,
        leitos_bloqueados,
        alta_medica,
        observacoes,
    ]
    logging.info(final_result)
    print(final_result)

    # Export data do gsheets
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("UTI's 2023 (Contratos Assistenciais)").worksheet(
        "Controle Diário"
    )
    last_row = len(sheet.get_all_records())
    sheet.insert_row(final_result, last_row + 2, value_input_option="user_entered")

    # Next page
    driver.switch_to.default_content()
    driver.execute_script(
        """document.querySelector('#eprmenu').contentWindow.document.querySelector('a[id*="MainMenuItemAnchor51470"]').click()"""
    )
    time.sleep(1)


driver.quit()
print("The End")
end_time = time.perf_counter()
logging.info(f"Runtime: ~{int(end_time - start_time)} seconds")
print(f"Runtime: ~{int(end_time - start_time)} seconds")
