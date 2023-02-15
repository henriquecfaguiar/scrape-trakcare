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

import pandas as pd
import logging
import time
import sys
import re
import os

start_time = time.perf_counter()

logging.basicConfig(
    level=logging.INFO,
    filename="ocorrencias.log",
    filemode="w",
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def get_chrome_path():
    with open("chrome-path.txt", mode="r") as f:
        return f.read()


# Get data from google sheets
def get_gsheets_data(contratada=None, tipo_uti=None):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open("google sheets file name").worksheet("sheet name")
    data = sheet.get_all_records(numericise_ignore=["all"])
    gsheets_df = pd.DataFrame(data)

    # Filter data accordingly to contratada + tipo_uti
    mask1 = gsheets_df["Contratada"] == f"{contratada}"
    mask2 = gsheets_df["Tipo UTI"] == f"{tipo_uti}"
    mask3 = gsheets_df["Tipo de alta"] == ""
    gsheets_df = gsheets_df[mask1 & mask2 & mask3]
    return gsheets_df[["Nome", "Nº SES"]]


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
        .replace("H.", "")
        .replace("Lago Sul", "")
        .replace("Sao", "São")
        .strip()
    )

tipo_uti_lst = []

for raw_contratada in raw_contratadas:
    tipo_uti_lst.append(
        raw_contratada.split("-")[1]
        .replace("UTI ", "")
        .replace("COVID", "Covid-19")
        .strip()
    )

writer = pd.ExcelWriter("output/ocorrencias.xlsx", engine="xlsxwriter")
empty_table_counter = 0

# Get data from each page
for current_page in range(1, total_pages + 1):
    driver.execute_script(
        f"""document.querySelector('#TRAK_main').contentWindow.document.querySelector('a[id*="P.Listz{current_page}"]').click()"""
    )
    time.sleep(1)
    driver.switch_to.frame("TRAK_main")

    # Look for second page
    try:
        number_of_pages = driver.execute_script(
            """return document.querySelector(".tlbListFooter").querySelector("small").textContent.trim()"""
        )
        pacient_names1 = driver.execute_script(
            """return Array.from(document.querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        pacient_names_waiting_list1 = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        ses_number1 = driver.execute_script(
            """return Array.from(document.querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )
        ses_number_waiting_list1 = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )

        driver.execute_script(
            """document.querySelector("#NextPageImage_PACWard_ListPatientsInWard").click()"""
        )
        time.sleep(1)

        pacient_names2 = driver.execute_script(
            """return Array.from(document.querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        pacient_names_waiting_list2 = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        ses_number2 = driver.execute_script(
            """return Array.from(document.querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )
        ses_number_waiting_list2 = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )

        pacient_names = pacient_names1 + pacient_names2
        ses_number = ses_number1 + ses_number2
        pacient_names_waiting_list = (
            pacient_names_waiting_list1 + pacient_names_waiting_list2
        )
        ses_number_waiting_list = ses_number_waiting_list1 + ses_number_waiting_list2

    except Exception as e:
        pacient_names = driver.execute_script(
            """return Array.from(document.querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        ses_number = driver.execute_script(
            """return Array.from(document.querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )
        pacient_names_waiting_list = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("label[id^=Surname]")).map(node => node.textContent.trim())"""
        )
        ses_number_waiting_list = driver.execute_script(
            """return Array.from(document.querySelector("#tPACWardRoom_ListPatients2").querySelectorAll("a[id^=URz]")).map(node => node.textContent.trim())"""
        )

    # Data cleansing and compare tables
    if pacient_names:

        for idx, pacient_name in enumerate(pacient_names):
            if not pacient_name.startswith("Filho"):
                pacient_names[idx] = pacient_name.split("(")[0]

    if pacient_names_waiting_list:

        for idx, pacient_name_waiting_list in enumerate(pacient_names_waiting_list):
            if not pacient_name_waiting_list.startswith("Filho"):
                pacient_names_waiting_list[idx] = pacient_name_waiting_list.split("(")[
                    0
                ]

    trakcare_df = pd.DataFrame(data={"Nome": pacient_names, "Nº SES": ses_number})
    waiting_list = pd.DataFrame(
        data={"Nome": pacient_names_waiting_list, "Nº SES": ses_number_waiting_list}
    )

    gsheets_df = get_gsheets_data(
        contratada=contratadas[current_page - 1],
        tipo_uti=tipo_uti_lst[current_page - 1],
    )
    gsheets_df.reset_index(inplace=True)
    gsheets_df.drop(labels="index", axis="columns", inplace=True)

    to_remove = gsheets_df[~gsheets_df["Nº SES"].isin(trakcare_df["Nº SES"])]
    to_add = trakcare_df[~trakcare_df["Nº SES"].isin(gsheets_df["Nº SES"])]
    to_add = to_add[~to_add["Nº SES"].isin(waiting_list["Nº SES"])]
    to_remove.insert(loc=2, column="Status", value="Remover")
    to_add.insert(loc=2, column="Status", value="Adicionar")

    result_df = pd.concat([to_remove, to_add])

    if not result_df.empty:
        result_df.to_excel(
            writer,
            sheet_name=f"{contratadas[current_page - 1]}_{tipo_uti_lst[current_page - 1]}",
            index=False,
        )
        for column in result_df:
            column_width = max(
                result_df[column].astype("str").map(len).max(), len(column)
            )
            col_idx = result_df.columns.get_loc(column)
            writer.sheets[
                f"{contratadas[current_page - 1]}_{tipo_uti_lst[current_page - 1]}"
            ].set_column(col_idx, col_idx, column_width)
    else:
        empty_table_counter += 1

    # Next page
    driver.switch_to.default_content()
    driver.execute_script(
        """document.querySelector('#eprmenu').contentWindow.document.querySelector('a[id*="MainMenuItemAnchor51470"]').click()"""
    )
    time.sleep(1)

writer.close()

driver.quit()
print(f"Empty tables: {empty_table_counter}/{total_pages}")
print("The End")
end_time = time.perf_counter()
logging.info(f"Runtime: ~{int(end_time - start_time)} seconds")
print(f"Runtime: ~{int(end_time - start_time)} seconds")
