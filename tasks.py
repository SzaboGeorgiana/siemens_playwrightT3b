from datetime import datetime
import os
from pathlib import Path

from openpyxl import Workbook, load_workbook
import requests
from robocorp import browser
from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel

FILE_NAME = "challenge.xlsx"
EXCEL_URL = f"https://rpachallenge.com/assets/downloadFiles/{FILE_NAME}"
OUTPUT_DIR = Path(os.getenv("ROBOT_ARTIFACTS", "output"))

import re
from playwright.sync_api import Page, expect

# Funcție pentru salvarea prețului și datei curente într-un fișier Excel
def salveaza_in_excel(titlu_produs, pret_produs):
    # Data curentă
    data_curenta = datetime.now().strftime("%Y-%m-%d")
    
    # Numele fisierului Excel
    nume_fisier = "preturi.xlsx"
    excel = Excel()
    excel.open_workbook(nume_fisier)

    # salvam pretul curent in Excel
    excel.append_rows_to_worksheet([data_curenta, titlu_produs, pret_produs])
    excel.save_workbook(nume_fisier)


def verifica_record(pret_produs):
    # Numele fisierului Excel
    nume_fisier = "preturi.xlsx"
    excel = Excel()
    excel.open_workbook(nume_fisier)

    # Luam datele dupa modelul generat automat la crearea proiectului
    rows = excel.read_worksheet_as_table("List1", header=True)

    # transformam pretul curent in float
    clean_value = pret_produs.replace(' RON', '').split()[0]
    print(clean_value)
    clean_value_float=float(clean_value.replace(',', '.'))

    nr=rows.size
    record=True
    
    # parcurgem datele ce le avem deja in Excel
    for i in range(1,nr,3):
        old_price=rows.get_row(i)
        value = list(old_price.values())[0]

        # transformam pe rand fiecare pret din Excel in float
        clean_value = value.replace('\xa0RON', '').split()[0]
        print(i,clean_value)
        clean_old_value_float=float(clean_value.replace(',', '.'))
    
        # comparam preturile anterioare din Excel cu pretul curent
        if clean_value_float>=clean_old_value_float:
            record=False

    if record:
        print("the price is the lowest recorded")


def cauta_produs(page):
    # Așteaptă să fie disponibil câmpul de căutare și butonul de căutare
    search_button_selector='button:has-text("Căutare")'
    page.wait_for_selector(search_button_selector)
    page.click(search_button_selector)

    search_input_selector = 'input[name="query"]'  
    page.wait_for_selector(search_input_selector)

    product_name= "Perie de păr Stitch"
    # Completează câmpul de căutare
    page.fill(search_input_selector, product_name)
    page.press(search_input_selector,"Enter")

    return product_name


@task
def my_task_3b():

    # Navigate to Site
    page=browser.goto("https://www.sinsay.com/ro/ro/")
    
    # page.wait_for_timeout(1000000)  # Wait for 1000 seconds before closing the browser
    page.click('button:has-text("OK")')

    # Căutăm denumirea produsului   
    product_name=cauta_produs(page)

    # Lista de rezultate a produselor
    page.wait_for_selector('div[data-testid="products-results"]')
    list_of_results = page.locator('div[data-testid="products-results"]')

    # Extragerea titlului produsului
    title_locator = list_of_results.locator('.ds-product-tile-name h2')
    title = title_locator.inner_text()

    # Extragerea prețului produsului
    price_locator = list_of_results.locator('.final-price')
    price = price_locator.inner_text()

    # Afișarea detaliilor produsului
    print(f"Titlu: {title}")
    print(f"Preț: {price}")
    if title==product_name:
        verifica_record(price)
        salveaza_in_excel(title,price)
    else:
        print("product not found")

