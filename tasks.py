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

    # dupa modelul generat automat la crearea proiectului
    rows = excel.read_worksheet_as_table("List1", header=True)

    clean_value = pret_produs.replace(' RON', '').split()[0]
    print(clean_value)
    clean_value_float=float(clean_value.replace(',', '.'))

    nr=rows.size
    record=True
    
    for i in range(1,nr,3):
        old_price=rows.get_row(i)
        value = list(old_price.values())[0]

        # Îndepărtează caracterul non-breaking space și orice alt caracter nedorit
        clean_value = value.replace('\xa0RON', '').split()[0]
        print(i,clean_value)
        clean_old_value_float=float(clean_value.replace(',', '.'))
    
        if clean_value_float>=clean_old_value_float:
            record=False

    if record:
        print("the price is the lowest recorded")

    excel.append_rows_to_worksheet([data_curenta, titlu_produs, pret_produs])
    excel.save_workbook(nume_fisier)


@task
def my_task_3b():

    # Navigate to Site
    page=browser.goto("https://www.sinsay.com/ro/ro/")
    
    # page.wait_for_timeout(1000000)  # Wait for 1000 seconds before closing the browser
    page.click('button:has-text("OK")')

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

    # Căutăm denumirea produsului   
    page.wait_for_selector('div[data-testid="products-results"]')

    # Selectăm toate imaginile din div-ul cu rezultatele produselor
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
        salveaza_in_excel(title,price)
    else:
        print("product not found")



@task
def solve_challenge():
    """
    Main task which solves the RPA challenge!

    Downloads the source data Excel file and uses Playwright to fill the entries inside
    rpachallenge.com.
    """
    browser.configure(
        browser_engine="chromium", 
        screenshot="only-on-failure", 
        headless=True 
    )
    try:
        # Reads a table from an Excel file hosted online.
        excel_file = download_file(
            EXCEL_URL, target_dir=OUTPUT_DIR, target_filename=FILE_NAME
        )
        excel = Excel()
        excel.open_workbook(excel_file)
        rows = excel.read_worksheet_as_table("Sheet1", header=True)

        # Surf the automation challenge website and fill in information from the table
        #  extracted above.
        page = browser.goto("https://rpachallenge.com/")
        page.click("button:text('Start')")
        for row in rows:
            fill_and_submit_form(row, page=page)
        element = page.locator("css=div.congratulations")
        browser.screenshot(element)
    finally:
        # A place for teardown and cleanups. (Playwright handles browser closing)
        print("Automation finished!")


def download_file(url: str, *, target_dir: Path, target_filename: str) -> Path:
    """
    Downloads a file from the given URL into a custom folder & name.

    Args:
        url: The target URL from which we'll download the file.
        target_dir: The destination directory in which we'll place the file.
        target_filename: The local file name inside which the content gets saved.

    Returns:
        Path: A Path object pointing to the downloaded file.
    """
    # Obtain the content of the file hosted online.
    response = requests.get(url)
    response.raise_for_status()  # this will raise an exception if the request fails
    # Write the content of the request response to the target file.
    target_dir.mkdir(exist_ok=True)
    local_file = target_dir / target_filename
    local_file.write_bytes(response.content)
    return local_file


def fill_and_submit_form(row: dict, *, page: browser.Page):
    """
    Fills a single form with the information of a single row from the table.

    Args:
        row: One row from the generated table out of the input Excel file.
        page: The page object over which the browser interactions are done.
    """
    field_data_map = {
        "labelFirstName": "First Name",
        "labelLastName": "Last Name",
        "labelCompanyName": "Company Name",
        "labelRole": "Role in Company",
        "labelAddress": "Address",
        "labelEmail": "Email",
        "labelPhone": "Phone Number",
    }
    for field, key in field_data_map.items():
        page.fill(f"//input[@ng-reflect-name='{field}']", str(row[key]))
    page.click("input:text('Submit')")
