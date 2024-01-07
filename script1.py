import os
import openpyxl
import logging
from time import sleep
from datetime import datetime
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

all_urls = [[ "https://www.amazon.de/-/en/gp/bestsellers/computers/430049031/ref=zg_bs_nav_computers_2_427958031" , 2, "Adapters"],
             ["https://www.amazon.de/-/en/gp/new-releases/computers/430049031/ref=zg_bsnr_nav_computers_2_427958031" , 2, "Adapters"],
             ["https://www.amazon.de/-/en/gp/most-wished-for/computers/430049031/ref=zg_mw_nav_computers_2_427958031" , 2, "Adapters"],
             ["https://www.amazon.de/-/en/gp/most-gifted/computers/430049031/ref=zg_mg_nav_computers_2_427958031" , 2, "Adapters"]
             ]

def setup_chrome_driver(extension_path):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_extension(extension_path)
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver

def handle_cookies(driver):
    try:
        reject_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, 'sp-cc-rejectall-link')))
        reject_button.click()
    except TimeoutException:
        pass

def load_or_create_workbook(Category_Name, Category_Type, excel_file_name):
    
    category_folder_name = Category_Name
    os.makedirs(category_folder_name, exist_ok=True)

    if Category_Type == 'Hot New Releases':
        category_type_folder_path = os.path.join(Category_Name, f'1. Skript 1 {Category_Type}')
    elif Category_Type == 'Best Sellers':
        category_type_folder_path = os.path.join(Category_Name, f'2. Skript 1 {Category_Type}')
    elif Category_Type == 'Movers & Shakers':
        category_type_folder_path = os.path.join(Category_Name, f'3. Skript 1 {Category_Type}')
    elif Category_Type == 'Most Wished For':
        category_type_folder_path = os.path.join(Category_Name, f'4. Skript 1 {Category_Type}')
    elif Category_Type == 'Most Gifted':
        category_type_folder_path = os.path.join(Category_Name, f'5. Skript 1 {Category_Type}')
    
    os.makedirs(category_type_folder_path, exist_ok=True)

    excel_file_path = os.path.join(category_type_folder_path, excel_file_name)

    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        worksheet = workbook.active
        logging.info(f"Workbook loaded successfully: {excel_file_path}")
    except FileNotFoundError:
        logging.info(f"New Workbook created: {excel_file_path}")
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

    return excel_file_path, workbook, worksheet

def scrape_and_write(driver, worksheet, level, url, row, main_category_type):
    logging.info(f"Scraping and writing data for level {level} from URL: {url}")

    for current_level in range(level, 4):
        level_urls = scrape_categories(driver, url)
        for level_url in level_urls:
            if level_url['url']:
                worksheet.cell(row=row, column=1).value = level_url['name']
                worksheet.cell(row=row, column=2).value = level_url['url']
                if str(current_level) == "0":
                    worksheet.cell(row=row, column=3).value = "100"
                else:
                    worksheet.cell(row=row, column=3).value = str(current_level)
                worksheet.cell(row=row, column=4).value = main_category_type
                row += 1

                if current_level < 4:
                    row = scrape_and_write(driver, worksheet, current_level + 1, level_url['url'], row, main_category_type)

    logging.info(f"Scraping and writing data for level {level} completed.")
    return row

def scrape_categories(driver, url):
    base_url = 'https://www.amazon.de/'
    max_retries = 3
    array_to_append = []

    for _ in range(max_retries):
        try:
            
            driver.get(url)
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, '_p13n-zg-nav-tree-all_style_zg-browse-group__88fbz')))

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            group_divs = soup.find_all('div', {'role': 'group', 'class': '_p13n-zg-nav-tree-all_style_zg-browse-group__88fbz'})

            for group_div in group_divs:
                categories = []
                item_divs = group_div.find_all('div', {'role': 'treeitem', 'class': '_p13n-zg-nav-tree-all_style_zg-browse-item__1rdKf'})

                for item_div in item_divs:
                    anchor_tag = item_div.find('a')
                    span_tag = item_div.find('span')

                    if span_tag:
                        categories.append({'name': span_tag.text.strip(), 'url': ''})

                    if anchor_tag:
                        category_name = anchor_tag.text.strip()
                        category_url = base_url + anchor_tag['href']
                        categories.append({'name': category_name, 'url': category_url})

                for category in categories:
                    array_to_append.append({'name': category['name'], 'url': category['url']})
            break

        except TimeoutException:
            logging.warning("Timed out waiting for page to load, retrying...")
    
    return array_to_append

def remove_duplicate_categories(worksheet):
    unique_categories = set()

    for row in range(worksheet.max_row, 1, -1):
        category_name = worksheet.cell(row=row, column=1).value
        if category_name in unique_categories or category_name is None:
            logging.info(f"Removing duplicate category: {category_name}")
            worksheet.delete_rows(row)
        else:
            unique_categories.add(category_name)

    logging.info("Duplicate categories removed successfully.")

def save_workbook(workbook, excel_file_path):
    try:
        workbook.save(excel_file_path)
        logging.info(f"Workbook saved successfully: {excel_file_path}")
    except Exception as e:
        logging.error(f"An error occurred while saving workbook {excel_file_path}: {str(e)}")

def main(driver, url_data):
    
    try:
        url = url_data[0]
        level = url_data[1]
        category_name = url_data[2]

        driver.get(url)
        sleep(5)
        handle_cookies(driver)

        soup = BeautifulSoup(driver.page_source, "lxml")
        page_title = soup.title.text.strip()
        main_category_name = category_name
        main_category_type = page_title[10:page_title.find(":")].strip()

        excel_file_name = f'./Script1_Germany_{main_category_name}_{datetime.now().strftime("%Y%m%d")}.xlsx'

        excel_file_path, workbook, worksheet = load_or_create_workbook(main_category_name, main_category_type, excel_file_name)

        worksheet.cell(row=1, column=1).value = "Category Name"
        worksheet.cell(row=1, column=2).value = "Category URL"
        worksheet.cell(row=1, column=3).value = "Category Level"
        worksheet.cell(row=1, column=4).value = "Category Type"

        row = 2
        row = scrape_and_write(driver, worksheet, level, url, row, main_category_type)

        remove_duplicate_categories(worksheet)
        save_workbook(workbook, excel_file_path)

    except Exception as e:
        logging.error(f"An error occurred while scraping URL {url}: {str(e)}")

    logging.info(f"All URLs for {main_category_type} processed successfully.")

if __name__ == "__main__":
    base_url = 'https://www.amazon.de/'
    extension_path = './amazoncrxextension.crx'
    driver = setup_chrome_driver(extension_path)
    handle_cookies(driver)
    for url_data in all_urls:
        main(driver, url_data)