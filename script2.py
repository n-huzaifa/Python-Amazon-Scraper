import os
import pyperclip
import openpyxl
import logging
from time import sleep
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

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

def get_excel_files(directory):
    excel_files = []
    
    if os.path.exists(directory) and os.path.isdir(directory):
        # Iterate through all subdirectories and files in the given directory
        for root, dirs, files in os.walk(directory):
            # Check if the current directory contains the specified text
            if "skript 1" in root.lower():
                for file in files:
                    # Check if the file has a .xls or .xlsx extension
                    if file.lower().endswith(('.xls', '.xlsx')):
                        # Create the full path to the Excel file
                        file_path = os.path.join(root, file)
                        excel_files.append(file_path)

    return excel_files

def load_amazon_urls(excel_file_path):
    try:
        wb = openpyxl.load_workbook(excel_file_path)
        sheet = wb.active
        amazon_urls = [tuple(sheet.cell(row=cell.row, column=col).value for col in range(1, sheet.max_column + 1)) for cell in sheet['A'] if cell.value and cell.row > 1]
        logging.info(f"Amazon URLs loaded from: {excel_file_path}")
        return amazon_urls
    
    except Exception as e:
        logging.error(f"Error loading Amazon URLs from {excel_file_path}: {e}")
        raise
    
def create_or_load_workbook(directory):
    
    try:
        folder_path = os.path.join(directory, f'6. Skript 2')
        os.makedirs(folder_path, exist_ok=True)

        excel_file_path = os.path.join(folder_path, f'./Script2_Germany_{directory}_{datetime.datetime.now().strftime("%Y%m%d")}.xlsx')

        if os.path.isfile(excel_file_path):
            workbook = openpyxl.load_workbook(excel_file_path)
            logging.info(f"Workbook loaded: {excel_file_path}")
            worksheet = workbook.active
            last_row = worksheet.max_row + 1
        else:
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.cell(1, 1).value = 'ASINs'
            worksheet.cell(1, 2).value = 'Category Name'
            worksheet.cell(1, 3).value = 'Category Level'
            worksheet.cell(1, 4).value = 'Date'
            workbook.save(excel_file_path)
            logging.info(f"New workbook created: {excel_file_path}")
            last_row = 2

        return workbook, excel_file_path, last_row

    except Exception as e:
        logging.error(f"Error creating or loading workbook: {e}")
        raise

def extract_asins_and_category(driver, url, worksheet, last_row, Category_Level):
    driver.get(url)
    sleep(3)
    handle_cookies(driver)

    max_attempts = 3  # You can adjust the maximum number of retry attempts

    for attempt in range(max_attempts):
        try:
            asin_extractor = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, 'azASINExtractorDropDown')))
            asin_extractor.click()

            all_asins_option = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'header') and contains(text(), 'All ASINs')]")))
            all_asins_option.click()

            sleep(3)
            captured_asins = pyperclip.paste()
            asins_list = captured_asins.split("\n")

            if not asins_list:
                raise ValueError("ASINs list is empty. Retrying...")

            category_name = driver.find_element(By.XPATH, "//h1[contains(@class, 'a-size-large')]").text
            category_name = category_name.replace("Best Sellers in ", "")

            for value in asins_list:
                worksheet.cell(row=last_row, column=1).value = value
                worksheet.cell(row=last_row, column=2).value = category_name
                worksheet.cell(row=last_row, column=3).value = Category_Level
                worksheet.cell(row=last_row, column=4).value = datetime.datetime.now().strftime("%Y-%m-%d")
                last_row += 1
            logging.info(f"ASINs extracted from: {category_name}")
            return last_row
            # Break out of the loop if extraction is successful

        except Exception as e:
            logging.error("Error:", e)
            if attempt < max_attempts - 1:
                logging.info(f"Retry attempt {attempt + 1}/{max_attempts}")
            else:
                logging.info("Max retry attempts reached. Exiting.")
                return last_row

def remove_duplicate_asins(worksheet):
    unique_asins = set()
    rows_to_delete = []

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, max_col=1):
        for cell in row:
            if cell.value in unique_asins or cell.value == None:
                rows_to_delete.append(cell.row)
                logging.warning(f"Duplicate ASIN found in row {cell.row}. Removing.")
            else:
                unique_asins.add(cell.value)

    for row_number in reversed(rows_to_delete):
        worksheet.delete_rows(row_number)

def save_workbook(workbook, excel_file_path):
    try:
        workbook.save(excel_file_path)
        logging.info(f"Workbook saved: {excel_file_path}")
    except Exception as e:
        logging.error(f"Error saving workbook: {e}")
        raise

def main(driver, directory = None):

    try:
        excel_files = get_excel_files(directory)
        workbook, asin_excel_file_path, last_row = create_or_load_workbook(directory)
                        
        for excel_file_path in excel_files:
            try:
                amazon_urls = load_amazon_urls(excel_file_path)

                for url in amazon_urls:
                    try:
                        Category_URL, Category_Level = url[1], url[2]

                        worksheet = workbook.active

                        last_row = extract_asins_and_category(driver, Category_URL, worksheet, last_row, Category_Level)

                        save_workbook(workbook, asin_excel_file_path)
                    
                    except Exception as e:
                        logging.error(f"Error processing URL {url}: {e}")

            except Exception as e:
                logging.error(f"Error processing Excel file {excel_file_path}: {e}")
        
        remove_duplicate_asins(worksheet)
        save_workbook(workbook, asin_excel_file_path)
        workbook.close()
        logging.info(f"Workbook closed: {asin_excel_file_path}")
        
        logging.info("Script execution completed successfully.")
    
    except Exception as e:
        logging.error(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    extension_path = './amazoncrxextension.crx'
    driver = setup_chrome_driver(extension_path)
    handle_cookies(driver)
    main(driver, "Adapters")