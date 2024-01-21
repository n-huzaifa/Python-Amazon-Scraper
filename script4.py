import smtplib
import os
import logging
import json
import datetime
import time
import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
from script1 import main as script1_main
from script2 import main as script2_main
from script3 import main as script3_main
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from amazoncaptcha import AmazonCaptcha

with open("data.json", 'r') as file:
    data = json.load(file)

recipient_email = data["recipient_email"]
sender_email = data["sender_email"]
sender_password = data["sender_password"]
all_urls = data["all_urls"]

def configure_driver():
    chrome_options = Options()
    chrome_options.add_extension('./amazoncrxextension.crx')
    chrome_options.add_argument(r'--user-data-dir=F:\OuoC')
    driver = webdriver.Chrome(options=chrome_options)
    driver.maximize_window()
    return driver

def handle_cookies(driver):
    try:
        # Use WebDriverWait to wait for the captcha image to be present
        image_element = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//div[@class='a-row a-text-center']//img")))
        link = image_element.get_attribute('src')

        # Solve the captcha
        captcha = AmazonCaptcha.fromlink(link)
        captcha_value = AmazonCaptcha.solve(captcha)

        # Find the input field and enter the captcha value
        input_field = driver.find_element(By.ID, "captchacharacters")
        input_field.send_keys(captcha_value)

        # Find the submit button and click it
        button = driver.find_element(By.CLASS_NAME, "a-button-text")
        button.click()
    except TimeoutException:
        pass
    try:
        reject_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, 'sp-cc-rejectall-link')))
        reject_button.click()
    except TimeoutException:
        pass

def get_category_folder(category):
    category_mapping = {
        'Electrical Goods': "Category 1",
        'Fashion & Accessories': "Category 2",
        'Home & Garden': "Category 3",
        'Office & Business Equipment': "Category 4",
        'DIY': "Category 5",
        'Drugstore & Cosmetics': "Category 6",
        'Baby & Child': "Category 7",
        'Sport & Leisure': "Category 8",
        'Pet Supplies': "Category 9",
        'Car & Motorbike': "Category 10",
        'Books, Media & Entertainment': "Category 11",
        'Food & Beverages': "Category 12",
        'Other': "Category 13"
    }
    return category_mapping.get(category, None)

def get_excel_files(category, script):

    excel_files = {}
    
    main_folder_name = get_category_folder(category)

    def is_skript_directory(path):
        if script == 0:
            return "Skript 1" in path or "Skript 2" in path
        elif script == 1:
            return "Skript 3" in path
        else:
            return False

    if os.path.exists(main_folder_name) and os.path.isdir(main_folder_name):
        # Iterate through all subdirectories and files in the given directory
        for root, dirs, files in os.walk(main_folder_name):
            # Check if the current directory contains the specified text
            if is_skript_directory(root):
                for file in files:
                    # Check if the file has a .xls or .xlsx extension
                    if file.lower().endswith(('.xls', '.xlsx')):
                        # Create the full path to the Excel file
                        file_path = os.path.join(root, file)
                        
                        # Update the dictionary with the latest file for each folder
                        folder_key = os.path.dirname(file_path)
                        if folder_key not in excel_files or os.path.getmtime(file_path) > os.path.getmtime(excel_files[folder_key]):
                            excel_files[folder_key] = file_path

    # Return a list of the latest Excel files from each folder
    return list(excel_files.values())

def create_email(sender_email, recipient, subject, message, excel_files):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    for file in excel_files:
        with open(file, 'rb') as file_content:
            excel_attachment = MIMEApplication(file_content.read(), _subtype="xlsx")
        excel_attachment.add_header('content-disposition', 'attachment', filename=file)
        msg.attach(excel_attachment)
        logging.info(f"file attached: {file}")

    return msg

def send_mail(sender_email, sender_password, recipient, subject, message, excel_files):
    smtp_server = smtplib.SMTP('smtp.gmail.com', 587)
    smtp_server.starttls()

    try:
        smtp_server.login(sender_email, sender_password)
        email_msg = create_email(sender_email, recipient, subject, message, excel_files)
        smtp_server.sendmail(sender_email, recipient, email_msg.as_string())
        logging.info("Email sent successfully.")
    except Exception as e:
        logging.error(f"Failed to send email. Error: {e}")
    finally:
        smtp_server.quit()

def main_mail(category, script):
    subject = 'Excel File'
    message = 'Please find the attached Excel file.'

    excel_files = get_excel_files(category, script)
    send_mail(sender_email, sender_password, recipient_email, subject, message, excel_files)

def is_end_of_week():
    today = datetime.datetime.now().weekday()
    if today == 6:
        global is_script1_done
        is_script1_done = False
    # Assuming Friday is the end of the week (Monday = 0, Sunday = 6)
    return today == 6 

def main1(driver):

    for category, url_data in all_urls.items():

        script1_main(driver, url_data)
        script2_main(driver, category)
        # main_mail(category, 0)
        
    global is_script1_done
    is_script1_done = True

def amalyse_login(driver):
    driver.get(f"https://shield.amalyze.com/auth/login?redirect=%2Fdashboard")
    # Wait for the space key press
    while True:
        if keyboard.is_pressed(' '):
            break
        time.sleep(1)

def main():

    driver = configure_driver()
    handle_cookies(driver)
    amalyse_login(driver)

    main1(driver)
    
    for _ in all_urls:
        if is_end_of_week(is_script1_done):
            if not is_script1_done:
                main1(driver)
        
        script3_main(driver)

if __name__ == "__main__":
    logging.basicConfig(filename=f'Log Files\\logfile_{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}.txt ', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    while True:
        main()
