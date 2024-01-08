import smtplib
import os
import logging
import json
import datetime
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

with open("data.json", 'r') as file:
    data = json.load(file)

recipient_email = data["recipient_email"]
sender_email = data["sender_email"]
sender_password = data["sender_password"]
all_urls_unsorted = data["all_urls"]
all_urls = sorted(all_urls_unsorted, key=lambda x: x[2])

extension_path = './amazoncrxextension.crx'
base_url = 'https://www.amazon.de/'

def configure_driver():
    chrome_options = Options()
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

def get_excel_files(directory, script):
    excel_files = []
    
    def is_skript_directory(path):
        if script == 0:
            return "skript 1" in path.lower() or "skript 2" in path.lower()
        elif script == 1:
            return "skript 3" in path.lower()
        else:
            return False

    if os.path.exists(directory) and os.path.isdir(directory):
        for root, dirs, files in os.walk(directory):
            if is_skript_directory(root):
                for file in files:
                    if file.lower().endswith(('.xls', '.xlsx')):
                        file_path = os.path.join(root, file)
                        excel_files.append(file_path)

    return excel_files

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

def main_mail(directory, script):
    subject = 'Excel File'
    message = 'Please find the attached Excel file.'

    excel_files = get_excel_files(directory, script)
    send_mail(sender_email, sender_password, recipient_email, subject, message, excel_files)

def main(all_urls):

    logging.basicConfig(filename=f'logfile_{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}.txt ', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    driver = configure_driver()
    handle_cookies(driver)

    previous_category = all_urls[0][2]
    for url_data in all_urls:
        category = url_data[2]
        
        if category != previous_category:
            script2_main(driver, previous_category)
            main_mail(previous_category, 0)
            script3_main(driver, previous_category)
            main_mail(previous_category, 1)
            previous_category = category
        
        script1_main(driver, url_data)

    script2_main(driver, category)
    main_mail(category, 0)
    script3_main(driver, category)
    main_mail(category, 1)

    # Close the WebDriver instance after all scripts have run
    driver.quit()

if __name__ == "__main__":
    main(all_urls)
