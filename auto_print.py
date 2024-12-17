from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
import os
import base64
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Setup Selenium WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run headless Chrome
chrome_options.add_argument("--disable-gpu")  # Disable GPU usage for headless mode
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--ignore-ssl-errors")
chrome_options.add_argument("--disable-usb")

def save_as_pdf(driver, file_path):
    # Print current page as PDF using CDP
    response = driver.execute_cdp_cmd("Page.printToPDF", {
        'paperWidth': 8.27,   # A4 size in inches
        'paperHeight': 11.69, # A4 size in inches
        'marginTop': 0.4,     # top margin in inches
        'marginBottom': 0.4,  # bottom margin in inches
        'marginLeft': 0.4,    # left margin in inches
        'marginRight': 0.4,   # right margin in inches
        'printBackground': True
    })
    pdf_data = base64.b64decode(response['data'])
    with open(file_path, 'wb') as file:
        file.write(pdf_data)

class AutoPrint:
    def __init__(self, save_folder):
        self.save_folder = save_folder
        if not os.path.exists(save_folder):
            os.makedirs(save_folder)

    def auto_print(self, url, doc_number):
        try:
            service = ChromeService(executable_path=ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=chrome_options)

            # Open a webpage
            driver.get(url)
            time.sleep(5)
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(5)
            if doc_number < 10:
                file_path = os.path.join(self.save_folder, f'0{doc_number}.pdf')
            else:
                file_path = os.path.join(self.save_folder, f'{doc_number}.pdf')
            # Save the current page as a PDF
            save_as_pdf(driver, file_path)
            print(f"Document {doc_number} saved successfully to {file_path}")

        except FileNotFoundError as fnf_error:
            print(fnf_error)
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            # Close the Selenium WebDriver
            if 'driver' in locals():
                driver.quit()

