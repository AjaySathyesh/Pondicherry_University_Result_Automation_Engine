from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import time

class WordChecker:
    def __init__(self, browser='chrome'):
        self.driver = self._get_driver(browser)

    def _get_driver(self, browser):
        if browser.lower() == 'chrome':
            service = ChromeService(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service)
        elif browser.lower() == 'firefox':
            service = FirefoxService(GeckoDriverManager().install())
            driver = webdriver.Firefox(service=service)
        else:
            raise ValueError("Unsupported browser. Please use 'chrome' or 'firefox'.")
        return driver

    def check_word_in_webpage(self, url, word, retries=3, delay=5):
        attempt = 0
        while attempt < retries:
            try:
                self.driver.get(url)
                page_source = self.driver.page_source
                return word in page_source
            except Exception as e:
                print(f"Attempt {attempt + 1} failed with error: {e}")
                attempt += 1
                time.sleep(delay)
        
        return False
    
    def close_driver(self):
        self.driver.quit()
