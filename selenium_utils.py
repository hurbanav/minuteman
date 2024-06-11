from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
import logging
import time


def _init_logger():
    logger = logging.getLogger('SeleniumUtils')
    logger.setLevel(logging.DEBUG)
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    ch.setFormatter(formatter)
    logger.addHandler(ch)
    return logger


class SeleniumUtils:
    def __init__(self, driver_path, browser='chrome', headless=False):
        self.logger = _init_logger()
        self.driver = self._init_driver(driver_path, browser, headless)

    def _init_driver(self, driver_path, browser, headless):

        driver = None
        if browser == 'chrome':
            from selenium.webdriver.chrome.service import Service as ChromeService
            from selenium.webdriver.chrome.options import Options as ChromeOptions
            options = ChromeOptions()
            service = ChromeService(driver_path)
        elif browser == 'firefox':
            from selenium.webdriver.firefox.service import Service as FirefoxService
            from selenium.webdriver.firefox.options import Options as FirefoxOptions
            options = FirefoxOptions()
            service = FirefoxService(driver_path)
        elif browser == 'edge':
            from selenium.webdriver.edge.service import Service as EdgeService
            from selenium.webdriver.edge.options import Options as EdgeOptions
            options = EdgeOptions()
            service = EdgeService(driver_path)
        else:
            raise ValueError("Unsupported browser! Use 'chrome', 'firefox', or 'edge'.")

        if headless:
            options.add_argument("--headless")
        options.add_argument("--disable-extensions")
        options.add_experimental_option("detach", True)
        
        try:
            if browser == 'chrome':
                driver = webdriver.Chrome(service=service, options=options)
            elif browser == 'firefox':
                driver = webdriver.Firefox(service=service, options=options)
            elif browser == 'edge':
                driver = webdriver.Edge(service=service, options=options)
            self.logger.info("WebDriver initialized successfully.")
            driver.maximize_window()
            return driver
        except WebDriverException as e:
            self.logger.error(f"Error initializing WebDriver: {e}")
            raise

    def open_url(self, url):
        try:
            self.driver.get(url)
            self.logger.info(f"Opened URL: {url}")
        except WebDriverException as e:
            self.logger.error(f"Error opening URL {url}: {e}")
            raise

    def click_element(self, locator, by='xpath', timeout=20, simple=False):
        time.sleep(0.5)
        if simple:
            self.wait_for_clickable_element('//th/input').click()
            self.logger.info(f"Clicked element by {by} with locator {locator}")
            return

        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = WebDriverWait(self.driver, timeout).until(
                ec.element_to_be_clickable((by, locator))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
            actions = ActionChains(self.driver)
            actions.move_to_element(element).click().perform()
            self.logger.info(f"Clicked element by {by} with locator {locator}")
            time.sleep(0.5)
        except (TimeoutException, NoSuchElementException) as e:
            self.logger.error(f"Error clicking element by {by} with locator {locator}: {e}")
            raise

    def find_element(self, locator, by='xpath', timeout=10):
        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = WebDriverWait(self.driver, timeout).until(
                ec.presence_of_element_located((by, locator))
            )
            self.logger.info(f"Found element by {by} with locator {locator}")
            return element
        except (TimeoutException, NoSuchElementException) as e:
            self.logger.error(f"Error finding element by {by} with locator {locator}: {e}")

    def send_keys(self, locator, keys, by='xpath', timeout=10):
        try:
            element = self.find_element(locator, by, timeout)
            element.clear()
            element.send_keys(keys)
            self.logger.info(f"Sent keys to element by {by} with locator {locator}")
        except (TimeoutException, NoSuchElementException) as e:
            self.logger.error(f"Error sending keys to element by {by} with locator {locator}: {e}")
            raise

    def quit(self):
        try:
            self.driver.quit()
            self.logger.info("WebDriver quit successfully.")
        except WebDriverException as e:
            self.logger.error(f"Error quitting WebDriver: {e}")
            raise

    def wait_for_element(self, locator, by='xpath', timeout=30):
        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = WebDriverWait(self.driver, timeout).until(
                ec.visibility_of_element_located((by, locator))
            )
            self.logger.info(f"Element by {by} with locator {locator} is visible.")
            return element
        except TimeoutException as e:
            self.logger.error(f"Timeout waiting for element by {by} with locator {locator}: {e}")
            raise

    def wait_for_clickable_element(self, locator, by='xpath', timeout=30):
        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = WebDriverWait(self.driver, timeout).until(
                ec.element_to_be_clickable((by, locator))
            )
            self.logger.info(f"Element by {by} with locator {locator} is visible.")
            return element
        except TimeoutException as e:
            self.logger.error(f"Timeout waiting for element by {by} with locator {locator}: {e}")
            raise

    def scroll_to_element(self, locator, by='xpath', timeout=10):
        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = self.find_element(locator, by, timeout)
            self.driver.execute_script("arguments[0].scrollIntoView();", element)
            self.logger.info(f"Scrolled to element by {by} with locator {locator}")
        except (TimeoutException, NoSuchElementException) as e:
            self.logger.error(f"Error scrolling to element by {by} with locator {locator}: {e}")
            raise

    def element_clickable(self, locator, by='xpath', timeout=10):
        find_by = {
            "id": By.ID,
            "xpath": By.XPATH,
            "link_text": By.LINK_TEXT,
            "partial_link_text": By.PARTIAL_LINK_TEXT,
            "name": By.NAME,
            "tag_name": By.TAG_NAME,
            "class_name": By.CLASS_NAME,
            "css": By.CSS_SELECTOR,
        }
        if by not in find_by:
            raise ValueError(f"Unsupported locator strategy: {by}")
        by = find_by[by]
        try:
            element = WebDriverWait(self.driver, timeout).until(
                ec.element_to_be_clickable((by, locator))
            )
            self.logger.info(f"Element by {by} with value {locator} is clickable.")
            return element
        except TimeoutException as e:
            self.logger.error(f"Timeout waiting for element by {by} with value {locator} to be clickable: {e}")
            raise
