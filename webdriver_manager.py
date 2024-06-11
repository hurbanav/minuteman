import os
import subprocess
import platform
import urllib.request
import zipfile
import logging
import requests
from packaging import version

class WebDriverManager:
    def __init__(self, driver_directory='drivers'):
        self.driver_directory = os.path.abspath(driver_directory)
        self.logger = self._init_logger()
        os.makedirs(self.driver_directory, exist_ok=True)
        self.versions_url = 'https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json'
        self.versions_dict = self.fetch_versions_dict(self.versions_url)
        self.driver_path = self.setup_driver_path()

    def _init_logger(self):
        logger = logging.getLogger('WebDriverManager')
        logger.setLevel(logging.DEBUG)
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        return logger

    def _get_driver_filename(self, driver_name):
        system = platform.system()
        if system == 'Windows':
            return f'{driver_name}.exe'
        elif system in ['Linux', 'Darwin']:
            return driver_name
        else:
            raise ValueError(f'Unsupported operating system: {system}')

    def _get_driver_filepath(self, driver_name):
        return os.path.join(self.driver_directory, driver_name, self._get_driver_filename('chromedriver'))

    def _download_driver(self, url, driver_name):
        self.logger.info(f"Downloading WebDriver from {url}")
        zip_path = os.path.join(self.driver_directory, f'{driver_name}.zip')
        urllib.request.urlretrieve(url, zip_path)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(self.driver_directory)
        os.remove(zip_path)
        self.logger.info(f"WebDriver downloaded and extracted to {os.path.join(self.driver_directory, driver_name)}")

    def get_chrome_version(self):
        """Gets the installed version of Chrome."""
        system = platform.system()
        if system == 'Windows':
            process = subprocess.run(['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version'], capture_output=True, text=True)
            output = process.stdout
            if 'version' in output:
                return output.split()[-1]
        elif system == 'Darwin':  # macOS
            process = subprocess.run(['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'], capture_output=True, text=True)
            output = process.stdout
            return output.split()[-1]
        elif system == 'Linux':
            process = subprocess.run(['google-chrome', '--version'], capture_output=True, text=True)
            output = process.stdout
            return output.split()[-1]
        return None

    def find_closest_webdriver(self, chrome_version, webdriver_versions):
        """Finds the closest WebDriver version for the given Chrome version."""
        chrome_version = version.parse(chrome_version)
        closest_version = None
        closest_diff = None
        for v in webdriver_versions:
            v_parsed = version.parse(v)
            # Calculate the difference between version components
            diff = sum(abs(a - b) for a, b in zip(chrome_version.release, v_parsed.release))
            if closest_diff is None or diff < closest_diff:
                closest_version = v
                closest_diff = diff
        self.logger.info(f'Closest webdriver version is {closest_version}')
        return closest_version

    def get_webdriver_download_url(self, platform, versions_dict):
        """Gets the WebDriver download URL for the given platform."""
        chrome_version = self.get_chrome_version()
        if not chrome_version:
            raise ValueError("Could not determine Chrome version.")

        builds = versions_dict["builds"]
        closest_version = self.find_closest_webdriver(chrome_version, builds.keys())

        for download in builds[closest_version]["downloads"]["chromedriver"]:
            if download["platform"] == platform:
                return download["url"]

        raise ValueError(f"No WebDriver found for platform: {platform}")

    def fetch_versions_dict(self, url):
        """Fetches the WebDriver versions dictionary from a URL."""
        try:
            response = requests.get(url)
            response.raise_for_status()
            versions_dict = response.json()
            self.logger.info("Fetched versions dictionary successfully.")
            return versions_dict
        except requests.RequestException as e:
            self.logger.error(f"Error fetching versions dictionary: {e}")
            raise

    def get_driver_filepath(self, driver_name, download_url=None):
        """Gets the file path of the WebDriver, downloading it if not present."""
        driver_filepath = self._get_driver_filepath(driver_name)
        if not os.path.exists(driver_filepath):
            if not download_url:
                raise ValueError("Download URL must be provided if WebDriver is not present.")
            self.logger.info(f"WebDriver not found at {driver_filepath}. Downloading...")
            self._download_driver(download_url, driver_name)
        else:
            self.logger.info(f"WebDriver found at {driver_filepath}")
        return driver_filepath

    def setup_driver_path(self):
        """Sets up the driver path by downloading if necessary."""
        system_platform = "win64" if platform.system() == "Windows" else "linux64" if platform.system() == "Linux" else "mac-x64"
        webdriver_url = self.get_webdriver_download_url(system_platform, self.versions_dict)
        driver_name = os.path.basename(webdriver_url).replace('.zip', '')
        driver_filepath = self.get_driver_filepath(driver_name, download_url=webdriver_url)
        return os.path.abspath(driver_filepath)
