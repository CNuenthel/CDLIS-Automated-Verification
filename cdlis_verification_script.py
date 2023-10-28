from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import os
from selenium.webdriver.chrome.service import Service
from datetime import date
import code
import csv

# Change to active directory
# os.chdir("C:/Users/cnuenthe/OneDrive - State of North Dakota/Desktop/MultiDriver CDLIS Checker")

# Advise activity
print("Getting Started!")

# Show todays month and day
print(f"Today's date is {date.today()}")


class Driver:
    def __init__(self, driver_data: dict):
        self.first_name = driver_data["First Name"]
        self.last_name = driver_data["Last Name"]
        self.oln = driver_data["OLN"]
        self.dob = driver_data["DOB"]
        self.dl_country = driver_data["Country"]
        self.dl_state = driver_data["State"]


class DriverDataParser:
    def __init__(self, csv_file_path: str):
        self.csv_path = csv_file_path
        self.csv_data = self._read_csv()

    # Read from file
    def _read_csv(self) -> list:
        csv_data = []
        with open(self.csv_path, newline='') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                self.csv_data.append(row)
        return csv_data

    # Create a list of driver objects and return the list
    def parse_driver(self, driver_data: dict) -> Driver:
        driver = Driver(driver_data)
        return driver


class CdlisWebCrawler:
    def __init__(self):
        self.driver = self._build_driver()
        self.page_flag = None
        self.login = None
        self.password = None

    def _build_driver(self) -> webdriver:
        # Create options object and add detach to keep window open
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)

        # Create service manager, this is a weak instantiation, this will break if not used on my work computer
        # service = Service(
        #     "C:/Users/cnuenthe/OneDrive - State of North Dakota/Desktop/MultiDriver CDLIS Checker/chromedriver-win64/chromedriver.exe")
        service = Service("chromedriver.exe")
        driver = webdriver.Chrome(service=service, options=chrome_options)  # Create webdriver object
        return driver

    # navigate to CDLIS website
    def navigate_to_cdlis_website(self):
        self.driver.get("https://cdlis.dot.gov/")
        self.page_flag = "Splash Page"

    # complete authorization splash page
    def navigate_through_splash_page(self):
        if self.page_flag == "Splash Page":
            self.driver.find_element(By.NAME, "btnAttentionIAgree").click()
            self.driver.find_element(By.NAME, "btnPrivacyIAgree").click()
            self.page_flag = "Login Page"
            return

        print(f"You must be on the splash page to navigate through it. Your page flag is currently: {self.page_flag}")

    # collect and enter login credentials
    def collect_credentials(self):
        self.login = input("Please input your CDLIS username:\n")
        self.password = input("Please input your CDLIS password:\n")

    def enter_credentials(self):
        if self.login is None or self.password is None:
            print("A username and password has not been collected")
            return

        if self.page_flag == "Login Page":
            self.driver.find_element(By.NAME, "UserName").send_keys(self.login)
            self.driver.find_element(By.NAME, "Password").send_keys(self.password)
            self.driver.find_element(By.XPATH, '//*[@id="loginForm"]/form/fieldset/input')
            return

        print(f"You must be on the login page to enter credentials. Your page flag is currently: {self.page_flag}")

    # select query filters
    def select_query_filters(self, driver_data: Driver):
        # collect a driver data object
        pass

    # input driver data

    # click search button

    # take a snapshot of given information

    # navigate back to query filter page


code.interact(local=locals())