"""
Limitations
    Currently this application cannot be run using headless mode as snapshots are taken of a non-maximized screen.
    I have not currently found a solution to creating a maximized screen in headless mode.
"""


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import os
from selenium.webdriver.chrome.service import Service
from datetime import date
import pandas as pd
from selenium.webdriver.support.ui import Select
from selenium.webdriver.remote.webelement import WebElement
import json
import time

WORKING_DIRECTORY = os.path.dirname(os.path.abspath(__file__))


# Advise activity
print("Getting Started!")

# Show todays month and day
print(f"Today's date is {date.today()}")


# Helper function to format date of birth data, splitting standard DOB "01/01/2000" to a dict {month: "01",
# day: "01", year: "2000"}
def parse_dob_str(dob_timestamp) -> dict:
    stringify_dob = dob_timestamp.strftime("%m/%d/%Y")
    month, day, year = stringify_dob.split("/")
    formatted_date = {
        "month": month,
        "day": day,
        "year": year
    }
    return formatted_date


# Helper function to rename pdf downloads from cdlis to driver name
def change_last_pdf_name(folder_path, new_name):
    # Get a list of all PDF files in the folder
    pdf_files = [file for file in os.listdir(folder_path) if file.lower().endswith('.pdf')]
    # Sort the PDF files by modification time (latest first)
    pdf_files.sort(key=lambda x: os.path.getmtime(os.path.join(folder_path, x)), reverse=True)

    if pdf_files:
        # Get the path of the last downloaded PDF
        last_pdf_path = os.path.join(folder_path, pdf_files[0])
        # Generate the new path with the desired name
        new_pdf_path = os.path.join(folder_path, f"{new_name}.pdf")
        # Rename the file
        os.rename(last_pdf_path, new_pdf_path)
        print(f"Renamed {pdf_files[0]} to {new_name}.pdf")
    else:
        print("No PDF files found in the folder.")


class Driver:
    def __init__(self, first_name, last_name, oln, dob, country, state):
        self.first_name = first_name
        self.last_name = last_name
        self.oln = oln
        self.dob = dob
        self.dl_country = country
        self.dl_state = state


class DriverDataParser:
    def __init__(self, csv_file_path: str):
        self.csv_path = csv_file_path
        self.df_data = self._read_xlsx()
        self.driver_pool = self._create_driver_objects()

    # Read from file
    def _read_xlsx(self) -> pd.core.frame.DataFrame:
        df = pd.read_excel(self.csv_path)
        return df

    def _create_driver_objects(self):
        driver_pool = []
        for i, row in self.df_data.iterrows():
            driver_pool.append(Driver(
                row["First Name"],
                row["Last Name"],
                row["OLN"],
                row["DOB"],
                row["Country"],
                row["State"]
            ))
        return driver_pool

    # Create a list of driver objects and return the list
    def get_driver(self) -> Driver or bool:
        if self.driver_pool:
            return self.driver_pool.pop()
        return False


class CdlisWebdataParser:
    def __init__(self, web_document_html: list[WebElement]):
        self.web_doc = web_document_html

    def parse_doc_to_lists(self) -> list[str]:
        string_data = " ".join([item.text for item in self.web_doc])
        split_string_data = [item for i, item in enumerate(string_data.splitlines()) if item != "" and i%2 == 0]
        return split_string_data


class CdlisWebCrawler:
    def __init__(self, driver_data_parser: DriverDataParser):
        self.data_parser = driver_data_parser
        self.crawler = self._build_crawler()
        self.failed_searches = []
        self.login = None
        self.password = None

    def _build_crawler(self) -> webdriver:
        # Create options object and add detach to keep window open
        chrome_options = Options()
        settings = {
            "recentDestinations": [{
                "id": "Save as PDF",
                "origin": "local",
                "account": "",
            }],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
            "isLandscapeEnabled": True
        }
        prefs = {'printing.print_preview_sticky_settings.appState': json.dumps(settings),
                 'savefile.default_directory': os.path.join(WORKING_DIRECTORY, 'output')}
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--kiosk-printing")

        # Create service manager, this is a weak instantiation, this will break if not used on my work computer
        service = Service(os.path.join(WORKING_DIRECTORY, "chromedriver.exe"))
        crawler = webdriver.Chrome(service=service, options=chrome_options)  # Create webdriver object
        return crawler

    # navigate to CDLIS website
    def navigate_to_cdlis_website(self):
        self.crawler.get("https://cdlis.dot.gov/")

    # complete authorization splash page
    def navigate_through_splash_page(self):
        self.crawler.find_element(By.NAME, "btnAttentionIAgree").click()
        self.crawler.find_element(By.NAME, "btnPrivacyIAgree").click()
        return

    def enter_credentials(self):

        while True:
            self.login = input("Please input your CDLIS username:\n")
            self.password = input("Please input your CDLIS password:\n")

            self.crawler.find_element(By.NAME, "UserName").send_keys(self.login)
            self.crawler.find_element(By.NAME, "Password").send_keys(self.password)
            self.crawler.find_element(By.XPATH, '//*[@id="loginForm"]/form/fieldset/input').click()

            try:
                self.crawler.find_element(By.NAME, "UserName")
                self.crawler.find_element(By.NAME, "UserName").clear() # Clear incorrect inputs
                self.crawler.find_element(By.NAME, "Password").clear()
                print("Sorry, your credentials were not validated, please try again")
            except NoSuchElementException:  # Indicates a successful login
                break

    # select query filters
    def select_query_filters(self, driver_data: Driver):
        # CDLIS webpage dropdown for territories uses a code system to select in dropdown
        query_library = {
            "Canada": "CN",
            "United States of America": "US",
            "Mexico": "MX",
            "Other Countries": "OTH",
            "United States Territories": "US-T"
        }

        territory_dropdown = Select(self.crawler.find_element(By.ID, "ddlCountryFilter"))
        territory_dropdown.select_by_value(query_library[driver_data.dl_country])
        jurisdiction_dropdown = Select(self.crawler.find_element(By.ID, "ddlJurisdiction"))
        self.crawler.implicitly_wait(1)  # Hold the page to update jurisdictional drop down options
        jurisdiction_dropdown.select_by_value(driver_data.dl_state)
        self.crawler.find_element(By.ID, "btnStartFilter").click()

    # input driver data
    def fill_driver_data(self, driver_data: Driver): # TODO add print statement to show driver being checked
        print(f"Checking driver: {driver_data.last_name}, {driver_data.first_name}")

        # Fill in OLN
        oln_input = self.crawler.find_element(By.ID, "DriverLicense")
        oln_input.send_keys(driver_data.oln)

        # Fill in last name
        ln_input = self.crawler.find_element(By.ID, "LastName")
        ln_input.send_keys(driver_data.last_name)

        # Fill in first name
        fn_input = self.crawler.find_element(By.ID, "FirstName")
        fn_input.send_keys(driver_data.first_name)

        # Format DOB info
        dob_dict = parse_dob_str(driver_data.dob)

        # Fill in DOB
        dob_month = self.crawler.find_element(By.ID, "DateOfBirthMonth")
        dob_month.send_keys(dob_dict["month"])

        dob_day = self.crawler.find_element(By.ID, "DateOfBirthDay")
        dob_day.send_keys(dob_dict["day"])

        dob_year = self.crawler.find_element(By.ID, "DateOfBirthYear")
        dob_year.send_keys(dob_dict["year"])

    def search_driver(self, driver_data: Driver) -> bool:
        self.crawler.find_element(By.XPATH, '//*[@id="RequestForm"]/fieldset/ol/li[6]/input[1]').click()

        # Verify if driver was found with given information
        try:
            self.crawler.find_element(By.ID, "DriverLicense")
            print(f"Driver not found with given credentials: {driver_data.last_name},{driver_data.first_name}")
            self.crawler.find_element(By.CLASS_NAME, "Cancel valid")
            return False
        except NoSuchElementException:
            return True

    def snapshot_driver_info(self, driver_data: Driver):
        self.crawler.execute_script("document.body.style.zoom='75%'")
        self.crawler.execute_script('window.print();')
        change_last_pdf_name(os.path.join(WORKING_DIRECTORY, "output"), f"{driver_data.last_name}_{driver_data.first_name}")
        time.sleep(2)

    # navigate back to query filter page
    def return_to_search_page(self):
        self.crawler.execute_script("document.body.style.zoom='100%'")
        self.crawler.find_element(By.CLASS_NAME, "Cancel").click()


# ______________________________________________________________________________________________________________________
def main():
    # Instantiate driver data parser to read driver data from excel sheet
    ddp = DriverDataParser(os.path.join(WORKING_DIRECTORY, "DriverData.xlsx"))

    # Create webcrawler and provide data parser to class
    cw = CdlisWebCrawler(ddp)

    # Navigate to CDLIS website, collect credentials, and login
    cw.navigate_to_cdlis_website()
    cw.navigate_through_splash_page()
    cw.enter_credentials()

    # Verify there are drivers to check and run the cycle on all drivers listed in excel
    while True:
        drv = cw.data_parser.get_driver() # Obtain a driver from the Excel data parser

        if drv: # Collect driver information from CDLIS
            cw.select_query_filters(drv)
            cw.fill_driver_data(drv)

            # If driver search fails, cut block and return to filter page
            if cw.search_driver(drv):
                cw.snapshot_driver_info(drv)
                cw.return_to_search_page()
            else:
                continue

        # End the cycle when driver data returns false
        else:
            break

    print("Driver checks complete! Please review the output folder for collected data")

    #failed_drivers = "\n".join(cw.failed_searches)
    #print(f"Failed driver checks: {failed_drivers}")

    cw.crawler.quit()









if __name__ == "__main__":
    main()
