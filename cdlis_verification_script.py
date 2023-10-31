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
import openpyxl
import usaddress
from typing_extensions import dataclass_transform
from dataclasses import dataclass


WORKING_DIRECTORY = os.path.dirname(os.path.abspath(__file__))


# Advise activity
print("Getting Started!")

# Show todays month and day
print(f"Today's date is {date.today()}")

# Helper function to format date of birth data, splitting standard DOB "01/01/2000" to a dict {month: "01", day: "01", year: "2000"}
def parse_dob_str(dob_timestamp) -> dict:
    stringify_dob = dob_timestamp.strftime("%m/%d/%Y")
    month, day, year = stringify_dob.split("/")
    formatted_date = {
        "month": month,
        "day": day,
        "year": year
    }
    return formatted_date


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
        chrome_options.add_argument("--start-maximized")

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

    def snapshot_driver_info(self, driver_data: Driver): # TODO Parse document information to excel sheet
        # Get table data from CDLIS website, creates a list of webelement objects
        table_data = self.crawler.find_elements(By.CLASS_NAME, "reportTable")
        cwdp = CdlisWebdataParser(table_data)
        
        if driver_data.dl_country == "Canada":
            formatted_webdata = cwdp.parse_doc_to_lists()
            driver_table_parser = DriverTableParser()
            print(driver_table_parser.parse_canada(formatted_webdata))

    # navigate back to query filter page
    def return_to_search_page(self):
        self.crawler.execute_script("document.body.style.zoom='100%'")  # Rezoom document or click coordinate will fail to execute
        self.crawler.find_element(By.CLASS_NAME, "Cancel").click()


@dataclass
class CanDriverData:
    first_name: str = ""
    middle_name: str = ""
    last_name: str = ""
    suffix: str = ""
    ssn: str = ""
    dob: str = ""
    height: str = ""
    weight: str = ""
    eye_color: str = ""
    sex: str = ""
    address: str = ""
    city: str = ""
    state: str = ""
    zip: str = ""
    jurisdiction: str = ""
    oln: str = ""
    issue_date: str = ""
    expiration_date: str = ""
    commercial_class: str = ""
    noncommercial_class: str = ""
    commercial_status: str = ""
    noncommercial_status: str = ""
    withdrawal_action: str = ""
    endorsement: str = ""
    convictions: int = ""
    accidents: int = ""
    withdrawals: int = ""
    permits: int = ""
    license_restrictions: int = ""


class DriverTableParser:
    """ This class holds methods to parse CDLIS string table data for use in formatting to an excel output """
    def __init__(self):
        self.output_file = "DriverDataOutput.xlsx"

    def parse_canada(self, can_table_data: list[str]) -> CanDriverData:
        """ Parses driver data from a canadian driver CDLIS table """
        driver = CanDriverData()

        # Parse name information to dictionary, row 1
        # Split string data, first row to components
        name_split = can_table_data[0].split(" ")

        # Parse data 
        driver.first_name = name_split[0]
        driver.middle_name = name_split[1]
        driver.last_name = name_split[2]
        driver.suffix = name_split[3]

        # Parse DOB information to dictionary, row 2
        # Split string data, second row to components
        dob_split = can_table_data[1].split(" ")

        # Parse data
        driver.ssn = dob_split[0]
        driver.dob = dob_split[1]
        driver.height = dob_split[2]

        # Match length of split data due to uncertainty of information. If weight and/or eye color is present the list will dynamically change
        # from the input data
        match len(dob_split):
            case 6:
                driver.weight = dob_split[3]
                driver.eye_color = dob_split[4]
                driver.sex = dob_split[5]
            case 5:
                driver.sex = dob_split[4]
            case 4:
                driver.sex = dob_split[3]


        # Parse address information to dictionary, row 3
        # Use usaddress module to parse address data 
        address = can_table_data[2]
        components = usaddress.parse(address)

        # Create lists to assign parsed data as usaddress parses each string to a category in a list -> [(data, category), ..]
        street_address = []
        city = []
        state = []
        zip = []

        # Parse categories to appropriate list
        for item in components:
            if item[1] in ["AddressNumber", "StreetName", "StreetNamePostType"]:
                street_address.append(item[0])
            elif item[1] in ["PlaceName"]:
                city.append(item[0])
            elif item[1] in ["StateName"]:
                state.append(item[0])
            elif item[1] in ["ZipCode"]:
                zip.append(item[0])
            else:
                # Advise of unassigned categories for further processing if necessary
                print(f"{item[1]} is unassigned")

        # Assign parsed data to DriverData object
        driver.address = " ".join(street_address)
        driver.city = " ".join(city)
        driver.state = " ".join(state)
        driver.zip = " ".join(state)

        # Parse DL information to dictionary, row 4
        # Split string data, fourth row to components
        dl_split = can_table_data[3].split(" ")

        driver.jurisdiction = dl_split[0]
        driver.oln = dl_split[1]
        driver.issue_date = dl_split[2]
        driver.expiration_date = dl_split[3]
        driver.commercial_class = dl_split[4]
        driver.noncommercial_class = dl_split[5]
        driver.commercial_status = dl_split[6]
        driver.noncommercial_status = dl_split[7]
        driver.withdrawal_action = dl_split[8]

        # Parse endorsement information, row 5
        driver.endorsement = can_table_data[4]

        # Parse confiction information, row 6
        # Split string data, sixth row to components, stringify expected num data
        convictions_split = [str(item) for item in can_table_data[5].split(" ")]

        driver.convictions = convictions_split[0]
        driver.accidents = convictions_split[1]
        driver.withdrawals = convictions_split[2]
        driver.permits = convictions_split[3]
        driver.license_restrictions = convictions_split[4]

        return driver


    # __________________________________________________________________________________________________________________________________
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
