import xlrd
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from values import inputs


# Test Case 20_a
class Exports:

    def __init__(self, driver):  # initialize each WebElement here
        self.driver = driver
        wait = WebDriverWait(self.driver.instance, 10)
        # Export Form data variables
        self.data_dropdown = None  # Data dropdown
        self.view_all_link = None  # View All link
        self.export_form_data_button = None  # click form exports
        self.prepare_export_button = None  # click prepare exports
        self.download_button = None  # click download
        self.find_data_by_ID_link = None  # Click findDataByID link
        self.find_data_by_ID_textbox = None  # Find data by ID textbox
        self.find_data_by_ID_button = None
        self.view_FormID = None
        self.womanName_HQ = None  # Property 'Woman's name' value on HQ

        # Export Case data variables
        self.export_case_data_link = None  # Export Case Data on the left panel
        self.add_export_button = None  # Add Export button
        self.export_case_data_button = None
        self.prepare_case_export_button = None
        self.case_download_button = None

    def form_exports(self):

        self.data_dropdown = WebDriverWait(self.driver.instance, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, '//*[@id="ProjectDataTab"]/a')))
        self.data_dropdown.click()
        print("Data drop down clicked")
        time.sleep(2)

        self.view_all_link = WebDriverWait(self.driver.instance, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, '//*[@id="ProjectDataTab"]/ul/li[6]/a')))
        self.view_all_link.click()
        print("View All link clicked")
        # time.sleep(2)

        self.export_form_data_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="export-list"]/div[2]/div/div[2]/table/tbody/tr/td[2]/a[1]')))
        self.export_form_data_button.click()
        print("Export form button clicked")
        # time.sleep(2)

        self.prepare_export_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="download-export-form"]/form/div[2]/div/div[2]/div[1]/button')))
        self.prepare_export_button.click()
        print("Prepare Form Export button clicked")
        # time.sleep(2)

        try:
            self.download_button = WebDriverWait(self.driver.instance, 10).until(
                EC.visibility_of_element_located((
                    By.XPATH, '//*[@id="download-progress"]/div/div/div[2]/div[1]/form/a')))
            self.download_button.click()
            print("Download form button clicked")
        except Exception as e:
            print(e)
            print("Download task failed to start")
        finally:
            time.sleep(2)

    def validate_downloaded_form_exports(self):
        path = inputs.download_path
        os.chdir(path)
        files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        oldest = files[0]
        newest = files[-1]
        print("Oldest:", oldest)
        print("Newest:", newest)
        # print ("All by modified oldest to newest:","\n".join(files))
        # check if size of file is 0
        if os.stat(newest).st_size == 0:
            print('File is empty')
        else:
            print('File is not empty')

        # Test Case 22_a
        # Reading an excel file using Python
        # To open Workbook
        wb = xlrd.open_workbook(newest)
        sheet = wb.sheet_by_index(0)
        # Extract a particular row value
        # print(sheet.row_values(1))
        # For row 0 and column 1, extracting first formID
        formID_colName = sheet.cell_value(0, 1)
        formID = sheet.cell_value(1, 1)
        print(formID_colName, ": ", formID)
        womanName_colName = sheet.cell_value(0, 2)
        womanName_excel = sheet.cell_value(1, 2)
        print("Woman's name in Excel- ", womanName_colName, ": ", womanName_excel)
        # return formID
        # def find_data_by_ID(self,validate_downloaded_form_exports):
        self.find_data_by_ID_link = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="hq-sidebar"]/nav/ul[1]/li[4]/a')))
        self.find_data_by_ID_link.click()
        print("Find data by ID link clicked")
        self.find_data_by_ID_textbox = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="find-form"]/div[2]/div[1]/input')))
        self.find_data_by_ID_textbox.clear()
        self.find_data_by_ID_textbox.send_keys(formID)
        print("Form ID fed in the textbox")
        self.find_data_by_ID_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="find-form"]/div[2]/div[2]/button')))
        self.find_data_by_ID_button.click()
        print("find_data_by_ID_button clicked")

        self.view_FormID = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="find-form"]/div[2]/div[1]/div[2]/a')))
        self.view_FormID.click()
        print("view_FormID link clicked")

        # Switch tab logic is required here

        self.womanName_HQ = self.driver.find_element_by_xpath(
            '//*[@id="form-data"]/div[3]/div/div/table/tbody/tr[2]/td[2]').text
        print("Woman's name on HQ")
        print(self.womanName_HQ)

        if (womanName_excel == self.womanName_HQ):
            print("Values match!")
        else:
            print("Values don't match")

    def case_exports(self):
        self.export_case_data_link = WebDriverWait(self.driver.instance, 10).until(
            EC.element_to_be_clickable((
                By.XPATH, '//*[@id="hq-sidebar"]/nav/ul[1]/li[2]/a')))
        self.export_case_data_link.click()
        print("export_case_data_link clicked")
        time.sleep(2)

        self.export_case_data_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="export-list"]/div[2]/div/div[2]/table/tbody/tr/td[3]/a[1]')))
        self.export_case_data_button.click()
        print("Export Case button clicked")
        # time.sleep(2)

        self.prepare_case_export_button = WebDriverWait(self.driver.instance, 10).until(
            EC.visibility_of_element_located((
                By.XPATH, '//*[@id="download-export-form"]/form/div[2]/div/div[2]/div[1]/button')))
        self.prepare_case_export_button.click()
        print("Prepare Case Export button clicked")
        time.sleep(2)

        try:
            self.case_download_button = WebDriverWait(self.driver.instance, 10).until(
                EC.visibility_of_element_located((
                    By.XPATH, '//*[@id="download-progress"]/div/div/div[2]/div[1]/form/a')))
            self.case_download_button.click()
            print("Download Case button clicked")
        except Exception as e:
            print(e)
            print("Download task failed to start")
        finally:
            time.sleep(2)
            
    #def validate_downloaded_form_exports(self):
        #WIP
    # Add Export code to be added here
    # self.add_export_button = WebDriverWait(self.driver.instance, 10).until(
    #   EC.element_to_be_clickable((
    #      By.XPATH, '//*[@id="create-export"]/p/a')))
    # self.add_export_button.click()
    # print("Add Export button clicked")
    # time.sleep(2)

