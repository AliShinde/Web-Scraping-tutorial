import pandas as pd
import requests
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
import time

# Add your own installation path of chrome driver
chrome_driver_path = "C:\\Chrome_driver\\chromedriver.exe"
# Your working directory where the excel file is stored
file_path = "C:\\Clemson\\DVL\\Web scraping\\DVL_Outreach.xlsx"
book = load_workbook(file_path)
service = Service(chrome_driver_path, log_output=os.devnull)

driver = webdriver.Chrome(service=service)


#Check request status before making a request
url = "https://www.clemson.edu/cafls/forestry-environmental-conservation/directory/faculty.html"



def check_url_status(url) -> bool:
    try:
        res = requests.get(url, timeout=5)
        if res.status_code >= 200 and res.status_code < 300:
            print(res.status_code)
            return True
        else:
            print(f"Error, the status code is {res.status_code}")
            return False
    except requests.exceptions.RequestException as e:
        print(f"Major error occured for {url} with text {e}")
        return False

isValidResponse = check_url_status(url)

data_body = []
if isValidResponse:
    driver.get(url)
    time.sleep(10)
    data_table = driver.find_element(By.TAG_NAME, "table")
    rows = data_table.find_elements(By.TAG_NAME, "tr")

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if not cells:
            # Header row
            headers = row.find_elements(By.TAG_NAME, "th")
            list_heading = [head.text.strip() for head in headers]
        else:
            data_subbody = [cell.text.strip() for cell in cells]
            data_body.append(data_subbody)

dataFrame = pd.DataFrame(data=data_body, columns=list_heading)


with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
    dataFrame.to_excel(writer, sheet_name="History Geography", index=False)
