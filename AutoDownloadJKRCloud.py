#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import time
import os

# ================= ATTACH TO EXISTING CHROME =================

chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

download_path = "/tmp"

prefs = {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)
# ================= LOGIN =================

driver.get("https://jkr.eazysaas.com/login.html?redir=https://jkr.eazysaas.com/")
time.sleep(5)

driver.find_element(By.NAME, "user name").send_keys("UMESH")
driver.find_element(By.NAME, "Password").send_keys("UMESH")
driver.find_element(By.XPATH, "//button[contains(text(),'Sign In')]").click()

time.sleep(15)

# ================= OPEN SALES ANALYZER =================

sales_analyzer_link = driver.find_element(
    By.XPATH, "//a[contains(@ui-sref,'SalesAnalyser')]"
)
driver.execute_script("arguments[0].click();", sales_analyzer_link)

time.sleep(12)

# ================= OPEN DATE PICKER =================

date_input = driver.find_element(By.XPATH, "//input[@daterange]")
driver.execute_script("arguments[0].click();", date_input)

time.sleep(2)

# ================= PASTE YESTERDAY =================

yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_range = f"{yesterday} - {yesterday}"

date_input.clear()
date_input.send_keys(date_range)

# ================= CLICK APPLY =================

apply_button = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//div[contains(@class,'daterangepicker') and contains(@style,'display: block')]//button[contains(@class,'applyBtn')]"
    ))
)

driver.execute_script("arguments[0].click();", apply_button)

# ================= SELECT MULTIPLE COLUMNS =================

columns_to_select = [
    "BARCODE","BILL DATE","BILL NO","COMPANY","PRODUCT",
    "PRODUCT GROUP","PURCHASE DATE","PURCHASE RATE",
    "SALE RATE","SALES MAN","SECTION","SIZE",
    "STOCK LOCATION","SUPPLIER","SUPPLIER CITY"
]

for col in columns_to_select:
    driver.find_element(
        By.XPATH, f"//td[normalize-space()='{col}']"
    ).click()

    time.sleep(2)

    driver.find_element(
        By.XPATH, "//div[contains(@id,'checkbox_selector')]"
    ).click()

# ================= OPEN COLUMN DIALOG =================

columns_button = driver.find_element(
    By.XPATH, "//button[contains(@class,'fa-columns')]"
)
columns_button.click()

time.sleep(3)

dialog = driver.find_element(
    By.XPATH, "//div[contains(@class,'ngdialog-message')]"
)

# ================= SELECT CHECKBOXES IN DIALOG =================

for col in columns_to_select:

    row = dialog.find_element(
        By.XPATH,
        f".//tr[.//td[normalize-space()='{col}']]"
    )

    checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
    driver.execute_script("arguments[0].click();", checkbox)

# ================= CLICK VIEW =================

view_button = driver.find_element(
    By.XPATH, "//button[@ng-click='ShowReport()' and contains(text(),'View')]"
)
driver.execute_script("arguments[0].click();", view_button)

time.sleep(5)

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================= EXPORT TO EXCEL =================

excel_button = WebDriverWait(driver, 60).until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[@ng-click=\"exporttoExcel(igrid,'Stock')\"]")
    )
)

driver.execute_script("arguments[0].scrollIntoView();", excel_button)
driver.execute_script("arguments[0].click();", excel_button)
time.sleep(10)
# ================= WAIT FOR DOWNLOAD =================

download_path = "/tmp"
target_folder = "/tmp"

os.makedirs(target_folder, exist_ok=True)

# Wait until Chrome finishes downloading
while any(fname.endswith(".crdownload") for fname in os.listdir(download_path)):
    time.sleep(1)

# Wait until file appears
while True:
    files = [f for f in os.listdir(download_path) if f.endswith((".xls", ".xlsx", ".csv"))]
    if files:
        break
    time.sleep(5)

# ================= RENAME & MOVE FILE =================

import shutil

latest_file = max(
    [os.path.join(download_path, f) for f in files],
    key=os.path.getctime
)

today = datetime.now().strftime("%d.%m.%Y")
extension = os.path.splitext(latest_file)[1]

new_location = os.path.join(target_folder, f"SALES-JKR-{today}{extension}")

shutil.move(latest_file, new_location)

import pandas as pd
import os

file_path = new_location   # downloaded .xls file

# ============================================================
# 1️⃣ READ HTML-BASED XLS FILE (skip STOCK row)
# ============================================================

tables = pd.read_html(file_path, header=1)
df = tables[0]

# ============================================================
# 🏆 2️⃣ DELETE ROW JUST AFTER HEADER
# ============================================================

df = df.iloc[1:]   # removes the unwanted row below header

# ============================================================
# 3️⃣ SAVE AS CLEAN XLSX
# ============================================================

xlsx_path = file_path.replace(".xls", ".xlsx")
df.to_excel(xlsx_path, index=False)

# Delete original file
os.remove(file_path)

print("✅ Clean report saved:", xlsx_path)

from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

file_drive = drive.CreateFile({'title': os.path.basename(xlsx_path)})
file_drive.SetContentFile(xlsx_path)
file_drive.Upload()

print("✅ Uploaded to Google Drive")

