#!/usr/bin/env python

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import os
import time
import shutil
import pandas as pd

# ================= GET LOGIN FROM GITHUB SECRETS =================

USERNAME = os.environ["USERNAME"]
PASSWORD = os.environ["PASSWORD"]

# ================= SETUP HEADLESS CHROME =================

download_path = "/tmp"

chrome_options = Options()

chrome_options.add_argument("--headless=new")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option("useAutomationExtension", False)

prefs = {
    "download.default_directory": "/tmp",
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(options=chrome_options)

# Remove automation flag
driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

wait = WebDriverWait(driver, 40)

# ================= LOGIN =================

driver.get("https://arniraja.eazysaas.com/login.html?redir=https://arniraja.eazysaas.com/")

wait.until(EC.presence_of_element_located((By.NAME, "user name"))).send_keys(USERNAME)
driver.find_element(By.NAME, "Password").send_keys(PASSWORD)
driver.find_element(By.XPATH, "//button[contains(text(),'Sign In')]").click()

print("✅ Logged in")

driver.save_screenshot("/tmp/after_login.png")
print("Screenshot saved")
time.sleep(15)

# ================= OPEN SALES ANALYZER =================

sales_analyzer_link = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(@ui-sref,'SalesAnalyser')]"))
)

driver.execute_script("arguments[0].click();", sales_analyzer_link)

print("✅ Opened Sales Analyzer")

# ================= SET DATE =================

date_input = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//input[@daterange]"))
)

driver.execute_script("arguments[0].click();", date_input)

yesterday = (datetime.now() - timedelta(days=1)).strftime("%d/%m/%Y")
date_range = f"{yesterday} - {yesterday}"

date_input.clear()
date_input.send_keys(date_range)

apply_button = wait.until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//div[contains(@class,'daterangepicker') and contains(@style,'display: block')]//button[contains(@class,'applyBtn')]"
    ))
)

driver.execute_script("arguments[0].click();", apply_button)

print("✅ Date applied")

# ================= OPEN COLUMN DIALOG =================

columns_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'fa-columns')]"))
)

driver.execute_script("arguments[0].click();", columns_button)

dialog = wait.until(
    EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'ngdialog-message')]"))
)

print("✅ Column dialog opened")

columns_to_select = [
    "BARCODE","BILL DATE","BILL NO","COMPANY","PRODUCT",
    "PRODUCT GROUP","PURCHASE DATE","PURCHASE RATE",
    "SALE RATE","SALES MAN","SECTION","SIZE",
    "STOCK LOCATION","SUPPLIER","SUPPLIER CITY"
]

for col in columns_to_select:
    row = dialog.find_element(
        By.XPATH,
        f".//tr[.//td[normalize-space()='{col}']]"
    )

    checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
    driver.execute_script("arguments[0].click();", checkbox)

print("✅ Columns selected")

# ================= CLICK VIEW =================

view_button = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[@ng-click='ShowReport()' and contains(text(),'View')]")
    )
)

driver.execute_script("arguments[0].click();", view_button)

print("✅ Report generated")

# ================= EXPORT TO EXCEL =================

excel_button = wait.until(
    EC.element_to_be_clickable(
        (By.XPATH, "//button[@ng-click=\"exporttoExcel(igrid,'Stock')\"]")
    )
)

driver.execute_script("arguments[0].scrollIntoView();", excel_button)
driver.execute_script("arguments[0].click();", excel_button)

print("✅ Export clicked")

# ================= WAIT FOR DOWNLOAD =================

while any(fname.endswith(".crdownload") for fname in os.listdir(download_path)):
    time.sleep(1)

files = [f for f in os.listdir(download_path) if f.endswith((".xls", ".xlsx", ".csv"))]

latest_file = max(
    [os.path.join(download_path, f) for f in files],
    key=os.path.getctime
)

print("✅ File downloaded:", latest_file)

# ================= CLEAN FILE =================

tables = pd.read_html(latest_file, header=1)
df = tables[0]

df = df.iloc[1:]

today = datetime.now().strftime("%d.%m.%Y")
xlsx_path = os.path.join(download_path, f"SALES-RR-{today}.xlsx")

df.to_excel(xlsx_path, index=False)

print("✅ Clean report saved:", xlsx_path)

# ================= SEND EMAIL =================

import smtplib
from email.message import EmailMessage

EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASS = os.environ["EMAIL_PASS"]
EMAIL_TO = os.environ["EMAIL_TO"]

msg = EmailMessage()
msg["Subject"] = "Daily Sales Report"
msg["From"] = EMAIL_USER
msg["To"] = EMAIL_TO
msg.set_content("Attached is the daily sales report.")

# Attach file
with open(xlsx_path, "rb") as f:
    file_data = f.read()
    file_name = os.path.basename(xlsx_path)

msg.add_attachment(
    file_data,
    maintype="application",
    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    filename=file_name
)

# Send Email
with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
    smtp.login(EMAIL_USER, EMAIL_PASS)
    smtp.send_message(msg)

print("✅ Email sent successfully")

driver.quit()
