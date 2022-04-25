import os.path

from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from openpyxl import load_workbook

workbook = load_workbook("C:\\Users\\Gh0sT\\Desktop\\work_LUK_flywheel\\volante_LUK.xlsx")
worksheet = workbook["Sheet1"]
column = worksheet["A"]
column_list = [column[x].value for x in range(len(column))]
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://aftermarket.schaeffler.com/en/catalog")
sleep(5)
for code in column_list:
    print(code)
    driver.find_element(By.XPATH, "//input[@class='mat-input-element mat-form-field-autofill-control ng-tns-c62-4 ng-untouched ng-pristine ng-invalid cdk-text-field-autofill-monitored']").click()
    sleep(1)
    driver.find_element(By.XPATH, "//input[@class='mat-input-element mat-form-field-autofill-control ng-tns-c62-4 ng-untouched ng-pristine ng-invalid cdk-text-field-autofill-monitored']").send_keys(code)
    sleep(1)
    driver.find_element(By.XPATH, "//button[@class='mat-focus-indicator mat-icon-button mat-button-base mat-accent ng-tns-c62-4']").click()
    sleep(5)
    driver.find_element(By.XPATH, "//mat-card[@class='mat-card mat-focus-indicator mat-elevation-z0 saam-mat-card--box-shadow saam-mat-card--linked']").click()
    sleep(5)
    for tag in driver.find_elements(By.TAG_NAME, "div"):
        if "saam-flex-row product-detail__body ng-tns" in tag.get_attribute("class"):
            img_src = tag.get_attribute("innerHTML")
    html = open(f"C:\\Users\\Gh0sT\\Desktop\\work_LUK_flywheel\\volante_LUK_photo\\{code}.html", "w")
    html.write(img_src)
    html.close()
    sleep(1)
    driver.get("https://aftermarket.schaeffler.com/en/catalog")
    sleep(5)
