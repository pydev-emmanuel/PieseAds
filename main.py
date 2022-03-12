# https://ro.e-cat.intercars.eu
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
from openpyxl import load_workbook

workbook = load_workbook("C:\\Users\\Gh0sT\\Desktop\\LESJOFORS.xlsx")
worksheet = workbook["Sheet1"]
column = worksheet["A"]
column_list = [column[x].value for x in range(len(column))]
column_set = set(column_list)
column_modified = list(column_set)
print(len(column_modified))
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get("https://ro.e-cat.intercars.eu")
driver.find_element(By.XPATH, "//input[@class='form-control form-control  bf-required']").send_keys("emanuel_b1998@yahoo.com")
driver.find_element(By.XPATH, "//input[@class='form-control form-control']").send_keys("Hala@Bilca2021")
sleep(1)
driver.find_element(By.XPATH, "//button[@class='btn btn-default btn col-sm-12']").click()
sleep(10)
for code in column_list[123:]:
    html_parameters = None
    html_other_numbers = None
    html_applications = None
    driver.find_element(By.XPATH, "//input[@class='header__searchinput js-search-field-input js-keyboardable-search js-onboarding-homepage-mainsearchinput ui-autocomplete-input']").send_keys(code)
    sleep(1)
    driver.find_element(By.XPATH, "//div[@class='header__searchbuttonsubmit js-search-button-submit']").click()
    sleep(5)
    try:
        link = driver.find_element(By.XPATH, "//div[@class='listingcollapsed__activenumbercontainer']/a").get_attribute("href")
        driver.get(link)
    except:
        pass

    for tab in driver.find_elements(By.XPATH, "//div[@class='tabs__item']"):
        if tab.text == "PARAMETERS":
            tab.click()
            sleep(2)
        if "OTHER" in tab.text:
            tab.click()
            sleep(2)
    for tab in driver.find_elements(By.XPATH, "//div[@class='tabs__item']"):
        if tab.text == "APPLICATIONS":
            tab.click()
            sleep(2)
            for branch in driver.find_elements(By.XPATH, "//div[@class='tree__branch']"):
                branch.click()
            sleep(2)
            for leaf in driver.find_elements(By.XPATH, "//div[@class='tree__leaf js-tree-trigger']"):
                leaf.click()
            sleep(2)
            html_applications = driver.find_element(By.XPATH, "//div[@class='layoutproductdetails__tabs layoutproductdetails__tabs--doublerow productprice--productdetails productretailprice--productdetails']").get_attribute("innerHTML")
            price = driver.find_element(By.XPATH, "//div[@class='buybox js-onboarding-productdetails-buybox buybox--']").get_attribute("innerHTML")
            html = open(f"C:\\Users\\Gh0sT\\Desktop\\HTML_LESJOFORS\\{code}.html", "w")
            html.write(html_applications)
            html.write(price)
            html.close()
            sleep(3)

