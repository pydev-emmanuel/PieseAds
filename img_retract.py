import os.path
import io
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
from openpyxl import load_workbook
from bs4 import BeautifulSoup as bs

workbook = load_workbook("C:\\Users\\Gh0sT\\Desktop\\WORK\\placute_frana_MINTEX\\placute_frana_MINTEX.xlsx")
worksheet = workbook["Sheet1"]
column = worksheet["A"]
column_list = [column[x].value for x in range(len(column))]
print(len(column_list))
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
action = ActionChains(driver)
for item in column_list:
    if os.path.isfile(f"C:\\Users\\Gh0sT\\Desktop\\WORK\\discuri_MINTEX\\placute_frana_MINTEX_info\\{item}.html"):
        print("Produs extractat")
    else:
        print(f"Produs in proces: {item}")
        driver.get(f"https://www.mintex.brakebook.com/bb/mintex/ro/{item}_402/datasheet.xhtml")
        sleep(5)
        try:
            superseded_item = driver.find_element(By.XPATH, "//div[@class='supersededBy']/a").get_attribute("text")
            driver.get(f"https://www.mintex.brakebook.com/bb/mintex/ro/{superseded_item}_82/datasheet.xhtml")
            sleep(5)
        except:
            pass
        try:
            driver.find_element(By.XPATH, "//div[@class='modelLabel']/a").click()
        except:
            pass
        sleep(5)
        src_code = driver.find_element(By.XPATH, "//td[@class='datasheetBody']").get_attribute("innerHTML")
        src_encoded = src_code.encode("utf-8")
        with io.open(f"C:\\Users\\Gh0sT\\Desktop\\WORK\\placute_frana_MINTEX\\placute_frana_MINTEX_info\\{item}.html", "w", encoding="utf-8") as html:
            html.write(src_code)



# for item in column_list:
#     driver.get("https://www.mintex.brakebook.com/bb/mintex/ro/applicationSearch.xhtml")
#     sleep(5)
#     driver.find_element(By.XPATH, "//input[@id='search_keywords']").click()
#     sleep(1)
#     action.key_down(Keys.CONTROL).send_keys("A").key_up(Keys.CONTROL).key_down(Keys.DELETE).perform()
#     sleep(5)
#     search_tag = driver.find_element(By.XPATH, "//input[@id='search_keywords']")
#     enter_tag = driver.find_element(By.XPATH, "//a[@title='Aici, puteţi utiliza un * pentru specificarea globală a unui fişier, de exemplu 2501* pentru toate articolele al căror număr articol începe cu 2501.']")
#     search_tag.click()
#     search_tag.send_keys(item)
#     enter_tag.click()
#     sleep(10)
