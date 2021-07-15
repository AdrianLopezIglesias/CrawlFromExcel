WebsitePass = "password"

import win32com.client
import math
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True

file = 'M:\\EXCEL.xlsm'

URLS= "https://Website.com

workbook = excel.Workbooks.Open(file)
WSheet = workbook.Worksheets("BaseActual")
workbook.RefreshAll()


options = webdriver.ChromeOptions()
options.add_argument('ignore-certificate-errors')
driver = webdriver.Chrome(chrome_options=options)


driver.get(URLS)
user = driver.find_element_by_id("login_user")
user.send_keys("userName")
pas = driver.find_element_by_id("login_pass")
pas.send_keys(WebsitePass)
login = driver.find_element_by_class_name("submit_input").click()

driver.get("https://Website.com/search")

driver.find_element_by_xpath("/html/body/div[2]/div[1]/div[1]/form/h2[2]").click()

time.sleep(1)
WSheet.UsedRange
LastRow = WSheet.UsedRange.Rows(WSheet.UsedRange.Rows.Count).Row
print(LastRow)

col = WSheet.Range("G1:G" + str(LastRow))
#print(len(col))
for cell in col:
    if cell.value == "Verificar" and WSheet.Range("V" + str(cell.Row)).Value != "" and WSheet.Range("V" + str(cell.Row)).Value != 0:
        try:
            print('Cell value: ')
            print(WSheet.Range("V" + str(cell.Row)).Value)
            WebsiteQuery = math.trunc(WSheet.Range("V" + str(cell.Row)).Value)
            driver.find_element_by_id("list_person_dni").clear()
            dni = driver.find_element_by_id("list_person_dni")
            dni.send_keys(WebsiteQuery)         
            driver.find_element_by_xpath("/html/body/div[2]/div[1]/div/form/input[2]").click()
            print(driver.find_element_by_class_name("list_table").text)
            try:
                WSheet.Range("AO" + str(cell.Row)).Value = driver.find_element_by_class_name("list_table").text
            except:
                WSheet.Range("AO" + str(cell.Row)).Value = "Error al realizar la operación. Comprobar manualmente"
        except:
            WSheet.Range("AO" + str(cell.Row)).Value = "Error al realizar la operación. Comprobar manualmente"
driver.close()
workbook.save()