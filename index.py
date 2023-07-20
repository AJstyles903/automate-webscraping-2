from time import sleep
from selenium import webdriver
from random import randint
from selenium.webdriver.common.by import By
from openpyxl import Workbook

driver = webdriver.Chrome('chromedriver.exe')
driver.get('https://acme-test.uipath.com/login')
print ("Opened Website")

sleep(randint(1,4))
driver.maximize_window()
driver.find_element('name','email').send_keys('dhruval@gmail.com')
driver.find_element('name','password').send_keys('Dhruval@123')
driver.find_element(By.CSS_SELECTOR, 'body > div > div.main-container > div > div > div > form > button').click()
driver.find_element(By.CSS_SELECTOR, '#dashmenu > div:nth-child(2) > a > button').click()

numOfHrows=len(driver.find_elements(By.XPATH,'/html/body/div/div[2]/div/table/tbody/tr[1]'))
numOfHcolumn=len(driver.find_elements(By.XPATH,'/html/body/div/div[2]/div/table/tbody/tr[1]/th'))
numOfrows=len(driver.find_elements(By.XPATH,'/html/body/div/div[2]/div/table/tbody/tr'))
numOfcolumn=len(driver.find_elements(By.XPATH,'/html/body/div/div[2]/div/table/tbody/tr[2]/td'))
numOfpage=len(driver.find_elements(By.XPATH,'/html/body/div/div[2]/div/nav/ul/li'))

print("Open New Workbook")
wb = Workbook()
ws=wb.active
ws.title="Work-items"

print("Write Data In Workbook")
for r in range(1,numOfHrows+1):
    ls=[]
    for c in range(2,numOfHcolumn+1):
        data = driver.find_element(By.XPATH,f"/html/body/div/div[2]/div/table/tbody/tr[{str(r)}]/th[{str(c)}]").text
        ls.append(data)
    ws.append(ls)
for i in range(1,numOfpage+1):
    for r in range(2,numOfrows+1):
        ls=[]
        for c in range(2,numOfcolumn+1):
            data = driver.find_element(By.XPATH,f"/html/body/div/div[2]/div/table/tbody/tr[{str(r)}]/td[{str(c)}]").text
            ls.append(data)
        ws.append(ls)
    driver.find_element(By.XPATH,f'/html/body/div/div[2]/div/nav/ul/li[{str(i)}]/a').click()
print("Save File In Current Location")
wb.save("ACME System.xlsx")