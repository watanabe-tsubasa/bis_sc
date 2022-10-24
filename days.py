from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import sys
import os
import pandas as pd

driver_path = r'C:\Users\341137\src\driver\chromedriver.exe'
access_URL = r'https://sts.aeonpeople.biz/adfs/ls/?SAMLRequest=hZJdT4MwGIX%2FCun9KLAJrhkkuF24ZCoR9MIbU9iLNCkt9i1O%2FfWy4cf0Yl739DmnT7pA3sqOpb1t1C0894DWeW2lQnY4iElvFNMcBTLFW0BmK5anVxsWuB7rjLa60pI4KSIYK7RaaoV9CyYH8yIquLvdxKSxtkNGKYfZrm5mnstBqw50J8EtxTvNG1GWWoJtXERN9%2FiAZjd5QZzVsEcovif%2FcNDiXwTf1kglUuKsVzF5rOo5RDMv8qchhHNeRn5ZRXU9Db3zaL6NzoYYYg9rhZYrG5PAC4KJ700Cr%2FBCFkRsGj4QJ%2Ft83oVQW6GeTrsoxxCyy6LIJuP6ezB4WD4ESLLYG2WHYnPk%2BDSWf4klyX8a8Vvjgh5Vjb0dux7Y61WmpajenFRKvVsa4BZi4hOajFd%2Bf4TkAw%3D%3D&RelayState=ss%3Amem%3Adc1cb8c90983406470be00896527695d630f88d3975670e8aebacf21a43f002b'
auth_id = '341137'
auth_pass = '3682tw'

driver = webdriver.Chrome(driver_path)
driver.implicity_wait(10)
driver.get(access_URL)

title = driver.title
print(title)

# ログイン
driver.find_element(By.CSS_SELECTOR, '#userNameInput').send_keys(auth_id)
driver.find_element(By.CSS_SELECTOR, '#passwordInput').send_keys(auth_pass)
driver.find_element(By.CSS_SELECTOR, '#submitButton').click()

# frameの定義
left_frame = driver.find_element(By.CSS_SELECTOR, 'html > frameset > frameset > frame:nth-child(1)')
right_frame = driver.find_element(By.CSS_SELECTOR, 'html > frameset > frameset > frame:nth-child(2)')

driver.switch_to.frame(left_frame)

# 日別売上情報へのアクセス
a_tags = driver.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
a_tags[4].click()

driver.switch_to.default_content()
driver.switch_to.frame(right_frame)

print('switch_right')

driver.find_elements(By.CSS_SELECTOR, '#kaishaCodeRadio')[2].click()

selector = driver.find_element(By.CSS_SELECTOR, '#displayPatternSelect1')
select = Select(selector)
print('selected')
select.select_by_value('7')

driver.find_elements(By.CSS_SELECTOR, '#periodTypeRadio')[2].click()

#orgLeft

# orgLeft > option:nth-child(2) 北関東 co_1
# orgLeft > option:nth-child(3) 南関東 co_2
# orgLeft > option:nth-child(4) 北陸信越 co_3
# orgLeft > option:nth-child(5) 東海 co_4
# orgLeft > option:nth-child(7) 近畿 6は旧近畿 co_5
# orgLeft > option:nth-child(8) 中四国 co_6
# selectAdd 追加

company_select = driver.find_element(By.CSS_SELECTOR, '#orgLeft')
select_company = Select(company_select)

select_company.select_by_index(1)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()
select_company.select_by_index(1)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()
select_company.select_by_index(1)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()

month_from_date = driver.find_element(By.CSS_SELECTOR, '#monthFromDate')
print('from_date_selected')
Select(month_from_date).select_by_value('202103')
month_to_date = driver.find_element(By.CSS_SELECTOR, '#monthToDate')
print('to_date_selected')
Select(month_to_date).select_by_value('202103')

driver.find_element(By.CSS_SELECTOR, '#excelDownload').click()

driver.quit()