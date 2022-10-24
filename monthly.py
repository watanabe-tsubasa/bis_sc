from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import sys
import os
import pandas as pd
import openpyxl

driver_path = r'C:\Users\341137\src\driver\chromedriver.exe'
access_URL = r'https://ae4wfh40.aeonpeople.biz/ar-aeon-hana-spring/auth'
auth_id = '341137'
auth_pass = '3682tw'

driver = webdriver.Chrome(driver_path)
driver.implicitly_wait(20)
wait = WebDriverWait(driver, 60)
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
original_window = driver.current_window_handle

# 確定営業情報へのアクセス
driver.switch_to.frame(left_frame)
a_tags = driver.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
a_tags[8].click()

# 検索条件の設定
driver.switch_to.default_content()
driver.switch_to.frame(right_frame)
print('switch_right')

driver.find_elements(By.CSS_SELECTOR, '#kaishaCodeRadio')[2].click()

selector = driver.find_element(By.CSS_SELECTOR, '#displayPatternSelect1')
select = Select(selector)
print('selected')
select.select_by_value('7')

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

select_company.select_by_index(7)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()
select_company.select_by_index(6)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()
select_company.select_by_index(4)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()

# 取得範囲の設定
list_17 = list(range(201703, 201713)) + list(range(201801, 201803))
list_18 = list(range(201803, 201813)) + list(range(201901, 201903))
list_19 = list(range(201903, 201913)) + list(range(202001, 202003))
list_20 = list(range(202003, 202013)) + list(range(202101, 202103))
list_21 = list(range(202103, 202113)) + list(range(202201, 202203))
list_22 = list(range(202203, 202210)) # + list(range(202201, 202203))

years = {
    2017: list_17,
    2018: list_18,
    2019: list_19,
    2020: list_20,
    2021: list_21,
    2022: list_22
}

# この部分から下をループして複数日付を取得する ===================================
for year, month_list in years.items():
    add_book = openpyxl.Workbook()
    add_book.save(f'./data/{year}.xlsx')

    for month in month_list:
        # 取得する日付の設定
        select_date = driver.find_element(By.CSS_SELECTOR, '#knk_to_ym')
        print('date_selected')
        Select(select_date).select_by_value(str(month))
        print('span_selected')
        driver.find_element(By.CSS_SELECTOR, '#viewKakuteiichiran').click()

        # window切り替え
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            print(driver.title)
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                print(f'switch_window:{driver.title}')
                break

        driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, '#Dottom'))
        driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, '#bottomFrame'))
        print('switch_frame')

        print('waiting_tbody')
        wait.until(EC.presence_of_element_located((By.TAG_NAME, 'tbody')))

        print('getting_html')
        html = driver.page_source.encode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        table_html = soup.select_one('body > table')
        tbody_html = soup.select_one('body > table > tbody')

        print('success')

        tr_list = tbody_html.find_all('tr')
        columns = []
        for i in range(len(tr_list[0].find_all('td'))):
            columns.append(tr_list[0].find_all('td')[i].get_text())

        # table = pd.DataFrame()
        # table['columns'] = columns

        row_list = []
        for j in range(1, len(tr_list)):
            row_data = []
            for k in range(len(tr_list[j].find_all('td'))):
                row_data.append(tr_list[j].find_all('td')[k].get_text().replace('\n', ''))
            # append_row = pd.DataFrame(row_data)
            row_list.append(row_data)

        table = pd.DataFrame(row_list, columns = columns)
        table.columns = table.columns.str.replace('\n', '').str.strip()
        
        with pd.ExcelWriter(f'./data/{year}.xlsx', mode = 'a') as writer:
            table.to_excel(writer, sheet_name=str(month))
        
        print(month)
        
        driver.close()
        driver.switch_to.window(original_window)
        print(driver.title)
        driver.switch_to.frame(right_frame)
        print('switch_right')

time.sleep(2)

driver.quit()

# os.rename('./data/result.xlsx', './data/data.xlsx')