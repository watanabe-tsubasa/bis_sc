from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
# import time
# import sys
# import os
import pandas as pd
import openpyxl

options = Options()
options.add_argument('--headless')
driver_path = r'C:\Users\341137\src\driver\chromedriver.exe'
access_URL = r'https://ae4wfh40.aeonpeople.biz/ar-aeon-hana-spring/auth'
auth_id = '341137'
auth_pass = '3682tw'

service = Service(executable_path=driver_path)
driver = webdriver.Chrome(service=service, options=options)
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

# 日別売上速報へのアクセス
driver.switch_to.frame(left_frame)
a_tags = driver.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
a_tags[4].click()

# 検索条件の設定
driver.switch_to.default_content()
driver.switch_to.frame(right_frame)
print('switch_right')

driver.find_elements(By.CSS_SELECTOR, '#kaishaCodeRadio')[2].click()

selector = driver.find_element(By.CSS_SELECTOR, '#displayPatternSelect1')
select = Select(selector)
print('selected')
# select.select_by_value('7') # 店別グループ別
select.select_by_value('L3') # カンパニー事業部店別ライン別

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
select_company.select_by_index(4)
driver.find_element(By.CSS_SELECTOR, '#selectAdd').click()
driver.find_elements(By.CSS_SELECTOR, '#periodTypeRadio')[2].click()


# 取得範囲の設定
list_17 = list(range(201703, 201713)) + list(range(201801, 201803))
list_18 = list(range(201803, 201813)) + list(range(201901, 201903))
list_19 = list(range(201903, 201913)) + list(range(202001, 202003))
list_20 = list(range(202003, 202013)) + list(range(202101, 202103))
list_21 = list(range(202103, 202113)) + list(range(202201, 202203))
list_22 = list(range(202203, 202213)) + list(range(202301, 202203))

years = {
    2017: list_17,
    2018: list_18,
    2019: list_19,
    2020: list_20,
    2021: list_21,
    2022: list_22
}

# years = {2022: [202301, 202302]} # デバッグ用
first_time = True # データ取得1回目かどうかの判定に用いる

# この部分から下をループして複数日付を取得する ===================================
for year, month_list in years.items():
    add_book = openpyxl.Workbook()
    # add_book_path = f'../data/{year}.xlsx'
    add_book_path = f'../data/{year}_line.xlsx'
    add_book.save(add_book_path)

    for month in month_list:
        # 取得する日付の設定
        select_date_from = driver.find_element(By.CSS_SELECTOR, '#monthFromDate')
        select_date_to = driver.find_element(By.CSS_SELECTOR, '#monthToDate')
        print('date_selected')
        # 1回目は初期値が最新の日付のため、select_date_fromから更新する。
        if first_time:
            Select(select_date_from).select_by_value(str(month))
            # time.sleep(1)
            Select(select_date_to).select_by_value(str(month))
            # time.sleep(1)
        # 2回目以降はselect_date_toから更新しないと日付が逆順となるエラーが発生する。
        else:
            Select(select_date_to).select_by_value(str(month))
            # time.sleep(1)
            Select(select_date_from).select_by_value(str(month))
            # time.sleep(1)            
        print('span_selected')
        driver.find_element(By.CSS_SELECTOR, '#viewNichibetsuichiran').click()

        # window切り替え
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            print(driver.title)
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                print(f'switch_window:{driver.title}')
                break

        driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, '#Dottom'))
        
        # 表示を全表示に切り替え      
        display_switch = driver.find_element(By.CSS_SELECTOR, '#displaySwith')
        Select(display_switch).select_by_value('0')
        
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
        table.columns = (table.columns
                         .str.replace('\n', '')
                         .str.replace(' ', '')
                         .str.replace('　', '')
                         )
        # print(table.columns)
        table = table[['店舗CD', 
                       '店舗', 
                    #    'グループCD', 
                    #    'グループ', 
                       '4ラインCD', 
                       '4ライン', 
                       '累計売上高', 
                       '累計日割差', 
                       '昨年累計売上高同規模同日', 
                       '昨年累計売上高同規模同曜', 
                       '累計客数', 
                       '昨年累計客数同規模同日', 
                       '昨年累計客数同規模同曜', 
                       '累計点数',
                       '昨年累計点数同規模同日',
                       '昨年累計点数同規模同曜',
                       ]]
        # info_cols = ['店舗CD','店舗', 'グループCD', 'グループ'] # グループ別
        info_cols = ['店舗CD','店舗', '4ラインCD', '4ライン'] # ライン別
        table_info = table[info_cols]
        table_values = table.drop(columns=info_cols)
        val_cols = table_values.columns
        for col in val_cols:
            # print(table_values[col])
            table_values[col] = table_values[col].str.replace(',', '').replace('\xa0', '0').astype(float)
            # print(table_values[col])
        
        table_values['累計日割'] = table_values['累計売上高'] + table_values['累計日割差']
        
        table = pd.concat([table_info, table_values], axis=1)
        table = table[['店舗CD', 
                       '店舗', 
                    #    'グループCD', 
                    #    'グループ', 
                       '4ラインCD', 
                       '4ライン', 
                       '累計売上高', 
                       '累計日割', 
                       '昨年累計売上高同規模同日', 
                       '昨年累計売上高同規模同曜', 
                       '累計客数', 
                       '昨年累計客数同規模同日', 
                       '昨年累計客数同規模同曜', 
                       '累計点数',
                       '昨年累計点数同規模同日',
                       '昨年累計点数同規模同曜',
                       ]]        
  
        with pd.ExcelWriter(add_book_path, mode = 'a') as writer:
            table.to_excel(writer, sheet_name=str(month), index=False)
        
        print(month)
        
        driver.close()
        driver.switch_to.window(original_window)
        print(driver.title)
        driver.switch_to.frame(right_frame)
        print('switch_right')
        
        first_time = False
        
    wb = openpyxl.load_workbook(add_book_path)
    wb.remove(wb.worksheets[0])
    wb.save(add_book_path)

# time.sleep(2)

driver.quit()

# os.rename('./data/result.xlsx', './data/data.xlsx')