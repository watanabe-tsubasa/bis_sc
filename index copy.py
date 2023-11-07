from module.bis_selenium import BisSelenium
from module.html_parser import html_parser
from module.generate_df import generate_df
from module.datelist import datelist

import polars as pl
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

if __name__ == '__main__':

  import time
  import os
  from dotenv import load_dotenv; load_dotenv()
  
  date_list = datelist('2022/04/01', '2023/03/31')
  
  driver_path = r'C:\Users\341137\src\driver\chromedriver.exe'
  auth_id = os.environ.get('AUTH_ID')
  auth_pass = os.environ.get('AUTH_PASS')
  
  bis_selenium = BisSelenium(driver_path, headless=True)
  bis_selenium.login(auth_id, auth_pass)
  bis_selenium.move_to_timesales()
  
  # element_selecter
  pattern_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#displayPatternSelect1')
  hour_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#hour')  
  
  # 条件選択部分
  bis_selenium.find_elements(By.CSS_SELECTOR, '#kaishaCodeRadio')[2].click()
  Select(pattern_selecter).select_by_value('7')
  
  # 組織選択
  north_org = ['0002_11000', '0003_12000', '0004_13000']
  south_org = ['0005_14000', '0007_16000', '0008_17000']
  bis_selenium.select_orgs(*south_org)
  
  # 商品分類選択
  product_list = ['[211]グロサリー', '[214]リカー', '[221]デイリーフーズ']
  bis_selenium.select_products(*product_list)
  
  # 日付選択
  Select(hour_selecter).select_by_value('10')
  for date in date_list:
    try:
      date_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#date')
      bis_selenium.input_without_JS(date_selecter, date)
      # 画面遷移
      report_viewer = bis_selenium.find_element(By.CSS_SELECTOR, '#viewJikanichiran')
      report_viewer.click()
      bis_selenium.move_to_new_window()
      first_frame = bis_selenium.find_element(By.CSS_SELECTOR, '#Dottom')
      bis_selenium.switch_to.frame(first_frame)
      
      # first_fetch
      bottom_frame = bis_selenium.find_element(By.CSS_SELECTOR, '#bottomFrame')
      bis_selenium.switch_to.frame(bottom_frame)
      html = bis_selenium.page_source.encode('utf-8')
      headers, body = html_parser(html)
      df_base, df_selected = generate_df(headers, body)
      df_concat = pl.concat([df_base, df_selected], how='horizontal')
      
      # second_third_fetch
      select_time_list = ['13', '16']
      for select_time in select_time_list:
        bis_selenium.switch_to.default_content()
        bis_selenium.switch_to.frame(first_frame)
        time_selector = bis_selenium.find_element(By.CSS_SELECTOR, '#TIME_SELECT')
        bottom_frame = bis_selenium.find_element(By.CSS_SELECTOR, '#bottomFrame')
        Select(time_selector).select_by_value(select_time)
        bis_selenium.switch_to.frame(bottom_frame)
        html = bis_selenium.page_source.encode('utf-8')
        headers, body = html_parser(html)
        _, df_selected = generate_df(headers, body)
        df_concat = pl.concat([df_concat, df_selected], how='horizontal')
      
      new_date = date.replace('/', '_')
      df_concat.write_csv(f'./south_org/{new_date}.csv')
      
      bis_selenium.move_to_origin()
      bis_selenium.switch_right()
    except:
      continue
  
  # ここまで
  
  time.sleep(4)
  bis_selenium.quit()