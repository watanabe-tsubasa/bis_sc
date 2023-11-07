from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import time

class BisSelenium(webdriver.Chrome):

  def __init__(
    self,
    driver_path,
    headless=False,
    access_URL=r'https://ae4wfh40.aeonpeople.biz/ar-aeon-hana-spring/auth',
    ):
    self.driver_path = driver_path

    self.headless = headless
    self.access_URL = access_URL
    
    self.options = Options()
    if self.headless:
      self.options.add_argument('--headless')
    service = Service(executable_path=self.driver_path)
    super().__init__(service=service, options=self.options)
    self.implicitly_wait(20)
    self.wait = WebDriverWait(self, 60)
    self.original_window = self.current_window_handle
   
  def login(self, auth_id, auth_pass):
    """login
    """
    self.auth_id = auth_id
    self.auth_pass = auth_pass
    self.get(self.access_URL)
    self.find_element(By.CSS_SELECTOR, '#userNameInput').send_keys(self.auth_id)
    self.find_element(By.CSS_SELECTOR, '#passwordInput').send_keys(self.auth_pass)
    self.find_element(By.CSS_SELECTOR, '#submitButton').click()

  def switch_left(self):
    """サイドバーに遷移
    """
    self.switch_to.default_content()
    left_frame = self.find_element(By.CSS_SELECTOR, 'html > frameset > frameset > frame:nth-child(1)')
    self.switch_to.frame(left_frame)

  def switch_right(self):
    """検索窓に遷移
    """
    self.switch_to.default_content()
    right_frame = self.find_element(By.CSS_SELECTOR, 'html > frameset > frameset > frame:nth-child(2)')
    self.switch_to.frame(right_frame)
    
  def move_to_daysales(self):
    """日別売上速報（一覧表示）
    """
    self.switch_left()
    a_tags = self.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
    a_tags[4].click()
    self.switch_right()
    
  def move_to_timesales(self):
    """時間帯別売上速報（一覧表示）
    """
    self.switch_left()
    a_tags = self.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
    a_tags[5].click()
    self.switch_right()

  def move_to_confirmedsales(self):
    """確定営業情報
    """
    self.switch_left()
    a_tags = self.find_elements(By.CSS_SELECTOR, '#liLink > div > a')
    a_tags[8].click()
    self.switch_right()
    
  def select_orgs(self, *args: *[str]):
    """組織を選択する（valueで指定）
    """
    org_reset_button = self.find_elements(By.CSS_SELECTOR, '#search_div_align > a')[0]
    org_reset_button.click()
    time.sleep(0.2)
    for arg in args:
      print(arg)
      self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#orgLeft')))
      org_selecter = self.find_element(By.CSS_SELECTOR, '#orgLeft')
      org_add_button = self.find_elements(By.CSS_SELECTOR, '#selectAdd')[0]
      Select(org_selecter).select_by_value(arg)
      org_add_button.click()

  def select_products(self, *args: *[str]):
    """商品分類を選択する（innerTextで指定）
    """
    product_reset_button = self.find_elements(By.CSS_SELECTOR, '#search_div_align > a')[1]
    product_reset_button.click()
    time.sleep(0.2)
    for arg in args:
      self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '#productLeft')))
      product_selecter = self.find_element(By.CSS_SELECTOR, '#productLeft')
      product_add_button = self.find_elements(By.CSS_SELECTOR, '#selectAdd')[1]
      Select(product_selecter).select_by_visible_text(arg)
      product_add_button.click()   
    
  def move_to_new_window(self):
    """新しく表示されたwindowをactiveにする
    """
    self.wait.until(EC.number_of_windows_to_be(2))
    for window_handle in self.window_handles:
        if window_handle != self.original_window:
            self.switch_to.window(window_handle)
            print(f'switch_window:{self.title}')
            break
    
  def move_to_origin(self):
    """新しいwindowを閉じて元のwindowをactiveにする
    """
    self.close()
    self.switch_to.window(self.original_window)
    
  def input_without_JS(self, selecter: WebElement, text: str):
    """JSが付随するinput要素を空にして入力

    Args:
        selecter (WebElement): input要素
        text (str): 入力文字列
    """
    self.execute_script('arguments[0].focus();', selecter)
    self.execute_script('arguments[0].value = "";', selecter)
    selecter.send_keys(text)
    
if __name__ == '__main__':

  import time
  import os
  from dotenv import load_dotenv; load_dotenv()
  
  driver_path = r'C:\Users\341137\src\driver\chromedriver.exe'
  auth_id = os.environ.get('AUTH_ID')
  auth_pass = os.environ.get('AUTH_PASS')
  
  bis_selenium = BisSelenium(driver_path)
  bis_selenium.login(auth_id, auth_pass)
  bis_selenium.move_to_timesales()
  
  # element_selecter
  pattern_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#displayPatternSelect1')
  date_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#date')
  hour_selecter = bis_selenium.find_element(By.CSS_SELECTOR, '#hour')  
  report_viewer = bis_selenium.find_element(By.CSS_SELECTOR, '#viewJikanichiran')

  # 条件選択部分
  bis_selenium.find_elements(By.CSS_SELECTOR, '#kaishaCodeRadio')[2].click()
  Select(pattern_selecter).select_by_value('7')
  
  # 組織選択
  org_list = ['0003_12000', '0004_13000']
  bis_selenium.select_orgs(*org_list)
  
  # 商品分類選択
  product_list = ['[211]グロサリー', '[214]リカー', '[221]デイリーフーズ']
  bis_selenium.select_products(*product_list)
  
  # 日付選択
  bis_selenium.input_without_JS(date_selecter, '2023/07/30')
  Select(hour_selecter).select_by_value('10')
  
  # 画面遷移
  report_viewer.click()
  bis_selenium.move_to_new_window()
  first_frame = bis_selenium.find_element(By.CSS_SELECTOR, '#Dottom')
  bis_selenium.switch_to.frame(first_frame)
  
  # first_fetch
  time_selector = bis_selenium.find_element(By.CSS_SELECTOR, '#TIME_SELECT')
  bottom_frame = bis_selenium.find_element(By.CSS_SELECTOR, '#bottomFrame')
  bis_selenium.switch_to.frame(bottom_frame)
  print(bis_selenium.page_source.encode('utf-8'))
  
  # select_date
  bis_selenium.switch_to.default_content()
  bis_selenium.switch_to.frame(first_frame)
  Select(time_selector).select_by_value('13')
  bis_selenium.switch_to.frame(bottom_frame)
  print(bis_selenium.page_source.encode('utf-8'))
  
  # second_fetch
  bis_selenium.move_to_origin()
  bis_selenium.switch_right()
  org_list = ['0005_14000', '0007_16000']
  bis_selenium.select_orgs(*org_list)
  bis_selenium.input_without_JS(date_selecter, '2023/07/29')
  Select(hour_selecter).select_by_value('10') 
  report_viewer.click()
  bis_selenium.move_to_new_window()
  
  bis_selenium.move_to_origin()
  bis_selenium.switch_right()

  # ここまで
  
  time.sleep(4)
  bis_selenium.quit()