import os
import openpyxl
import glob

dir_path = r'C:\Users\341137\OneDrive - AEON\デスクトップ\月度別店別営業利益'

rename_files = glob.iglob(dir_path + '\*.xlsx')
print(rename_files)

for file_path in rename_files:
    # print(file_path) それぞれのpathからA3（日付記載）の値を取得
    temp_book = openpyxl.load_workbook(file_path)
    sheet = temp_book['Sheet1']
    date = sheet['A3'].value
    # print(f'{dir_path}\{date}.xlsx')
    os.rename(file_path, f'{dir_path}\{date}.xlsx') # ファイル名を取得した日付の値に変更
    
print('success')