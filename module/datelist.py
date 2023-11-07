from datetime import datetime, timedelta

def datelist(start_date_str: str, end_date_str: str):

  # 開始日と終了日を設定
  start_date = datetime.strptime(start_date_str, '%Y/%m/%d')
  end_date = datetime.strptime(end_date_str, '%Y/%m/%d')

  # 日付のリストを作成
  date_list = []
  while start_date <= end_date:
      date_list.append(start_date.strftime('%Y/%m/%d'))  # 必要に応じてフォーマットを変更
      start_date += timedelta(days=1)
      
  return date_list

if __name__ == '__main__':
  date_list = datelist('2022/04/01', '2023/03/31')
  print(date_list)
