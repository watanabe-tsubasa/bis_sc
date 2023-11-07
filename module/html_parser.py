import lxml
from bs4 import BeautifulSoup

def html_parser(html: bytes):
  soup = BeautifulSoup(html, 'lxml')
  table = soup.find('tbody')
  header_base = table.find('tr').find_all('td')
  headers = [header.get_text().replace('\n', '').replace(' ', '') for header in header_base]
  rows_base = table.find_all('tr')[1:]
  body = []
  for row_base in rows_base:
    row = []
    for row_elem in row_base.find_all('td'):
      row.append(row_elem.get_text().replace('\n', '').replace(' ', '') )
    body.append(row)
    
  return (headers, body)
  