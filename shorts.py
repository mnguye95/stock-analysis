
import requests
from datetime import date
from bs4 import BeautifulSoup
import pandas

def get_shorts():

    url = "https://www.highshortinterest.com/all/{}"

    rows_found = False

    tickers = []
    pages = 2

    for i in range(1, pages+1):
        while not rows_found:
            try:
                soup = BeautifulSoup(requests.get(url.format(i)).text, 'lxml')
                table = soup.find("table",{"class" : "stocks"})
                rows = table.find_all('tr')
                for row in rows[1:]:
                    row_obj = {}
                    cell_count = 0
                    cols = row.find_all('td')
                    if len(cols) > 1:
                        for cell in cols:
                            datum = cell
                            if cell_count == 0:
                                row_obj['Symbol'] = datum.find('a').get_text()
                            elif cell_count == 1:
                                row_obj['Company'] = datum.get_text()
                            elif cell_count == 2:
                                row_obj['Exchange'] = datum.get_text()
                            elif cell_count == 3:
                                row_obj['ShortInterest'] = datum.get_text()
                            elif cell_count == 4:
                                row_obj['Float'] = datum.get_text()
                            elif cell_count == 5:
                                row_obj['Outstanding'] = datum.get_text()
                            elif cell_count == 6:
                                row_obj['Industry'] = datum.get_text()
                            cell_count = cell_count + 1
                        tickers.append(row_obj)
                rows_found = True
            except Exception as e:
                print(e)
        rows_found = False

    return pandas.DataFrame(tickers)