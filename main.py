# Required Modules
import datetime, time, wget, os, errno, ctypes
from datetime import date
from sys import exit
from urllib.request import urlopen  # the lib that handles the url stuff
import random
import requests
import pickle
import json
import pandas
from bs4 import BeautifulSoup
from tabulate import tabulate
import imgkit
from pretty_html_table import build_table
import traceback

# Selenium for web automation and data extraction
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options  
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException

# Email module
import email, smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

path_wkthmltoimage = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe'
config = imgkit.config(wkhtmltoimage=path_wkthmltoimage)

def get_shorts():

    url = "https://www.highshortinterest.com/all/{}"

    rows_found = False

    shorts_csv = []
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
                            # row_string = row_string + ' ' + datum
                            cell_count = cell_count + 1
                        # print(row_string)
                        shorts_csv.append(row_obj)
                rows_found = True
            except Exception as e:
                print(e)
        rows_found = False

    return pandas.DataFrame(shorts_csv)
    # make_excel(shorts_csv)

def get_proxy():
    print('Loading Proxies...')

    proxy_list_url = "https://raw.githubusercontent.com/clarketm/proxy-list/master/proxy-list-raw.txt"

    def format_ip(n):
        return n.decode("utf-8").replace('\n', '')

    ip_addresses = urlopen(proxy_list_url) # it's a file like object and works just like a file
    ip_addresses = list(map(format_ip, ip_addresses))
    proxy_index = 0
    working_proxy = ''

    found_proxy = False
    print('Locating Working Proxy...')
    while not found_proxy:
        try:
            proxy_index = random.randint(0, len(ip_addresses) - 1)
            current_ip = ip_addresses[proxy_index]
            proxies = {"http": "http://{}".format(current_ip), "https": "https://{}".format(current_ip)}
            print('Attempting {}'.format(current_ip))
            r = requests.get("https://www.google.com/", proxies=proxies, timeout=3)
            found_proxy = True
            working_proxy = current_ip
            print('Proxy Located... {}'.format(current_ip))
        except Exception as e:
            broken_ip = ip_addresses.pop(proxy_index)
    
    return working_proxy

def is_third_friday(s):
    return s.weekday() == 4 and 15 <= s.day <= 21

def make_link(x,y): 
    g = datetime.datetime.strptime(y, '%m/%d/%y')
    d = g.strftime('%Y-%m-%d')
    o = 'm' if is_third_friday(g) else 'w'
    return '<a href="https://www.barchart.com/stocks/quotes/{}/options?expiration={}-{}" target="_blank">{}</a>'.format(x,d,o,x)

def send_email(file_data, shorts_data, file_path):

    today_date = date.today().strftime("%b %d %Y")
    
    subject = "Options Report {}".format(today_date)
    sender_email = "uoa.alerts@gmail.com"
    password = ""

    # Creating Body of
    text = """
    Stock Nitro Report {date}

    Most Volume
    {volume}

    Best Calls
    {calls}

    Best Puts
    {puts}

    Best Calls For Short Tickers
    {shorts}

    Good luck!
    
    """

    html = """
    <html><head>
    <style>
      h3, h4 {{text-align: center;}}
      html, body {{ height: 100%; margin: 0; padding: 0; }}
      .main {{display: table; width: 100%; height: 100%;}}
      .row {{display: table-row; height: 100%;}}
      .inner-table {{display: table-cell; padding: 7px;}}
      table {{ margin: 0 auto; border-collapse: collapse; text-align: left; overflow: hidden; }}
      td, th {{ border-top: 1px solid #ECF0F1; padding: 5px; }}
      td {{ border-left: 1px solid #ECF0F1; border-right: 1px solid #ECF0F1; }}
      th {{ background-color: #4ECDC4; }}
      tr:nth-of-type(even) td {{ background-color: lighten(#4ECDC4, 35%); }}
    </style>
    </head>
    <body>
    <h4>Stock Nitro Report {date}</h4>
    <div class="main">
    <div class="row">
    <div class="inner-table">
    <h3>Most Volume</h3>
    {volume}
    </div>
    <div class="inner-table">
    <h3>Best Calls</h3>
    {calls}
    </div>
    <div class="inner-table">
    <h3>Best Puts</h3>
    {puts}
    </div>
    <div class="inner-table">
    <h3>Best Calls For Short Tickers</h3>
    {shorts}
    </div>
    </div>
    </div>
    <h4>The Put-call ratio (PCR) is to say that a higher ratio means it's time to sell and a lower ratio means it's time to buy, because when the ratio is high it suggests that people are either expecting or protecting more readily against a future decline in the price of the underlying. A Put-Call ratio between 0.5 and 1 is considered a sideways trend in the markets.</h4>
    <h4>Good luck!</h4>
    </body></html>
    """

    file_data['Symbol'] = file_data.apply(lambda x: make_link(x['Symbol'], x['Exp Date']), axis=1)
    shorts_data['Symbol'] = shorts_data.apply(lambda x: make_link(x['Symbol'], x['Exp Date']), axis=1)

    # file_data.set_index('Symbol', inplace=True)
    # shorts_data.set_index('Symbol', inplace=True)
    file_data = file_data.rename_axis(None, axis=1)
    shorts_data = shorts_data.rename_axis(None, axis=1)
    file_data.reset_index()
    shorts_data.reset_index()
    print(file_data.keys())

    buys = file_data.sort_values(by='PCR', ascending=True)
    sells = file_data.sort_values(by='PCR', ascending=False)
    shorts_data = shorts_data.sort_values(by='PCR', ascending=True)

    # text = text.format(volume=tabulate(file_data.head(30), headers='keys', tablefmt="grid"), calls=tabulate(buys.head(30), headers='keys', tablefmt="grid"),
    #     puts=tabulate(sells.head(30), headers='keys', tablefmt="grid"), shorts=tabulate(shorts_data.head(30), headers='keys', tablefmt="grid"), date=today_date, showindex=False)
    # html = html.format(volume=tabulate(file_data.head(30), headers='keys', tablefmt="html"), calls=tabulate(buys.head(30), headers='keys', tablefmt="html"),
    #     puts=tabulate(sells.head(30), headers='keys', tablefmt="html"), shorts=tabulate(shorts_data.head(30), headers='keys', tablefmt="html"), date=today_date, showindex=False)
    text = text.format(volume=tabulate(file_data.head(30), headers='keys', tablefmt="grid"), calls=tabulate(buys.head(30), headers='keys', tablefmt="grid"),
        puts=tabulate(sells.head(30), headers='keys', tablefmt="grid"), shorts=tabulate(shorts_data.head(30), headers='keys', tablefmt="grid"), date=today_date, showindex=False)
    html = html.format(volume=file_data.head(30).to_html(escape=False, index=False), calls=buys.head(30).to_html(escape=False, index=False),
        puts=sells.head(30).to_html(escape=False, index=False), shorts=shorts_data.head(30).to_html(escape=False, index=False), date=today_date, showindex=False)
    
    file_name_format = date.today().strftime("%b_%d_%Y")
    imgkit.from_string(html, "./screenshots/{}_EOD_Options_Report.jpg".format(file_name_format), config=config)
    
    # Add body to email
    message = MIMEMultipart("alternative", None, [MIMEText(text), MIMEText(html,'html')])

    # Create a multipart message and set headers
    message["From"] = "Stock Nitro"
    message["Subject"] = subject
    # message["Bcc"] = receiver_email  # Recommended for mass emails

    filename = file_path  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {file_name_format}_EOD_Options_Report.xlsx",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        recipients = ['mnguye95@gmail.com']
        for recipient in recipients:
            print('Sending to {}'.format(recipient))
            message["To"] = recipient
            server.sendmail(sender_email, recipient, text)

def make_excel(data, done=False):

    main_df = None

    if not done:
        main_df = pandas.DataFrame(data)
    else:
        main_df= data
        main_df = main_df.parse("main")

    all_volume = main_df.groupby(['Symbol', 'Type'])['Volume'].sum().reset_index()
    by_exp = main_df.groupby(['Symbol', 'Type', 'Exp Date'])['Volume'].sum().reset_index()

    volume_html = all_volume.pivot_table('Volume', ['Symbol'], 'Type')
    exp_html = by_exp.pivot_table('Volume', ['Exp Date','Symbol'], 'Type')

    volume_html = volume_html.reset_index(0)
    exp_html = exp_html.reset_index(0)
    exp_html = exp_html.reset_index(0)

    exp_html['PCR'] = exp_html['Put'] / exp_html['Call']
    volume_html['PCR'] = volume_html['Put'] / volume_html['Call']

    exp_html = exp_html.sort_values(by=['Call', 'PCR'], ascending=[False, True])
    volume_html = volume_html.sort_values(by=['Call', 'PCR'], ascending=[False, True])

    file_name = './results/options_{}.xlsx'.format(date.today().strftime("%b-%d-%Y"))

    writer = pandas.ExcelWriter(file_name, engine='xlsxwriter')

    main_df.to_excel(writer, sheet_name='main', index=False)

    exp_html.to_excel(writer, sheet_name='by_expiration', index=False)

    volume_html.to_excel(writer, sheet_name='by_volume', index=False)

    # details = {
    # 'Symbol' : ['BLNK', 'GME', 'ROOT'],
    # 'Float' : ['23M', '21M', '34M'],
    # 'ShortInterest' : ['43%', '65%', '34%'],
    # }
    
    # _shorts = pandas.DataFrame(details)
    # shorts = _shorts

    shorts = get_shorts()
    shorts = shorts[['Symbol', 'Float', 'ShortInterest']]
    high_shorts = shorts.merge(exp_html, on=['Symbol'], how='inner')
    high_shorts = high_shorts.sort_values(by='PCR', ascending=True)

    high_shorts.to_excel(writer, sheet_name='high_shorts', index=False)
    high_shorts.to_html('high_shorts.html', index=False)

    writer.save()

    ask_email = False

    while not ask_email:
        ans = input('Send emails? (y,n)\n')
        if ans.lower() == 'y' or 'y' in ans:
            try:
                send_email(exp_html, high_shorts, file_name)
                ask_email = True
            except Exception as e:
                traceback.print_exc()
        else:
            ask_email = True
            print('Done')

def start_scrape(proxy):
    url = "https://www.barchart.com/options/unusual-activity/stocks"
    chrome_options = Options()
    chrome_options.add_argument('--proxy-server=http://%s' % proxy)
    chrome_options.add_argument("--start-maximized")
    # chrome_options.add_argument("--headless")
    chrome_options.add_experimental_option("prefs", {
    "download.default_directory": r"C:\Users\Michael\Documents\projects\options\results",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True,
    })
    driver = None

    SHORT_TIMEOUT = 30

    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(e)
        # ctypes.windll.user32.MessageBoxW(None, "Please download Chromedriver and put at the root of this program.", "No Chrome Driver Found.", 0)
        exit()

    driver.get(url)

    rows_found = False
    done = False

    data_csv = []
    row_count = 1

    while not done:
        while not rows_found:
            try:
                print('Waiting for page to load...')
                WebDriverWait(driver, SHORT_TIMEOUT
                        ).until(EC.presence_of_element_located((By.XPATH, '//*[@sly-repeat="row in rows.data"]')))
                page_source = driver.page_source
                soup = BeautifulSoup(page_source, 'lxml')
                print('Page loaded')
                sym_list = soup.find_all(attrs={"sly-repeat" : "row in rows.data"})
                try:
                    for row in sym_list:
                        # row_string = ''
                        cell_count = 0
                        row_obj = {}
                        for cell in row.find_all(attrs={"data-ng-bind" : "cell"}):
                            datum = cell.get_text()
                            if cell_count == 0:
                                row_obj['Symbol'] = datum
                            elif cell_count == 1:
                                row_obj['Price'] = float(datum.replace(',', '')) if datum != 'N/A' else 'N/A'
                            elif cell_count == 2:
                                row_obj['Type'] = datum
                            elif cell_count == 3:
                                row_obj['Strike'] = datum
                            elif cell_count == 4:
                                row_obj['Exp Date'] = datum
                            elif cell_count == 5:
                                row_obj['Day(s) to Expire'] = datum
                            elif cell_count == 6:
                                row_obj['Bid'] = float(datum)
                            elif cell_count == 7:
                                row_obj['Midpoint'] = float(datum)
                            elif cell_count == 8:
                                row_obj['Ask'] = float(datum)
                            elif cell_count == 9:
                                row_obj['Last'] = float(datum)
                            elif cell_count == 10:
                                row_obj['Volume'] = int(float(datum.replace(',', '')))
                            elif cell_count == 11:
                                row_obj['Open Interest'] = int(float(datum.replace(',', '')))
                            elif cell_count == 12:
                                row_obj['Vol/OI'] = float(datum)
                            elif cell_count == 13:
                                row_obj['Implied Volitity'] = datum
                            elif cell_count == 14:
                                row_obj['Order Date'] = datum
                            # row_string = row_string + ' ' + datum
                            cell_count = cell_count + 1
                        # print(row_string)
                        data_csv.append(row_obj)
                    rows_found = True
                except Exception as e:
                    print(e)
            except:
                driver.refresh()
                print('Could not find rows.')
        
        print("Page {} Processed".format(row_count))
        row_count = row_count + 1
        
        is_next_page = False
        while not is_next_page:
            try:
                cookie_accept = None
                try:
                    cookie_accept = driver.find_element_by_xpath("//button[contains(text(), 'Accept all')]")
                except:
                    print()
                if cookie_accept:
                    cookie_accept.click()
                else:
                    next_button = driver.find_element_by_class_name("next")
                    
                    if next_button.is_displayed():
                        next_button.click()
                        rows_found = False
                    else:
                        print('Last page .')
                        done = True

                    is_next_page = True
            except Exception as e:
                print(e)

    if done:
        make_excel(data_csv)

    return True

scrape_finish = False

while not scrape_finish:
    pickle_exist = False
    try:
        pickle.load(open("proxy.pickle", "rb"))
        pickle_exist = True
    except:
        pickle_exist = False
    if pickle_exist:
        new_proxy = input('Use cached proxy? (y,n) \n')
        if new_proxy.lower() == 'y' or 'y' in new_proxy:
            try:
                scrape_finish = start_scrape(pickle.load(open("proxy.pickle", "rb")))
            except (OSError, IOError) as e:
                print(e)
        else:
            try:
                cur_proxy = get_proxy()
                pickle.dump(cur_proxy, open("proxy.pickle", "wb"))
                scrape_finish = start_scrape(cur_proxy)
            except:
                print('Pickle existed, but a new one was requested and crashed.')
    else:
        try:
            cur_proxy = get_proxy()
            pickle.dump(cur_proxy, open("proxy.pickle", "wb"))
            scrape_finish = start_scrape(cur_proxy)
        except:
            print('Pickle did not existed, so a new one was made but crashed.')
        
        

# send_email('./results/options_Jan-26-2021.xlsx')
# make_excel(pandas.ExcelFile('./results/options_{}.xlsx'.format(date.today().strftime("%b-%d-%Y")), engine='openpyxl'), True)
