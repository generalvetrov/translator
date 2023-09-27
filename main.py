import time
import openpyxl
from selenium import webdriver
#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By


headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    # 'Cookie': 'LKQEshop=Login=False&CatalogSwitchViewType=Row; _fbp=fb.1.1694899067962.827195960; _ga_WB4E636C76=GS1.1.1694947436.4.1.1694947442.0.0.0; OptanonConsent=isGpcEnabled=0&datestamp=Sun+Sep+17+2023+13%3A43%3A57+GMT%2B0300+(%D0%B7%D0%B0+%D1%81%D1%85%D1%96%D0%B4%D0%BD%D0%BE%D1%94%D0%B2%D1%80%D0%BE%D0%BF%D0%B5%D0%B9%D1%81%D1%8C%D0%BA%D0%B8%D0%BC+%D0%BB%D1%96%D1%82%D0%BD%D1%96%D0%BC+%D1%87%D0%B0%D1%81%D0%BE%D0%BC)&version=202210.1.0&isIABGlobal=false&hosts=&genVendors=&consentId=1e191346-f3f2-42bc-9cc6-bf34e9f65a0c&interactionCount=2&landingPath=NotLandingPage&groups=C0001%3A1%2CC0002%3A1%2CC0004%3A1&geolocation=UA%3B30&AwaitingReconsent=false; _ga=GA1.1.1444528616.1694788081; _gat_UA-170320319-1=1; _gcl_au=1.1.744428976.1694899068; _gid=GA1.2.1857193055.1694899068; __RequestVerificationToken=yFhw5r96iB8XqAaB_T3ecRLxDws9HBbI-SA2S7WNRMyb4yZumArNh2G2F2R-L_PxKudsZSMDBBu7oSangP3IiAgXbDL3FUwhn1S1sFooVIE1; OptanonAlertBoxClosed=2023-09-16T21:17:47.611Z; ASP.NET_SessionId=gk15q4iwfwapokgph511gelt',
    'Referer': 'https://www.lkq.cz/homepage/car',
    'Cache-Control': 'max-age=0',
    'Host': 'www.lkq.cz',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.3 Safari/605.1.15',
    'Accept-Language': 'uk-UA,uk;q=0.9',
    # 'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
}

FILE_NAME = "StarLine.xlsx"                                                          # enter the file name and sheet names
#MAIN_SHEET_NAME = 'New Assortiment'
MAIN_SHEET_NAME = 'Sheet1'

workbook = openpyxl.open(FILE_NAME)
main_sheet = workbook[MAIN_SHEET_NAME]


def registration(link):
    options = webdriver.ChromeOptions()
    for key, value in headers.items():
        options.add_argument(f'--header={key}:{value}')

    browser = webdriver.Chrome(options=options)

    browser.get(link)
    browser.implicitly_wait(1)

    WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#onetrust-accept-btn-handler'))).click()

    languages = browser.find_element(By.CSS_SELECTOR, '#lang_nav')
    languages.click()
    ua = browser.find_elements(By.CSS_SELECTOR, '.dropdown-language')[1]
    ua.click()

    excel_go(browser)

    workbook.save(FILE_NAME)
    workbook.close()


def excel_go(browser):
    for (art_num, brand, name_ua, _, _, _, name_pl, count, id) in main_sheet.iter_rows(min_row=2):
        name_ua.value = find_name(browser, art_num.value)


def find_name(browser, art_num):
    inputClientNumber = browser.find_element(By.CSS_SELECTOR, '#SearchFocus')
    inputClientNumber.send_keys(art_num)

    time.sleep(1)     #work

    try:
        result = browser.find_element(By.CLASS_NAME, 'ProductName')
        print(result.text)

        #time.sleep(1)

        inputClientNumber.clear()
        #time.sleep(1)

        return result.text
    except:
        inputClientNumber.clear()
        return 'no result'


link = 'https://www.lkq.cz/homepage/car'

registration(link)

