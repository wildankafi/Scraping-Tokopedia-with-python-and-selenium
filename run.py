import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import json
import glob
import pandas as pd


def run_url(link):
    dir = os.path.dirname(__file__)
    chrome_driver_path = dir + "\msedgedriver.exe"
    # create a new Chrome session
    driver = webdriver.Edge(chrome_driver_path)
    driver.implicitly_wait(30)
    driver.maximize_window()
    driver.get(link)
    delay = 15  # seconds
    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
    try:
        WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'css-1mswclh')))
        print("Page is ready!")
    except TimeoutException:
        print(
            "Loading took too much time!")

    data = driver.page_source
    driver.quit()
    soup = BeautifulSoup(data, "html.parser")
    # soup = BeautifulSoup(data, "html5lib")
    return soup

def get_url():
    soup = run_url(f"https://www.tokopedia.com/p/komputer-laptop/media-penyimpanan-data/harddisk-external")
    count = soup.find_all(attrs={'class': 'css-k21wea-unf-pagination-item e19tp72t1'})
    pages = None
    for pages in count:pass
    if pages:
        lastpages = pages.get_text()
    totalpages = int(lastpages.replace('.', ''))
    for p in range(1,totalpages):
        soupPages = run_url(f"https://www.tokopedia.com/p/komputer-laptop/media-penyimpanan-data/harddisk-external?page={p}")
        product = soupPages.find_all(attrs={'class': 'css-bk6tzz e1nlzfl3'})
        # urls = []
        data = {}
        data['product'] = []
        no = 1
        for link in product:
            url = link.find('a')['href']
            data['product'].append({
                'no': no,
                'url': url,
                'page': p
            })
            no += 1
            # End Write file json
    with open('data.json', 'w') as outfile:
        json.dump(data, outfile)
    # End Write file json
    print('Get URL link....')

def detail_product(url):
    print('Get Detail Product', url)
    dir = os.path.dirname(__file__)
    chrome_driver_path = dir + "\msedgedriver.exe"
    # create a new Chrome session
    driver = webdriver.Edge(chrome_driver_path)
    driver.implicitly_wait(30)
    driver.maximize_window()
    driver.get(url)
    delay = 15  # seconds
    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
    try:
        WebDriverWait(driver, delay).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'css-1mswclh')))
        print("Page is ready!")
    except TimeoutException:
        print(
            "Loading took too much time!")

    data = driver.page_source
    driver.quit()
    soup = BeautifulSoup(data, "html.parser")
    title = soup.find(attrs={'data-testid': 'lblPDPDetailProductName'}).get_text()
    rating = soup.find(attrs={'data-testid': 'lblPDPDetailProductRatingNumber'}).get_text()
    review = soup.find(attrs={'data-testid': 'lblPDPDetailProductRatingCounter'}).get_text().replace("(", "").replace(")", "")
    sale = soup.find(attrs={'data-testid': 'lblPDPDetailProductSuccessRate'}).get_text()
    view = soup.find(attrs={'data-testid': 'lblPDPDetailProductSeenCounter'}).get_text().split("x", 1)[0]
    price = soup.find(attrs={'data-testid': 'lblPDPDetailProductPrice'}).get_text().replace('Rp','')
    weight = soup.find(attrs={'data-testid': 'PDPDetailWeightValue'}).get_text()
    condition = soup.find(attrs={'data-testid': 'PDPDetailConditionValue'}).get_text()
    desc = soup.find(attrs={'data-testid': 'pdpDescriptionContainer'})
    if desc is None:
        description = "None"
    else:
        description = desc.get_text()
    store = soup.find(attrs={'data-testid': 'llbPDPFooterShopName'}).get_text()
    location = soup.find(attrs={'data-testid': 'lblPDPFooterLastOnline'}).get_text().split("\xa0•\xa0", 1)[0]

    dict_data = {
        'Title': title,
        'Rating': rating,
        'Review': review,
        'Sale': sale,
        'View': view,
        'Price': price,
        'Weight': weight,
        'Condition': condition,
        'Description': description,
        'Store': store,
        'Location': location
    }
    print(description)
    punctuations = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''
    no_punct = ""
    for char in title:
        if char not in punctuations:
            no_punct = no_punct + char
    file = no_punct #title.replace(" ", "-").replace("[","").replace("]","").replace("","")
    with open('./results/{}.json'.format(file), 'w') as outfile:
        json.dump(dict_data, outfile)

def create_csv():
    files = sorted(glob.glob('./results/*.json'))
    datas = []
    for file in files:
        with open(file) as json_file:
            data = json.load(json_file)
            datas.append(data)
    df = pd.DataFrame(datas)
    # df.to_csv('results.csv', index=False)
    df.to_excel('results.xlsx', index=False)
    print('Generate CSV')

def run():
    get_url()
    with open('data.json') as json_file:
        data = json.load(json_file)
        for p in data['product']:
            detail_product(p['url'])
    create_csv()

if __name__ == '__main__':
    run()