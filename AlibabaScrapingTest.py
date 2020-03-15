import requests
from bs4 import BeautifulSoup
import AlibabaConfig as config


url = "https://www.alibaba.com/trade/search?fsb=y&IndexArea=product_en&CatId=&SearchText=earphones&dmtrack_pageid=3b5c7ef00ab0a9875e6dbe1a170dcae99ab2147154"
headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
r = requests.get(url=url, headers=headers)



soup = BeautifulSoup(r.text, features="html.parser")

for i in soup.find_all('div'):
    productSpecs = {}
    if i.get('class') == "organic-offer-wrapper organic-gallery-offer-inner m-gallery-product-item-v2 img-switcher-parent".split(' '):

        #Title
        for j in i.find_all('h4'):
            productSpecs.update({'title': j.get('title')})

        #URL
        for j in i.find_all('a'):
            if j.get('class')[0] == "organic-gallery-title":
                print("Fetching from: ", j.get('href').split('//')[1])
                productSpecs.update({'pid': j.get('data-domdot').split('pid:')[1]})
                productSpecs.update({'url': j.get('href').split('//')[1]})

        #Price
        for j in i.find_all('p'):
            if j.get('class')[0] == "gallery-offer-price":
                productSpecs.update({'price': j.get('title')})

        #Min Order
        for j in i.find_all('p'):
            if j.get('class')[0] == "gallery-offer-minorder":
                productSpecs.update({'min_order': j.text.split('(')[0]})

        #Company Name
        for j in i.find_all('a'):
              if j.get('class')[0] == "organic-gallery-offer__seller-company":
                productSpecs.update({'supplier': j.text})

        #Country
        for j in i.find_all('span'):
            if j.get('class') == "seller-tag__country gallery-offer-seller-tag bg-visible".split(" "):
                productSpecs.update({'country': j.get('title')})

        #Year of Duration
        for j in i.find_all('span'):
            if j.get('class') == "seller-tag__year gallery-offer-seller-tag".split(" "):
                productSpecs.update({'year_of_operation': j.text})

        #Rating
        for j in i.find_all('span'):
            try:
                if j.get('class')[0] == "seb-supplier-review__score":
                    productSpecs.update({'rating': j.text})
            except:
                continue
        sheetName = 'AlibabaSample.xlsx'
        check = config.check_sheet_exists(sheetName)
        if check is False:
            config.createSheet(sheetName, productSpecs)
            config.appendSheet(sheetName, productSpecs)
        elif check is True:
            config.appendSheet(sheetName, productSpecs)
