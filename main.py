import requests
import shelve
from datetime import datetime

import pandas as pd


def returning_report():
    posting_numbers = []
    endpoint = 'https://api-seller.ozon.ru/v1/report/returns/create'
    headers = {
        "Client-Id": "",
        "Api-Key": ""
    }
    data = {

    }
    response = requests.post(endpoint, headers=headers, json=data)
    if response.status_code == 200:
        return posting_numbers
    else:
        print("Error: Request returned status code", response.status_code, "with message", response.reason)

def order_returning_date(posting_number):
    endpoint = 'https://api-seller.ozon.ru/v2/returns/company/fbs'
    headers = {
        "Client-Id": "",
        "Api-Key": ""
    }
    data = {
        "filter": {
            "posting_number": [
                posting_number
            ],
            "status": "returned_to_seller"
        },
        "limit": 1,
        "offset": 0
    }
    response = requests.post(endpoint, headers=headers, json=data)
    if response.status_code == 200:
        row_date = str(response.json()['result']['returns'][0]['returned_to_seller_date_time'])
        return datetime.strptime(row_date, "%Y-%m-%dT%H:%M:%S%z").strftime("%d.%m.%Y")
    else:
        print("Error: Request returned status code", response.status_code, "with message", response.reason)

def order_shipment_date(posting_number):
    endpoint = "https://api-seller.ozon.ru/v3/posting/fbs/get"
    headers = {
        "Client-Id": "",
        "Api-Key": ""
    }
    data = {
        "posting_number": posting_number,
        "with": {
            "analytics_data": False,
            "barcodes": False,
            "financial_data": False,
            "product_exemplars": False,
            "translit": False
        }
    }
    response = requests.post(endpoint, headers=headers, json=data)
    if response.status_code == 200:
        row_date = str(response.json()['result']['shipment_date'])
        return datetime.strptime(row_date, "%Y-%m-%dT%H:%M:%SZ").strftime("%d.%m.%Y")
    else:
        print("Error: Request returned status code", response.status_code, "with message", response.reason)

def find_matching():
    returns = pd.read_excel('returns.xlsx')
    for i in range(returns.shape[0]):
        code = str(returns.iloc[i]['Артикул товара'])
        posting_number = str(returns.iloc[i]['Номер отправления'])
        if '.' in code:
            try:
                with shelve.open('data/AE-QID') as data:
                    code = data[code]
            except:
                print(f'Next code is not finding: {code}')
        returns.loc[i, 'Артикул товара'] = code
        returns.loc[i, 'Дата акта'] = order_returning_date(posting_number)
        returns.loc[i, 'Дата отгрузки'] = order_shipment_date(posting_number)
    date_string = datetime.now().strftime('%Y-%m-%d')
    returns.to_excel(f'{date_string} returns.xlsx', index=False)

if __name__ == '__main__':
    find_matching()