import base64
import datetime
import gzip
import http.client
import json
import math
import os
import tempfile
import time
from collections import OrderedDict
from urllib.parse import urlparse
import random
from urllib.request import urlopen, Request

import openpyxl
import pygsheets
import requests
from bs4 import BeautifulSoup
from google.auth.transport import requests as req
import pandas as pd

import test
from fake_MS import *

moi_sclad_token = '3a702a18a2627dd392a7d467b2d335aa528b3703'
API_WB = (
    "eyJhbGciOiJFUzI1NiIsImtpZCI6IjIwMjQwMjI2djEiLCJ0eXAiOiJKV1QifQ.eyJlbnQiOjEsImV4cCI6MTcyODk0NzMwNSwiaWQiOiJiMGMzYTllNC0zMDU0LTQ2MzgtODMyZS1hMWU1ZjY5NmI1ZWUiLCJpaWQiOjEyMTAyMjE5LCJvaWQiOjIxNzA5NywicyI6NTEwLCJzaWQiOiJhODEwYzI0NC0zNTZkLTQwYmUtYjQzMC00NWQ3NWUzZGY5ODgiLCJ0IjpmYWxzZSwidWlkIjoxMjEwMjIxOX0.E1O4eDZ1--U_dRD3b5v4kT9x_A9AT5_5m_GYRFIIjXYt_lWWoaTJOfeJ2_8IAsIqhORuLpgWP_-_gS6gIFbfFw")

OZON_CLIENT_ID_USA = '163276'
OZON_USA_TOKEN = '6d33c2e7-6b37-4814-9cca-835bc5cfaeed'
X10_OZON_CLIENT_ID_USA = '1261586'
X10_OZON_USA_TOKEN = '56edbddd-7de7-4569-925d-5806050482b9'
YANDEX_CLIENT_ID = '96d46df945294aef8f2f6893d24441a0'
CLIENT_SECRET_ID = 'dd5b2fae7e4344d1a020cf225c84d638'
YANDEX_TOKEN = 'y0_AgAAAABvN8ktAAvjhQAAAAEGbFyUAADJiTQ0hJJL84AO3hkbydMiEkEBjw'

now = datetime.datetime.now()
week_ago = now - datetime.timedelta(weeks=1)
begin_time = str(week_ago.strftime("%Y-%m-%d %H:%M:%S"))
end_time = str(now.strftime("%Y-%m-%d %H:%M:00"))

users = [{
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.81 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.104 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.86 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.94 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.139 Safari/537.36"
    },
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.87 Safari/537.36"
    }]


def compare_moment(x, y):
    return datetime.datetime.strptime(x['moment'], '%Y-%m-%d %H:%M:%S') - datetime.datetime.strptime(y['moment'],
                                                                                                     '%Y-%m-%d %H:%M:%S')

def complex_to_string(value):
    if isinstance(value, dict):
        return json.dumps(value, ensure_ascii=False)
    elif isinstance(value, (list, tuple)):
        return json.dumps(value, ensure_ascii=False)
    else:
        return str(value)

def find_value(dictionary, key):
    if key in dictionary:
        return dictionary[key]
    else:
        for k in dictionary:
            if isinstance(dictionary[k], dict):
                result = find_value(dictionary[k], key)
                if result is not None:
                    return result
        return None


def split_list(lst, n):
    return [lst[i:i + n] for i in range(0, len(lst), n)]


def get_yandex_prods(prods_data_moysclad):
    header = {"Authorization": f"Bearer {YANDEX_TOKEN}"}
    body = {
        "businessId": 74559019,
        "dateFrom": str(week_ago.strftime("%Y-%m-%d")),
        "dateTo": str(now.strftime("%Y-%m-%d")),
        "grouping": "OFFERS"
    }
    conn = http.client.HTTPSConnection(
        urlparse('https://api.partner.market.yandex.ru/reports/shows-sales/generate').hostname)
    conn.request('POST', 'https://api.partner.market.yandex.ru/reports/shows-sales/generate',
                 json.dumps(body), headers=header)
    response = conn.getresponse()
    data = response.read()
    result = json.loads(data)
    conn.close()
    time.sleep(250)
    anal_id = result['result']["reportId"]
    conn = http.client.HTTPSConnection(
        urlparse(f'https://api.partner.market.yandex.ru/reports/info/{anal_id}').hostname)
    conn.request('GET', f'https://api.partner.market.yandex.ru/reports/info/{anal_id}', headers=header)
    response = conn.getresponse()
    data = response.read()
    result = json.loads(data)
    conn.close()
    try:
        file = result['result']['file']
    except:
        time.sleep(250)
        conn = http.client.HTTPSConnection(
            urlparse(f'https://api.partner.market.yandex.ru/reports/info/{anal_id}').hostname)
        conn.request('GET', f'https://api.partner.market.yandex.ru/reports/info/{anal_id}', headers=header)
        response = conn.getresponse()
        data = response.read()
        result = json.loads(data)
        conn.close()
        file = result['result']['file']


    with tempfile.TemporaryDirectory() as tmpdirname:
        temp_file_path = os.path.join(tmpdirname, 'file.xlsx')
        response = requests.get(file)
        with open(temp_file_path, 'wb') as temp_file:
            temp_file.write(response.content)

        workbook = openpyxl.load_workbook(temp_file_path)
        sheet = workbook.active
        for i in range(0, len(prods_data_moysclad)):

            for row in sheet.iter_rows(values_only=True):
                # Если значение в колонке F совпадает с переменной
                if prods_data_moysclad[i]['article'] in row[3]:


                    if prods_data_moysclad[i].get('ordered_ya', None):
                        print(f'Была цена {prods_data_moysclad[i]["ordered_ya"]}')
                        prods_data_moysclad[i]['ordered_ya'] = float(prods_data_moysclad[i]['ordered_ya']) + float(row[10])
                        print(f'Стала цена {prods_data_moysclad[i]["ordered_ya"]}')

                    else:
                        prods_data_moysclad[i]['ordered_ya'] = row[10]
                        body = {'offerIds': [row[3]]}
                        conn = http.client.HTTPSConnection(
                            urlparse(
                                f'https://api.partner.market.yandex.ru/businesses/74559019/offer-mappings').hostname)
                        conn.request('POST', f'https://api.partner.market.yandex.ru/businesses/74559019/offer-mappings',
                                     json.dumps(body), headers=header)
                        response = conn.getresponse()
                        data = response.read()
                        result = json.loads(data)
                        conn.close()
                        price = result['result']['offerMappings'][0]['offer']['basicPrice']['value']
                        prods_data_moysclad[i]['price_ya'] = price


                    prods_data_moysclad[i]['article_ya'] = prods_data_moysclad[i]['article']
                    print('Артикул яндекс: ', prods_data_moysclad[i]['article_ya'])



        # Закрыть файл
        workbook.close()
    for i in range(0, len(prods_data_moysclad)):
        if prods_data_moysclad[i].get('ordered_ya', None):
            pass
        else:
            prods_data_moysclad[i]['article_ya'] = ''
            prods_data_moysclad[i]['ordered_ya'] = ''
            prods_data_moysclad[i]['price_ya'] = ''
    print(prods_data_moysclad)
    return prods_data_moysclad


def process_cards(cards, prods_for_pricing_and_anal):
    period = {"begin": begin_time, "end": end_time}
    body = {"nmIDs": [int(card["nmID"]) for card in cards], "timezone": "Europe/Moscow", "period": period, "page": 1}

    headers = {'Content-Type': 'application/json', 'Authorization': API_WB}

    try:
        response = requests.post('https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail', json=body,
                                 headers=headers)
        result = response.json()
        llls = result['data']["cards"]
        for pr in prods_for_pricing_and_anal:
            for lll in llls:
                if pr['article'] == lll['vendorCode']:
                    # temp_prod = {**pr}
                    # prods_for_pricing_and_anal.remove(temp_prod)
                    pr['sold_wb'] = find_value(lll, 'buyoutsCount')
                    pr['article_wb'] = lll["nmID"]
                    # prods_for_pricing_and_anal.append(temp_prod)
    except Exception as e:
        print(f"Response error: {str(e)}")
        time.sleep(61)
        response = requests.post('https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail', json=body,
                                 headers=headers)
        result = response.json()
        llls = result['data']["cards"]
        for pr in prods_for_pricing_and_anal:
            for lll in llls:
                if pr['article'] == lll['vendorCode']:
                    pr['sold_wb'] = find_value(lll, 'buyoutsCount')
                    pr['article_wb'] = lll["nmID"]

# def process_card(card, prod, prods_for_pricing_and_anal):
#     cardd = [int(card["nmID"])]
#     period = {"begin": begin_time, "end": end_time}
#     body = {"nmIDs": cardd, "timezone": "Europe/Moscow", "period": period, "page": 1}
#
#     headers = {'Content-Type': 'application/json', 'Authorization': API_WB}
#
#     while True:
#         response = requests.post('https://seller-analytics-api.wildberries.ru/api/v2/nm-report/detail', json=body, headers=headers)
#         try:
#             result = response.json()
#             lll = result['data']["cards"][0]
#             temp_prod = {**prod}
#             temp_prod['sold_wb'] = find_value(lll, 'buyoutsCount')
#             temp_prod['arcticle_wb'] = cardd[0]
#             prods_for_pricing_and_anal.pop(prods_for_pricing_and_anal.index(prod))
#             prods_for_pricing_and_anal.append(temp_prod)
#             break
#         except Exception as e:
#             print(f"Response error: {str(e)}")
#             time.sleep(60)


def get_wb_prods(prods_for_pricing_and_anal):
    print(f'Товаров всего - {len(prods_for_pricing_and_anal)}')

    headers = {
        'Content-Type': 'application/json',
        'Authorization': API_WB
    }

    body = {
        "settings": {
            "sort": {
                "ascending": False
            },
            "cursor": {
                "limit": 100
            },
            "filter": {
                "textSearch": "",
                "allowedCategoriesOnly": True,
                "tagIDs": [],
                "objectIDs": [],
                "brands": [],
                "imtID": 0,
                "withPhoto": -1
            }
        }
    }

    def fetch_data(body):
        request = Request('https://suppliers-api.wildberries.ru/content/v2/get/cards/list', data=json.dumps(body).encode(), headers=headers)
        try:
            response = urlopen(request)
            return response.read().decode()
        except Exception as e:
            print(f'Response error: {str(e)}')
            time.sleep(61)
            return fetch_data(body)

    alll = fetch_data(body)
    all_cards = [*json.loads(alll)['cards']]
    cursor = json.loads(alll)['cursor']

    updated_at = cursor['updatedAt']
    nm_id = cursor['nmID']
    total = cursor['total']
    body['settings']['cursor']['updatedAt'] = updated_at
    body['settings']['cursor']['nmID'] = nm_id

    while total >= 100:
        print('Полуачем след страницы товаров')
        try:
            total = int(cursor['total'])
        except Exception as e:
            print('Товары закончились')

        if total >= 100:
            try:
                alll = fetch_data(body)
                cursor = json.loads(alll)['cursor']
                updated_at = cursor['updatedAt']
                nm_id = cursor['nmID']
                body['settings']['cursor']['updatedAt'] = updated_at
                body['settings']['cursor']['nmID'] = nm_id
                all_cards.extend(json.loads(alll)['cards'])
            except Exception as e:
                print(f'Response error: {str(e)}')
                time.sleep(65)
                alll = fetch_data(body)
                cursor = json.loads(alll)['cursor']
                updated_at = cursor['updatedAt']
                nm_id = cursor['nmID']
                body['settings']['cursor']['updatedAt'] = updated_at
                body['settings']['cursor']['nmID'] = nm_id
                all_cards.extend(json.loads(alll)['cards'])
        else:
            try:
                alll = fetch_data(body)
                all_cards.extend(json.loads(alll)['cards'])
            except Exception as e:
                print(f'Response error: {str(e)}')
                time.sleep(65)
                alll = fetch_data(body)
                all_cards.extend(json.loads(alll)['cards'])


        process_cards(all_cards, prods_for_pricing_and_anal) # Поменять на массовое получение информации о товарах

    print(total)
    for prod in prods_for_pricing_and_anal:
        fgfg = {"limit": 10, "filterNmID": prod.get('article_wb', None)}
        if prod.get('sold_wb', None) is None:
            prod['article_wb'] = ''
            prod['sold_wb'] = ''
        else:
            url = 'https://discounts-prices-api.wb.ru/api/v2/list/goods/filter'
            url += '?' + '&'.join([f"{k}={v}" for k, v in fgfg.items()])
            try:
                try:
                    request = Request(url, headers=headers, method='GET')
                    response = urlopen(request)
                    response = json.loads(response.read().decode())
                    price = response['data']['listGoods'][0]['sizes'][0]['price']
                    prod['price_wb'] = price
                except Exception as e:
                    print(f'Response error: {str(e)}')
                    time.sleep(65)
                    request = Request(url, headers=headers, method='GET')
                    response = urlopen(request)
                    response = json.loads(response.read().decode())
                    price = response['data']['listGoods'][0]['sizes'][0]['price']
                    prod['price_wb'] = price
            except Exception as e:
                prod['price_wb'] = ''

    print(prods_for_pricing_and_anal)
    return prods_for_pricing_and_anal

def get_ozon_prods(prods_data_moysclad):
    ozon_prods_1 = ozon_inside_part(OZON_CLIENT_ID_USA, OZON_USA_TOKEN, prods_data_moysclad, 'usa-rusa')
    time.sleep(35)
    ozon_prods_2 = ozon_inside_part(X10_OZON_CLIENT_ID_USA, X10_OZON_USA_TOKEN, ozon_prods_1, 'x10')
    return ozon_prods_2



def get_last_invents_prods():
    bad_cats = ['', None, 'Архив', 'Архив/Архив Сырья', 'Архив/GRA', 'Офис/Менеджеры/Симкарты/Активные',
                'Офис/Менеджеры/Канцелярия', 'Офис/Менеджеры/Симкарты', 'Офис/Менеджеры/Симкарты/Неактивные',
                'Сырьё/Перефасовка', 'Офис/Менеджеры/Симкарты/Проблемы', 'Офис/Производство/Для упаковки',
                'Офис/Производство/Экипировка', 'Офис/Менеджеры/техника', 'Сырьё/Порошки', 'Сырьё/Расходники',
                'Склад/Оборудование', 'Сырьё/Этикетки']

    headers = {
        "Authorization": "Bearer " + moi_sclad_token,
        "Accept-Encoding": "gzip"
    }

    prods_data_moysclad = []
    offsets = [0, 1000, 2000]
    url_prods = []
    for o in offsets:
        conn = http.client.HTTPSConnection(
            urlparse(f'https://api.moysklad.ru/api/remap/1.2/report/stock/bystore?limit=1000&offset={o}').hostname)
        conn.request('GET', f'https://api.moysklad.ru/api/remap/1.2/report/stock/bystore?limit=1000&offset={o}',
                     json.dumps({}), headers)
        response = conn.getresponse()

        # Check if the response is gzip-compressed
        if response.getheader('Content-Encoding') == 'gzip':
            data = gzip.decompress(response.read())
            data = data.decode('utf-8')
        else:
            data = response.read().decode('utf-8')

        data = json.loads(data)
        conn.close()
        rows = data['rows']
        if rows:
            print('Получаем ссылки МС')
            for r in rows:
                try:
                    url_prods.append(r)
                except:
                    pass

    for u in url_prods:
        try:
            response = requests.get(u['meta']['href'], headers=headers)
            res = response.json()
            if res['pathName'] in bad_cats:
                continue
            else:
                all_c = 0
                for s in u['stockByStore']:
                    all_c += int(s['stock'])
                if all_c > 0:
                    prods_data_moysclad.append({"url": u['meta']['href'], "name": res['name'], 'article': res['code']})
                else:
                    continue
        except Exception as e:
            print(f"Error: {e}")
            continue

    print(prods_data_moysclad)
    return prods_data_moysclad

def ozon_inside_part(client_id, api_token, prods_data_moysclad, whats_ozon):
    headers = {
        'Client-Id': client_id,
        'Api-Key': api_token,
        'Content-Type': 'application/json'
    }
    arts = [prod['article'] for prod in prods_data_moysclad if 'article' in prod]


    splitted_list = split_list(arts, 1000)
    for l in splitted_list:
        time.sleep(20)
        body = {

            "language": "DEFAULT",
            "offer_id": l,
            "search": "",
            "sku": [],
            "visibility": "ALL"

        }

        conn = http.client.HTTPSConnection(
            urlparse('https://api-seller.ozon.ru/v1/report/products/create').hostname)
        conn.request('POST', 'https://api-seller.ozon.ru/v1/report/products/create', json.dumps(body), headers=headers)
        response = conn.getresponse()
        data = response.read()
        result = json.loads(data)
        conn.close()
        body_2 = {
            "code": result['result']['code']
        }
        time.sleep(60)
        conn = http.client.HTTPSConnection(
            urlparse('https://api-seller.ozon.ru/v1/report/info').hostname)
        conn.request('POST', 'https://api-seller.ozon.ru/v1/report/info', json.dumps(body_2), headers=headers)
        response = conn.getresponse()
        data = response.read()
        result = json.loads(data)
        conn.close()
        df = pd.read_csv(result["result"]['file'], sep=';', encoding='utf-8',
                         index_col=False)  # DEFAULT!!!!!!!!!!!!!!!!!!!!!!

        df = df[['Артикул', 'Barcode', 'FBO OZON SKU ID', 'Текущая цена с учетом скидки, ₽']]
        dictionary = df.set_index('Артикул').apply(lambda x: [x['Barcode'], x['FBO OZON SKU ID'], x['Текущая цена с учетом скидки, ₽']], axis=1).to_dict()
        new_dictionary = {k.replace("'", ""): v for k, v in dictionary.items()}
        for prod in prods_data_moysclad:
            article = prod['article']
            if article in new_dictionary:
                if new_dictionary[article] != 0:
                    try:

                        if 'OZN' in new_dictionary[article][0]:
                            prod[whats_ozon]['barcode_ozon'] = new_dictionary[article][0]
                            prod[whats_ozon]['sku_ozon'] = new_dictionary[article][1]
                            prod[whats_ozon]['price_ozon'] = new_dictionary[article][2]
                        else:
                            prod[whats_ozon]['barcode_ozon'] = ''
                            prod[whats_ozon]['sku_ozon'] = ''
                            prod[whats_ozon]['price_ozon'] = ''
                    except:
                        if type(new_dictionary[article][0]) != type(2.5):
                            if 'OZN' in new_dictionary[article][0]:
                                prod[whats_ozon] = {'barcode_ozon': new_dictionary[article][0]}
                                prod[whats_ozon]['sku_ozon'] = new_dictionary[article][1]
                                prod[whats_ozon]['price_ozon'] = new_dictionary[article][2]
                            else:
                                prod[whats_ozon] = {'barcode_ozon': ''}
                                prod[whats_ozon]['sku_ozon'] = ''
                                prod[whats_ozon]['price_ozon'] = ''
                        else:
                            prod[whats_ozon] = {'barcode_ozon': ''}
                            prod[whats_ozon]['sku_ozon'] = ''
                            prod[whats_ozon]['price_ozon'] = ''

                    print(prod)

            else:
                pass
                # prod[whats_ozon] = {'barcode_ozon': ''}
                # prod[whats_ozon]['sku_ozon'] = ''
                # prod[whats_ozon]['price_ozon'] = ''

        article_generator = [prod_data['article'] for prod_data in prods_data_moysclad]
        for new_art_k, new_art_v in new_dictionary.items():
            if new_art_k not in article_generator:
                prods_data_moysclad.append({'article': new_art_k, whats_ozon: {'barcode_ozon': new_art_v[0], 'sku_ozon': new_art_v[1], 'price_ozon': new_art_v[2]}})

        body_counts_sell = {
            "date_from": begin_time,
            "date_to": end_time,
            "metrics": [
                "ordered_units"
            ],
            "dimension": [
                "sku"
            ],
            "filters": [],
            "sort": [
                {
                    "key": "ordered_units",
                    "order": "DESC"
                }
            ],
            "limit": 1000,
            "offset": 0
        }
        conn = http.client.HTTPSConnection(
            urlparse('https://api-seller.ozon.ru/v1/analytics/data').hostname)
        conn.request('POST', 'https://api-seller.ozon.ru/v1/analytics/data', json.dumps(body_counts_sell),
                     headers=headers)
        response = conn.getresponse()
        data = response.read()
        result = json.loads(data)
        conn.close()
        if result.get("result", None) is None:
            time.sleep(65)
            conn.request('POST', 'https://api-seller.ozon.ru/v1/analytics/data', json.dumps(body_counts_sell),
                         headers=headers)
            response = conn.getresponse()
            data = response.read()
            result = json.loads(data)
            conn.close()

        for d in result["result"]["data"]:
            # Проверяем, есть ли словарь в prods_data_moysclad с таким же sku_ozon
            for i in range(len(prods_data_moysclad)):
                try:
                    if str(prods_data_moysclad[i][whats_ozon].get("sku_ozon", None)) == str(d["dimensions"][0]["id"]):
                        # if prods_data_moysclad[i][whats_ozon].get("sold_ozon", None) and prods_data_moysclad[i][whats_ozon].get("sold_ozon", None) != '':
                        if prods_data_moysclad[i][whats_ozon].get("sold_ozon", None):
                            prods_data_moysclad[i][whats_ozon]["sold_ozon"] = round(
                                float(prods_data_moysclad[i][whats_ozon]["sold_ozon"])) + round(float(d["metrics"][0]))
                        else:
                            prods_data_moysclad[i][whats_ozon]["sold_ozon"] = round(float(d["metrics"][0]))
                            break
                except:
                    pass

    for i in range(0, len(prods_data_moysclad)):
        # if prods_data_moysclad[i].get('barcode_ozon', None):
        try:
            if prods_data_moysclad[i].get(whats_ozon, None):
                if prods_data_moysclad[i][whats_ozon].get("sold_ozon", None) is None or prods_data_moysclad[i][whats_ozon].get("sold_ozon", None) == '':
                    prods_data_moysclad[i][whats_ozon]['sold_ozon'] = 0
                if prods_data_moysclad[i][whats_ozon].get("sku_ozon", None) is None or prods_data_moysclad[i][whats_ozon].get("sku_ozon", None) == '':
                    prods_data_moysclad[i][whats_ozon] = ''
            else:
                prods_data_moysclad[i][whats_ozon] = ''
        except:
            print(f'Ошибка: {prods_data_moysclad[i]}')
    return prods_data_moysclad

def write_to_table(prods_data):
    gc = pygsheets.authorize(service_account_file=f"{os.getcwd()}/test-mps-rentable-02e4e2bdc3d4.json".replace('\\', '/'))
    sh = gc.open('Pythonvauto')

    # Открытие листа по имени
    wks = sh.worksheet_by_title('Лист1')
    wks.clear()
    # Преобразование данных в формат, подходящий для записи в таблицу
    headers = ['url', 'name', 'article', 'article_wb', 'sold_wb', "price_wb", "usa-rusa", "x10", 'article_ya', 'ordered_ya', 'price_ya']

    # rows = [list(OrderedDict(sorted(item.items(), key=lambda t: headers.index(t[0]))).values()) for item in prods_data]
    wks.update_row(1, headers)
    rows = []
    for item in prods_data:
        row = []
        for header in headers:
            value = item.get(header, '')
            row.append(complex_to_string(value))
        print(row)
        rows.append(row)
    # Запись данных
    wks.update_values('A2', rows)
def main():
    write_to_table(get_yandex_prods(get_ozon_prods(get_wb_prods(ms))))
    # get_ozon_prods(ms)
    # write_to_table(get_ozon_prods(get_last_invents_prods()))


if __name__ == '__main__':
    main()