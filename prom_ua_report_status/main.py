import logging
import re

import colorlog
import pandas as pd
import requests
from dotenv import load_dotenv
from tqdm import tqdm

from get_status import Status
from headers_cookies import headers, cookies
from config import CONFIG


load_dotenv()


handler = colorlog.StreamHandler()
handler.setFormatter(colorlog.ColoredFormatter(
    "%(log_color)s%(levelname)s:%(name)s:%(message)s",
    log_colors={
        'DEBUG': 'cyan',
        'INFO': 'green',
        'WARNING': 'yellow',
        'ERROR': 'red',
        'CRITICAL': 'bold_red',
    }
))

logger = colorlog.getLogger('Prom.ua report status')
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)

logging.getLogger().handlers.clear()

BASE_URL = CONFIG['BASE_URL']


class ExportProm:

    def __init__(self, custom_status_id, month='current_month', current_course=41.85, status=False):
        self.custom_status_id = custom_status_id
        self.month = month
        self.current_course = current_course
        self.status = status

    def __valid_pp(self, sku, quantity):
        if '||' in sku:
            return float(sku.split('||')[-1]) * quantity * self.current_course
        elif '|' in sku:
            return float(sku.split('|')[-1]) * quantity
        else:
            print(f"Didn't find price: {sku}")

    def get_delivery_data(self, id, doid, ctp=None):
        if doid == 4898969:
            params = {
                'order_id': f'{id}',
                'delivery_option_id': f'{doid}',
                'cart_total_price': f'{ctp}',
            }

            response = requests.get(
                'https://my.prom.ua/remote/delivery/nova_poshta/init_data_order',
                params=params,
                cookies=cookies,
                headers=headers,
            ).json()

            barcode = response.get('data').get('intDocNumber')
            price = response.get('data').get('packageCost')

            return barcode, price

        elif doid == 10119216:

            params = {
                'order_id': f'{id}',
                'delivery_option_id': f'{doid}',
            }

            response = requests.get(
                'https://my.prom.ua/remote/delivery/ukrposhta/init_data_order',
                params=params,
                cookies=cookies,
                headers=headers,
            ).json()

            barcode = response.get('data').get('declarationId')
            price = response.get('data').get('declaredCost')

            return barcode, price

        else:
            barcode = None
            price = None

            return barcode, price

    def __collect_barcode(self, data):
        res = re.search(r'[0-9]{12,14}', data)
        if res:
            return res.group(0)
        else:
            return None

    def __priceFinderFromSku(self, price, sku):
        pattern = r'[-рсолхэпик ]\s*(\d+)'

        match = re.search(pattern, sku)
        if match:
            price = match.group(1)
            return price
        else:
            return price

    def _get_pagination(self):

        url = f'{BASE_URL}?custom_status_id={self.custom_status_id}&company_client_id=null&page=1&per_page=100&new_cabinet=true&search_term'

        response = requests.get(
            url=url,
            cookies=cookies,
            headers=headers,
        )

        if response.status_code != 200:
            logger.error('Server error ')

        response = response.json()

        num_pages = response.get('pagination').get('num_pages')

        if num_pages == 0:
            logger.error('No orders found')
            return None

        return num_pages

    def _get_orders(self, page):
        url = f'{BASE_URL}?custom_status_id={self.custom_status_id}&company_client_id=null&page={page}&per_page=100&new_cabinet=true&search_term'

        response = requests.get(url=url,
                                cookies=cookies,
                                headers=headers,
                                )

        return response.json()

    def get_data(self):
        """ """
        data_list = []

        pagination = self._get_pagination()

        for page in tqdm(range(1, pagination + 1)):

            orders = self._get_orders(page)
            for order in orders:
                id = order.get('id')

                order_type = order.get('type')

                client_first_name = order.get('client_first_name')
                client_last_name = order.get('client_last_name')
                client_full_name = client_last_name + ' ' + client_first_name

                payment_option_name = order.get('payment_option_name')

                labels = order.get('labels')

                comments = ', '.join(
                    [label.get('name').replace(' ', '') for label in labels])

                added_items = order.get('added_items')

                delivery_option_id = order.get('delivery_option_id')

                price_text = order.get('price_text')

                cart_total_price = price_text[:-1].replace(',', '.').replace('\xa0', '').strip() if '₴' in price_text \
                    else price_text.replace(',', '.').replace('\xa0', '').strip()
                cart_total_price = cart_total_price

                barcode, price = self.get_delivery_data(
                    id=id, doid=delivery_option_id, ctp=cart_total_price)

                if price == '':
                    price = self.__priceFinderFromSku(price, comments)

                pattern = r"Пром-оплата|олхрс|оплрс"

                try:
                    if re.search(pattern, payment_option_name):
                        pc = price
                    elif re.search(pattern, comments):
                        pc = price
                    else:
                        pc = ''
                except:
                    print("pattern", pattern)
                    print("id", id)

                pattern_1 = r'денис-(\d+)'
                match_1 = re.search(pattern_1, comments)

                pattern_2 = r'взялиидениса(\d+)'
                match_2 = re.search(pattern_2, comments)

                if match_1:
                    den = int(match_1.group(1))
                elif match_2:
                    den = -match_2.group(1)
                else:
                    den = ''

                pattern = r"дропмард-(\d+)"
                match = re.search(pattern, comments)

                if match:
                    mard = match.group(1)
                    price = int(mard)
                else:
                    mard = ''

                purchase_price = None
                margin = None

                if barcode is None:
                    barcode = self.__collect_barcode(data=comments)

                try:

                    if self.status:
                        st = Status()
                        status = st.getStatus(barcode=barcode)
                    else:
                        status = None
                except Exception as ex:
                    print(ex)
                    print(id)

                # ----------------------------------------------------------------------------------
                if len(added_items) > 1:

                    purchase_price = 0

                    for item in added_items:
                        sku = item.get('sku')
                        quantity = item.get('quantity')

                        purchase_price += self.__valid_pp(sku, quantity)

                    sku = added_items[0].get('sku')
                    quantity = added_items[0].get('quantity')

                    pattern = r'пе(\d+)'
                    match = re.search(pattern, comments)

                    if match:
                        purchase_price += match.group(1)
                        pe = match.group(1)
                    else:
                        pe = ''

                    data_list.append(
                        [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                         price, purchase_price, margin, pe, pc, den, mard, status])

                    # Остальные позиции проставляем с ценой и стоимостью доставки в ноль, что бы не дублировать
                    for item in added_items[1:]:
                        quantity = item.get('quantity')

                        sku = ''
                        price = 0
                        purchase_price = 0
                        pc = 0
                        pe = ''
                        den = 0
                        mard = 0

                        data_list.append(
                            [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                             price,
                             purchase_price, margin, pe, pc, den, mard, status])

                else:
                    for item in added_items:
                        sku = item.get('sku')
                        quantity = item.get('quantity')

                        purchase_price = self.__valid_pp(sku, quantity)

                        pattern = r'пе(\d+)'
                        match = re.search(pattern, comments)

                        if match:
                            # print(match.group(1))
                            # purchase_price += match.group(1)
                            pe = match.group(1)
                        else:
                            pe = ''

                        data_list.append(
                            [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                             price, purchase_price, margin, pe, pc, den, mard, status])

        df = pd.DataFrame(data_list,
                          columns=['id замовлення', 'Спосіб замовлення', 'ПІБ', 'Спосіб оплати', 'Кількість', 'Артикул',
                                   'Коментарі', 'ТТН', 'Ціна продажу', 'Ціна закупу', 'Прибуток', 'Ми >> Пе',
                                   'РС >> СЛ', 'Денис >> СЛ', 'Мард >> СЛ', 'Статус замовлення'])

        df.to_excel(f'data/{self.month}.xlsx', index=None, header=True)


def main():

    options = {
        'custom_status_id': 150740,
        'month': 'May',
        'current_course': 41.85,
        'status': True,
    }

    ex = ExportProm(**options)

    # ex.get_data()

    ex._get_pagination()


if __name__ == '__main__':
    main()
