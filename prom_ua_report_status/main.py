import re
import os

import pandas as pd
import requests

from tqdm import tqdm

from get_status import Status
from headers_cookies import headers, cookies


class ExportProm:

    def __init__(self, custom_status_id, month='current_month', current_course=37.5, status=False):
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

            # print(barcode)
            # print(price)

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

            # print(barcode)
            # print(price)

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

    def __pePattern(self, purchase_price, comments):
        pattern = r'пе(\d+)'
        match = re.search(pattern, comments)

        if match:
            purchase_price += match.group(1)
            return purchase_price
        else:
            return purchase_price

    def get_data(self):
        """ """
        data_list = []

        response = requests.get(
            f'https://my.prom.ua/remote/order_api/orders?custom_status_id={self.custom_status_id}&'
            f'company_client_id=null&page=1&per_page=100&new_cabinet=true&search_term',
            cookies=cookies,
            headers=headers,
        ).json()

        pagination = response.get('pagination').get('num_pages')

        for page in tqdm(range(1, pagination + 1)):

            response = requests.get(
                f'https://my.prom.ua/remote/order_api/orders?custom_status_id={self.custom_status_id}&company_client_id=null&page={page}&'
                f'per_page=100&new_cabinet=true&search_term',
                cookies=cookies,
                headers=headers,
            ).json()

            orders = response.get('orders')
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

                    purchase_price = self.__pePattern(purchase_price, comments)

                    data_list.append(
                        [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                         price, purchase_price, margin, status])

                    # Остальные позиции проставляем с ценой и стоимостью доставки в ноль, что бы не дублировать
                    for item in added_items[1:]:
                        quantity = item.get('quantity')

                        sku = ''
                        price = 0
                        purchase_price = 0

                        data_list.append(
                            [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                             price,
                             purchase_price, margin, status])

                else:
                    for item in added_items:
                        sku = item.get('sku')
                        quantity = item.get('quantity')

                        purchase_price = self.__valid_pp(sku, quantity)

                        purchase_price = self.__pePattern(purchase_price, comments)

                        data_list.append(
                            [id, order_type, client_full_name, payment_option_name, quantity, sku, comments, barcode,
                             price, purchase_price, margin, status])

        df = pd.DataFrame(data_list,
                          columns=['id замовлення', 'Спосіб замовлення', 'ПІБ', 'Спосіб оплати', 'Кількість', 'Артикул',
                                   'Коментарі', 'ТТН', 'Ціна продажу', 'Ціна закупу', 'Прибуток', 'Статус замовлення'])

        df.to_excel(f'data/{self.month}.xlsx', index=None, header=True)


def main():
    ex = ExportProm(custom_status_id=144381, month='December',
                    current_course=38.5)
    ex.get_data()


if __name__ == '__main__':
    main()
