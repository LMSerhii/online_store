import json
import re
import os

import ezsheets

from dotenv import load_dotenv

from openpyxl import load_workbook

load_dotenv()


class GPrice:
    __ITEMS_SHEET_LIST = []
    __ITEMS_EXCEL_LIST = []

    def __init__(self, sheet_id, path_file=None, vcc_gs='B', price_col_gs='E', av_col_gs='H', vcc_ex='B',
                 price_col_ex='H', av_col_ex='G'):
        self.sheet_id = sheet_id
        self.path_file = path_file
        self.vcc_gs = vcc_gs
        self.price_col_gs = price_col_gs
        self.av_col_gs = av_col_gs
        self.vcc_ex = vcc_ex
        self.price_col_ex = price_col_ex
        self.av_col_ex = av_col_ex

    def __valid_vendor_code(self, vc):
        title_list = ['ФОТО', 'Артикул', 'Наименование товара', 'Основные параметры',
                      'Цена/USD', 'Цена/UAH', 'Количество в ящике', 'Наличие', 'Цена']

        if vc == '':
            return True
        elif vc is None:
            return True

        for name in title_list:
            if re.match(fr'^{name}', vc):
                return True

    def connect_to_google(self):
        ss = ezsheets.Spreadsheet(self.sheet_id)
        sheet = ss.sheets[0]
        return sheet

    def __add_to_list_sheet(self, sheet):
        for i in range(1, sheet.rowCount + 1):
            vc_sheet = sheet[f'{self.vcc_gs}{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            self.__ITEMS_SHEET_LIST.append(vc_sheet)

    def __connect_to_excel(self):
        ee = load_workbook(self.path_file)
        excel_sheet = ee.active
        return excel_sheet

    def __add_to_list_excel(self, excel_sheet):
        for j in range(1, excel_sheet.max_row + 1):
            vc_excel = excel_sheet[f'{self.vcc_ex}{j}'].value

            if self.__valid_vendor_code(vc_excel):
                continue

            self.__ITEMS_EXCEL_LIST.append(vc_excel)

    def __valid_price(self, price):

        if isinstance(price, str) and '$' in price:
            price = price.replace('$', '').replace(',', '.').strip()

        if price is None or price == '-':
            price = 0

        return float(price)

    def __not_availability(self, sheet):

        for i in range(1, sheet.rowCount + 1):
            vc_sheet = sheet[f'{self.vcc_gs}{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            for item in self.__ITEMS_SHEET_LIST:
                if vc_sheet == item:
                    sheet[f"{self.av_col_gs}{i}"] = "FALSE"

    def updatePrice(self):
        sheet = self.connect_to_google()
        self.__add_to_list_sheet(sheet)

        excel_sheet = self.__connect_to_excel()
        self.__add_to_list_excel(excel_sheet)

        for i in range(1, sheet.rowCount + 1):
            vc_sheet = sheet[f'{self.vcc_gs}{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            price = self.__valid_price(sheet[f"{self.price_col_gs}{i}"])

            availability = sheet[f"{self.av_col_gs}{i}"]

            for j in range(1, excel_sheet.max_row + 1):
                vc_excel = excel_sheet[f'{self.vcc_ex}{j}'].value

                if self.__valid_vendor_code(vc_excel):
                    continue

                price_excel = self.__valid_price(excel_sheet[f"{self.price_col_ex}{j}"].value)
                availability_excel = excel_sheet[f"{self.av_col_ex}{j}"].value

                if vc_excel == vc_sheet:

                    self.__ITEMS_SHEET_LIST.remove(vc_sheet)

                    try:
                        self.__ITEMS_EXCEL_LIST.remove(vc_excel)
                    except Exception as ex:
                        print(ex)
                        print(vc_excel)

                    if price_excel != price:
                        sheet[f"{self.price_col_gs}{i}"] = price_excel

                    if availability_excel == '+':
                        sheet[f"{self.av_col_gs}{i}"] = 'TRUE'

                    if availability_excel == '-':
                        sheet[f"{self.av_col_gs}{i}"] = 'FALSE'

                    # print(f'{vc_sheet} -- {price} -- {availability} /'
                    #       f' {vc_excel} -- {price_excel} -- {availability_excel}')

        self.__not_availability(sheet)

        # print(f'NEW ITEMS: {self.__ITEMS_EXCEL_LIST}')
        # print(f'NOT AVAILABILITY: {self.__ITEMS_SHEET_LIST}')

        data_dict = {
            'NEW ITEMS': self.__ITEMS_EXCEL_LIST,
            'NOT AVAILABILITY': self.__ITEMS_SHEET_LIST
        }

        with open(f'grand_eltos.json', 'w', encoding='utf-8') as file:
            json.dump(data_dict, file, indent=4, ensure_ascii=False)

    def __replaceSpace(self, cell_value):

        return cell_value.strip().replace('  ', '')

    def editGS(self, target_column="C"):
        sheet = self.connect_to_google()

        for i in range(1, sheet.rowCount + 1):
            cell_value = sheet[f'{target_column}{i}']

            if self.__valid_vendor_code(cell_value):
                continue

            new_value = self.__replaceSpace(cell_value)

            sheet[f'{target_column}{i}'] = new_value + " ELTOS"


def main():
    up = GPrice(
        sheet_id=os.getenv('GRAND_ELTOS'),
        path_file=r"C:\Users\user\Downloads\price.xlsx",
        vcc_ex='B',
        price_col_ex='H',
        av_col_ex='G'
    )
    up.updatePrice()

    # res = up.connect_to_google()
    # print(res.title)
    # print(os.getenv('GRAND'))

    # up.editGS(target_column="C")


if __name__ == '__main__':
    main()
