import re

import ezsheets
from openpyxl import load_workbook


class GPrice:

    __ITEMS_SHEET_LIST = []
    __ITEMS_EXCEL_LIST = []

    def __init__(self, sheet_id, path_file=None):
        self.sheet_id = sheet_id
        self.path_file = path_file

    def connect_to_google(self):
        ss = ezsheets.Spreadsheet(self.sheet_id)
        sheet = ss.sheets[0]
        return sheet

    def __add_to_list_sheet(self, sheet):
        for i in range(1, 113):
            vc_sheet = sheet[f'C{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            self.__ITEMS_SHEET_LIST.append(vc_sheet)

    def __connect_to_excel(self):
        ee = load_workbook(self.path_file)
        excel_sheet = ee.active
        return excel_sheet

    def __add_to_list_excel(self, excel_sheet):
        for j in range(1, excel_sheet.max_row + 1):
            vc_excel = excel_sheet[f'B{j}'].value

            if self.__valid_vendor_code(vc_excel):
                continue

            self.__ITEMS_EXCEL_LIST.append(vc_excel)

    def __valid_vendor_code(self, vc):
        if vc == '' or vc is None:
            return True
        elif re.match(r'^Артикул', vc):
            return True

    def __valid_price(self, price):
        if isinstance(price, str) and '$' in price:
            price = price.replace('$', '').replace(',', '.').strip()
        return float(price)

    def __not_availability(self, sheet):

        for i in range(1, 113):
            vc_sheet = sheet[f'C{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            for item in self.__ITEMS_SHEET_LIST:
                if vc_sheet == item:
                    sheet[f"I{i}"] = "FALSE"

    def updatePrice(self):
        sheet = self.connect_to_google()
        self.__add_to_list_sheet(sheet)

        excel_sheet = self.__connect_to_excel()
        self.__add_to_list_excel(excel_sheet)

        for i in range(1, 113):
            vc_sheet = sheet[f'C{i}']

            if self.__valid_vendor_code(vc_sheet):
                continue

            price = self.__valid_price(sheet[f"F{i}"])
            availability = sheet[f"I{i}"]

            for j in range(1, excel_sheet.max_row + 1):
                vc_excel = excel_sheet[f'B{j}'].value

                if self.__valid_vendor_code(vc_excel):
                    continue

                price_excel = self.__valid_price(excel_sheet[f"H{j}"].value)
                availability_excel = excel_sheet[f"G{j}"].value

                if vc_excel == vc_sheet:

                    self.__ITEMS_SHEET_LIST.remove(vc_sheet)
                    self.__ITEMS_EXCEL_LIST.remove(vc_excel)

                    if price_excel != price:
                        sheet[f"F{i}"] = price_excel

                    if availability_excel == '+':
                        sheet[f"I{i}"] = 'TRUE'

                    # print(f'{vc_sheet} -- {price} -- {availability} /'
                    #       f' {vc_excel} -- {price_excel} -- {availability_excel}')

        self.__not_availability(sheet)

        print(f'NEW ITEMS: {self.__ITEMS_EXCEL_LIST}')
        print(f'NOT AVAILABILITY: {self.__ITEMS_SHEET_LIST}')


def main():
    # up = GPrice(
    #     sheet_id='1MGwJnpVdbBygL2SWSb7DQyybirzMYrH9RhsM_vzcLAI',
    #     path_file=r"C:\Users\admin\Downloads\Электроинструмент Grand д.xlsx"
    # )
    # up.updatePrice()
    print("hello")

if __name__ == '__main__':
    main()
