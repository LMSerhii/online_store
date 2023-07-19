import json
import math
import re
from datetime import date

import ezsheets

from openpyxl import load_workbook



class UpdatePrice:
    def __init__(self, export_path, marg, or_marg, curr, rate_sell, vcc):
        self.export_path = export_path
        self.margin = marg
        self.or_margin = or_marg
        self.curr = curr
        self.rate_sell = rate_sell
        self.vendor_code_coll = vcc

    def royalty(self, price, rate):
        result = (price / (100 - rate)) * 100
        return math.ceil(result)

    def __vendor_code(self, worksheet, itr, rate):

        vendor_code = worksheet[f'{self.vendor_code_coll}{itr}'].value

        if re.match(r'^OR|\*', vendor_code):

            if '||' in vendor_code:

                pp = vendor_code.split('||')[-1]

                new_price = self.royalty((float(pp) * self.curr + self.or_margin), rate)

                old_price = self.royalty(new_price, self.rate_sell)

                return new_price, old_price

            elif '|' in vendor_code:

                pp = vendor_code.split('|')[-1]

                if re.search(r'Т', vendor_code):
                    new_price = self.royalty((float(pp) + self.or_margin + 150), rate)
                else:
                    new_price = self.royalty((float(pp) + self.or_margin), rate)

                old_price = self.royalty(new_price, self.rate_sell)

                return new_price, old_price

        else:

            if '||' in vendor_code:

                pp = vendor_code.split('||')[-1]

                new_price = self.royalty((float(pp) * self.curr + self.margin), rate)

                old_price = self.royalty(new_price, self.rate_sell)

                return new_price, old_price

            else:

                pp = vendor_code.split('|')[-1]

                if re.search(r'Т', vendor_code):
                    new_price = self.royalty((float(pp) + self.margin + 150), rate)
                else:
                    new_price = self.royalty((float(pp) + self.margin), rate)

                old_price = self.royalty(new_price, self.rate_sell)

                return new_price, old_price

    def __availability(self, worksheet, itr):

        availability = worksheet[f'P{itr}'].value

        if availability == '+' or availability == '!':
            discount = f'{self.rate_sell}%'
            date_start = date.today().strftime('%d.%m.%Y')
            date_end = date.today().replace(day=date.today().day + 7).strftime('%d.%m.%Y')

            return discount, date_start, date_end

        else:
            discount = ''
            date_start = ''
            date_end = ''

            return discount, date_start, date_end

    def __get_rate(self, worksheet, itr, rate_column):
        export_rate_id = worksheet[f'{rate_column}{itr}'].value

        with open('prom_rate.json', 'r', encoding='utf-8') as f:
            file = json.load(f)

        for item in file:
            prom_rate_id = int(item.get('cat_id'))

            if export_rate_id == prom_rate_id:
                return float(item.get('rate')[:-1])

    def __vendor_validation(self, worksheet, itr):

        vc_export = worksheet[f'{self.vendor_code_coll}{itr}'].value

        if vc_export is None:
            return True

        if re.match(r'^Код_товара', vc_export):
            return True

        return False

    def __get_price(self, sheet_id, vc_export):
        ss = ezsheets.Spreadsheet(sheet_id)
        sheet = ss.sheets[0]

        for i in range(10, sheet.rowCount + 1):
            vendor_code = sheet[f'B{i}']

            if vendor_code == vc_export:
                # print(f'{vendor_code} == {vc_export}')

                price = sheet[f'E{i}']

                if isinstance(price, str) and '$' in price:
                    price = price.replace('$', '').replace(',', '.').strip()

                availability = sheet[f'H{i}']
            else:
                price = 0
                availability = ''

            return price, availability


    def __put_id_prom(self, worksheet, itr, sheet_id):

        vencod_export = worksheet[f'{self.vendor_code_coll}{itr}'].value
        vc_export = vencod_export.split('|')[-3]

        price, availability = self.__get_price(sheet_id=sheet_id, vc_export=vc_export)

        if price != 0:

            if availability:
                worksheet[f'P{itr}'].value = '!'
            else:
                worksheet[f'P{itr}'].value = '-'

            if re.match(r'^OR|\*', vencod_export):
                worksheet[f'{self.vendor_code_coll}{itr}'].value = f'OR|{vc_export}||000{price}'

            else:
                worksheet[f'{self.vendor_code_coll}{itr}'].value = f'{vc_export}||000{price}'

    def updateProm(self, rate_column='AA', from_price=None):
        wb = load_workbook(filename=self.export_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):

            if self.__vendor_validation(worksheet=ws, itr=i):
                continue

            rate = self.__get_rate(worksheet=ws, itr=i, rate_column=rate_column)

            if rate == None:
                rate = 0

            if from_price is not None:
                self.__put_id_prom(worksheet=ws, itr=i, sheet_id=from_price)

            new_price, old_price = self.__vendor_code(worksheet=ws, itr=i, rate=rate)
            discount, date_start, date_end = self.__availability(worksheet=ws, itr=i)

            ws[f'I{i}'].value = str(old_price)
            ws[f'AE{i}'].value = discount
            ws[f'AI{i}'].value = date_start
            ws[f'AJ{i}'].value = date_end

        wb.save(self.export_path)

        return "Successful updated"

    def updateEpik(self, rate):
        wb = load_workbook(filename=self.export_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):
            new_price, old_price = self.__vendor_code(worksheet=ws, itr=i, rate=rate)

            ws[f'E{i}'].value = new_price
            ws[f'F{i}'].value = old_price

        wb.save(self.export_path)

        return "Successful updated"

    def updateRozetka(self, rate):
        wb = load_workbook(filename=self.export_path)
        ws = wb.active

        for i in range(2, ws.max_row + 1):
            new_price, old_price = self.__vendor_code(worksheet=ws, itr=i, rate=rate)

            ws[f'I{i}'].value = new_price
            ws[f'J{i}'].value = old_price

        wb.save(self.export_path)

        return "Successful updated"


def main():
    export = UpdatePrice(
        export_path=r"C:\Users\admin\Desktop\grand_or.xlsx",
        marg=100,
        or_marg=430,
        curr=37.5,
        rate_sell=25,
        vcc='A'
    )

    print(export.updateProm(from_price='1MGwJnpVdbBygL2SWSb7DQyybirzMYrH9RhsM_vzcLAI'))

    # res = export.royalty(36 * 37.5 + 50, 0)
    # print(res)
    # print(export.royalty(res, 25))

    # while True:
    #     margin = 430
    #     rate = 18
    #     rate_sell = 35
    #     price = int(input('Enter price: '))
    #     print(prom.royalty(price + margin, rate))
    #     print(prom.royalty(prom.royalty(price + margin, rate), rate_sell))


if __name__ == '__main__':
    main()
