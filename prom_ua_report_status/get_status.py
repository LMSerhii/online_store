import os
import re

from openpyxl import load_workbook
from tqdm import tqdm

from api_request import sendRequest



class Status:

    def getStatus(self, barcode):

        req = sendRequest()
        if re.match(r'^204|^59', f'{barcode}'):
            response = req.request_to_np(barcode=barcode)
            return response.get('data')[0].get('Status')
        elif re.match(r'^050', f'{barcode}'):
            response = req.request_to_ukr(barcode=barcode)
            return response.get('eventName')
        elif re.match(r'^50', f'{barcode}'):
            response = req.request_to_ukr(barcode=f"0{barcode}")
            return response.get('eventName')
        else:
            return ''

    def readExcel(self, path_to_path, ttn_col='H', status_col='L'):
        wb = load_workbook(filename=path_to_path)
        sheet = wb.active

        for i in tqdm(range(2, sheet.max_row + 1), position=0):
            barcode = sheet[f'{ttn_col}{i}'].value
            sheet[f'{status_col}{i}'].value = self.getStatus(barcode)

        name = os.path.split(path_to_path)[-1]

        wb.save(f"data/with_status/{name}")


def main():
    path_file = r"C:\Users\admin\Desktop\order_list_February.xlsx"
    ttn_col = 'H'
    status_col = 'L'

    st = Status()
    # st.readExcel(path_file, ttn_col, status_col)
    print(st.getStatus(barcode='0504270497177'))


if __name__ == "__main__":
    main()
