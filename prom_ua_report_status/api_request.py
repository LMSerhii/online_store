import json
import os
import time

import requests
from dotenv import load_dotenv

load_dotenv()


class sendRequest:
    # NP
    __NP_API_KEY = os.getenv('NP_API_KEY')
    __NP_PHONE = os.getenv('NP_PHONE')

    # UKR
    __PRODUCTION_BEARER_StatusTracking = os.getenv('PRODUCTION_BEARER_StatusTracking')

    def __request_repeat(self, url, method, data=None, headers=None, retry=5):
        """ """
        try:
            if data is not None:
                response = requests.post(url=url, data=data, headers=headers)
            else:
                response = requests.get(url=url, headers=headers)
        except Exception:
            time.sleep(3)
            if retry:
                return self.__request_repeat(url, method, data, headers, retry=(retry - 1))
            else:
                raise
        else:
            return response

    def request_to_np(self, barcode):
        """  """
        url = 'https://api.novaposhta.ua/v2.0/json/'

        headers = {'Content-type': 'application/json',
                   'Accept': 'text/plain',
                   'Content-Encoding': 'utf-8'}

        data = {
            "apiKey": self.__NP_API_KEY,
            "modelName": "TrackingDocument",
            "calledMethod": "getStatusDocuments",
            "methodProperties": {
                "Documents": [
                    {
                        "DocumentNumber": f"{barcode}",
                        "Phone": self.__NP_PHONE
                    }
                ]
            }
        }
        return self.__request_repeat(url, 'post', json.dumps(data), headers).json()

    def request_to_ukr(self, barcode):
        """ """
        headers = {
            'Authorization': f'Bearer {self.__PRODUCTION_BEARER_StatusTracking}',
            'Content-Type': 'application/json',
        }
        url = f'https://www.ukrposhta.ua/status-tracking/0.0.1/statuses/last?barcode={barcode}'

        return self.__request_repeat(url, 'get', headers)


def main():
    # np = sendRequest()
    # res = np.request_to_np(ttn='20450735460753')

    ukr = sendRequest()
    res = ukr.request_to_ukr(barcode='0504584249666')
    print(res)

    # with open('data/result.json', 'w', encoding='utf-8') as file:
    #     json.dump(res, file, indent=4, ensure_ascii=False)


if __name__ == '__main__':
    main()
