import json
import os
import time
import pprint

import requests
from dotenv import load_dotenv

load_dotenv()


class sendRequest:
    # NP
    __NP_API_KEY = os.getenv('NP_API_KEY')
    __NP_PHONE = os.getenv('NP_PHONE')

    # UKR
    __PRODUCTION_BEARER_StatusTracking = os.getenv('PRODUCTION_BEARER_StatusTracking')

    def __request_repeat(self, url, data=None, headers=None, retry=5):
        """ """
        try:
            if data is not None:
                response = requests.post(url=url, data=data, headers=headers)
            else:
                response = requests.get(url=url, headers=headers)
        except Exception:
            time.sleep(3)
            if retry:
                return self.__request_repeat(url, data, headers, retry=(retry - 1))
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
        return self.__request_repeat(url=url, data=json.dumps(data), headers=headers).json()

    def request_to_ukr(self, barcode):
        """ """
        headers = {
            'Authorization': f'Bearer {self.__PRODUCTION_BEARER_StatusTracking}',
            'Content-Type': 'application/json',
        }
        url = f'https://www.ukrposhta.ua/status-tracking/0.0.1/statuses/last?barcode={barcode}'

        return self.__request_repeat(url=url, headers=headers).json()


def main():
    POST_NAME = 'UKR'

    if POST_NAME == 'NP':
        np = sendRequest()
        res = np.request_to_np(barcode='20450749477905')
        pprint.pprint(res.get('data')[0].get('Status'))

    elif POST_NAME == 'UKR':
        ukr = sendRequest()
        res = ukr.request_to_ukr(barcode='0504270497177')
        pprint.pprint(res.get('eventName'))


if __name__ == '__main__':
    main()
