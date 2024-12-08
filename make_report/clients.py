from make_report.base import BaseAPIClient


class NovaPoshtaClient(BaseAPIClient):
    def __init__(self, api_key: str, phone: str):
        self.api_key = api_key
        self.phone = phone
        self.base_url = 'https://api.novaposhta.ua/v2.0/json/'

    def get_status(self, barcode: str) -> str:
        data = {
            "apiKey": self.api_key,
            "modelName": "TrackingDocument",
            "calledMethod": "getStatusDocuments",
            "methodProperties": {
                "Documents": [{"DocumentNumber": barcode, "Phone": self.phone}]
            }
        }
        response = self._make_request(self.base_url, "POST", json=data)
        return response['data'][0]['Status']


class UkrPoshtaClient(BaseAPIClient):
    def __init__(self, bearer_token: str):
        self.bearer_token = bearer_token
        self.base_url = 'https://www.ukrposhta.ua/status-tracking/0.0.1'

    def get_status(self, barcode: str) -> str:
        headers = {'Authorization': f'Bearer {self.bearer_token}'}
        url = f'{self.base_url}/statuses/last?barcode={barcode}'
        response = self._make_request(url, headers=headers)
        return response['eventName']
