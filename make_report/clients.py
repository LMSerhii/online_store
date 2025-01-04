from base import BaseAPIClient
from logger import logger
from config import PromExportConfig
from typing import Any, Dict, List, Optional, Tuple

import pprint


class NovaPoshtaClient(BaseAPIClient):
    def __init__(self, api_key: str, phone: str, base_url: str):
        self.api_key = api_key
        self.phone = phone
        self.base_url = base_url

    def get_status(self, barcode: str) -> str:
        data = {
            "apiKey": self.api_key,
            "modelName": "TrackingDocument",
            "calledMethod": "getStatusDocuments",
            "methodProperties": {
                "Documents": [
                    {
                        "DocumentNumber": barcode,
                        "Phone": self.phone
                    }
                ]
            }
        }
        response = self.post(self.base_url, json=data)
        return response['data'][0]['Status']


class UkrPoshtaClient(BaseAPIClient):
    def __init__(self, bearer_token: str, base_url: str):
        self.bearer_token = bearer_token
        self.base_url = base_url

    def get_status(self, barcode: str) -> str:
        headers = {'Authorization': f'Bearer {self.bearer_token}'}
        url = f'{self.base_url}/statuses/last?barcode={barcode}'
        response = self.get(url, headers=headers)
        return response['eventName']


class PromClient(BaseAPIClient):
    def __init__(self, config: PromExportConfig):
        self.config = config

    def get_status(self, barcode: str, provider: str) -> str:

        json_data = {
            'shipments': [
                {
                    'barcode': barcode,
                    'provider': provider,
                },
            ],
        }

        response = self.post(
            url=self.config.shipments_statuses_url,
            headers=self.config.headers,
            json=json_data
        )
        return response.get('data', {})

    def get_pagination(self) -> int:
        """Отримати кількість сторінок для пагінації"""

        try:
            params = {
                "custom_status_id": self.config.custom_status_id,
                "page": 1,
                "per_page": 100,
                "company_client_id": 'null',
                "new_cabinet": 'true',
                "search_term": '',
            }

            response = self.get(
                url=f"{self.config.prom_base_url}/orders",
                params=params,
                headers=self.config.headers,
                cookies=self.config.cookies,
            )
            return response['pagination']['num_pages']
        except Exception as e:
            logger.error(f"Failed to get pagination: {e}")
            return 0

    def get_orders(self, page: int) -> List[Dict[Any, Any]]:
        """Отримати замовлення для конкретної сторінки"""
        try:
            params = {
                "custom_status_id": self.config.custom_status_id,
                "page": page,
                "per_page": 100,
                "company_client_id": 'null',
                "new_cabinet": 'true',
                "search_term": '',
            }

            response = self.get(
                url=f"{self.config.prom_base_url}/orders",
                params=params,
                headers=self.config.headers,
                cookies=self.config.cookies,
            )

            return response['orders']
        except Exception as e:
            logger.error(f"Failed to get orders for page {page}: {e}")
            return []

    def get_delivery_data(self, id: int, delivery_option_id: int) -> Tuple[Optional[str], Optional[float]]:

        if not id:
            return None, None

        delivery_configs = {
            4898969: {  # Нова Пошта
                'url': self.config.np_delivery_data_url,
                'params': {
                    'order_id': str(id),
                    'delivery_option_id': str(delivery_option_id),
                    'cart_total_price': 200,
                },
                'barcode_key': 'intDocNumber',
                'price_key': 'packageCost'
            },
            10119216: {  # Укрпошта
                'url': self.config.ukr_delivery_data_url,
                'params': {
                    'order_id': str(id),
                    'delivery_option_id': str(delivery_option_id)
                },
                'barcode_key': 'declarationId',
                'price_key': 'declaredCost'
            }
        }

        config = delivery_configs.get(delivery_option_id)

        if not config:
            return None, None

        try:
            response = self.get(
                url=config['url'],
                params=config['params'],
                cookies=self.config.cookies,
                headers=self.config.headers
            )
            data = response.get('data', {})
            return data.get(config['barcode_key']), data.get(config['price_key'])
        except Exception as e:
            logger.error(f"Failed to get delivery data for order {id}: {e}")
            return None, None
