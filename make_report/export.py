from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import requests
from tqdm import tqdm

from .config import PromExportConfig
from .logger import logger
from .status import StatusService


class ExportProm:
    def __init__(self, config: PromExportConfig, status_service: StatusService):
        self.config = config
        self.status_service = status_service

    def __valid_pp(self, sku: str, quantity: int) -> float:
        """Валідація та розрахунок ціни"""
        if '||' in sku:
            return float(sku.split('||')[-1]) * quantity * self.config.current_course
        elif '|' in sku:
            return float(sku.split('|')[-1]) * quantity
        logger.warning(f"Didn't find price: {sku}")
        return 0

    def get_delivery_data(self, id: int, doid: int, ctp: Optional[float] = None) -> Tuple[Optional[str], Optional[float]]:
        """Отримання даних доставки"""
        if doid == 4898969:  # Nova Poshta
            params = {
                'order_id': str(id),
                'delivery_option_id': str(doid),
                'cart_total_price': str(ctp) if ctp else None,
            }
            response = requests.get(
                'https://my.prom.ua/remote/delivery/nova_poshta/init_data_order',
                params=params,
                cookies=self.config.cookies,
                headers=self.config.headers,
            ).json()
            return response.get('data', {}).get('intDocNumber'), response.get('data', {}).get('packageCost')

        elif doid == 10119216:  # UkrPoshta
            # ... код для УкрПошти ...
            pass

    def _get_pagination(self) -> int:
        """Отримати кількість сторінок для пагінації"""
        try:
            params = {
                "status_id": self.config.custom_status_id,
                "date_from": self._get_date_from(),
                "limit": 100
            }
            response = requests.get(
                f"{self.base_url}/orders/list",
                headers=self.headers,
                params=params
            )
            response.raise_for_status()
            data = response.json()
            return (data['count'] // 100) + 1
        except Exception as e:
            logger.error(f"Failed to get pagination: {e}")
            return 0

    def _get_date_from(self) -> str:
        """Отримати дату початку для фільтрації"""
        if self.config.month == 'current_month':
            return datetime.now().strftime("%Y-%m-01")
        return f"2024-{self.config.month}-01"

    def _get_orders(self, page: int) -> List[Dict[Any, Any]]:
        """Отримати замовлення для конкретної сторінки"""
        try:
            params = {
                "status_id": self.config.custom_status_id,
                "date_from": self._get_date_from(),
                "limit": 100,
                "page": page
            }
            response = requests.get(
                f"{self.base_url}/orders/list",
                headers=self.headers,
                params=params
            )
            response.raise_for_status()
            return response.json()['orders']
        except Exception as e:
            logger.error(f"Failed to get orders for page {page}: {e}")
            return []

    def _process_order_item(self, item: Dict, basic_info: Dict, is_first: bool) -> Dict:
        """Обробка окремого товару в замовленні"""
        try:
            purchase_price = self._calculate_purchase_price(
                item['sku'],
                item['quantity']
            )

            return {
                "ID": basic_info['id'] if is_first else '',
                "Дата": basic_info['date'] if is_first else '',
                "Статус": basic_info['status'] if is_first else '',
                "ТТН": basic_info['ttn'] if is_first else '',
                "Товар": item['name'],
                "Кількість": item['quantity'],
                "Ціна продажу": item['price'],
                "Ціна закупки": purchase_price,
                # Додайте інші поля за необхідності
            }
        except Exception as e:
            logger.error(f"Failed to process order item: {e}")
            return {}

    def _calculate_purchase_price(self, sku: str, quantity: int) -> float:
        """Розрахунок ціни закупки"""
        try:
            if '||' in sku:
                return float(sku.split('||')[-1]) * quantity * self.config.current_course
            elif '|' in sku:
                return float(sku.split('|')[-1]) * quantity
            logger.warning(f"Couldn't find price in SKU: {sku}")
            return 0
        except Exception as e:
            logger.error(
                f"Failed to calculate purchase price for SKU {sku}: {e}")
            return 0

    def get_data(self) -> List[Dict]:
        """Основний метод для отримання всіх даних"""
        try:
            data_list = []
            pagination = self._get_pagination()

            if not pagination:
                logger.error("Failed to get pagination")
                return data_list

            for page in tqdm(range(1, pagination + 1), desc="Processing orders"):
                orders = self._get_orders(page)
                for order in orders:
                    basic_info = {
                        'id': order['id'],
                        'date': order['date_created'],
                        'status': order['status'],
                        'ttn': order.get('delivery_info', {}).get('tracking_number', '')
                    }

                    for idx, item in enumerate(order['items']):
                        order_data = self._process_order_item(
                            item,
                            basic_info,
                            is_first=(idx == 0)
                        )
                        if order_data:
                            data_list.append(order_data)

            return data_list

        except Exception as e:
            logger.error(f"Failed to get data: {e}")
            return []

    def export_to_excel(self, data: List[Dict], filename: str):
        """Експорт даних в Excel"""
        try:
            import pandas as pd
            df = pd.DataFrame(data)
            df.to_excel(filename, index=False)
            logger.info(f"Data exported successfully to {filename}")
        except Exception as e:
            logger.error(f"Failed to export data to Excel: {e}")
