from abc import ABC, abstractmethod
from typing import Any, Dict

import requests
from logger import logger


class BaseAPIClient(ABC):
    @abstractmethod
    def get_status(self, barcode: str) -> str:
        """Абстрактний метод, який повинен бути реалізований в дочірніх класах"""
        pass

    def _make_request(self, url: str, method: str = "GET", **kwargs) -> Dict[Any, Any]:
        try:
            response = requests.request(method, url, **kwargs)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            logger.error(f"API request failed: {e}")
            raise
