from abc import ABC, abstractmethod
from typing import Any, Dict
import time
import requests
from logger import logger


class BaseAPIClient(ABC):
    @abstractmethod
    def get_status(self, barcode: str) -> str:
        """Абстрактний метод, який повинен бути реалізований в дочірніх класах"""
        pass

    def _make_request(self, url: str, method: str = "GET", retry = 5, **kwargs) -> Dict[Any, Any]:
        try:
            response = requests.request(method, url, **kwargs)
            response.raise_for_status()
            return response.json()
        except requests.RequestException as e:
            logger.error(f"API request failed: {e}")
            time.sleep(3)
            if retry:
                logger.warning(f"Retry: {retry}")
                return self._make_request(url, method, retry=(retry - 1), **kwargs )
            else:
                raise

    def get(self, url: str, **kwargs) -> Dict[Any, Any]:
        return self._make_request(url, "GET", **kwargs)

    def post(self, url: str, **kwargs) -> Dict[Any, Any]:
        return self._make_request(url, "POST", **kwargs)
