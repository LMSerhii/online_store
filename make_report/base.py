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

    def _make_request(self, url: str, method: str = "GET", retry: int = 5, **kwargs) -> Dict[Any, Any]:
        kwargs.setdefault("timeout", 15)
        try:
            response = requests.request(method, url, **kwargs)

            if response.status_code == 404:
                logger.warning(f"Resource not found (404): {url}")
                return {}

            if 400 <= response.status_code < 500:
                logger.error(f"Client error {response.status_code}: {url}")
                return {}

            response.raise_for_status()
            return response.json()

        except requests.exceptions.Timeout as e:
            logger.error(f"Request timed out: {e}")
            if retry:
                attempt = 6 - retry
                sleep_time = min(2 ** attempt, 30)
                logger.warning(f"Retry: {retry} (wait {sleep_time}s)")
                time.sleep(sleep_time)
                return self._make_request(url, method, retry=(retry - 1), **kwargs)
            raise

        except requests.exceptions.ConnectionError as e:
            logger.error(f"Connection error: {e}")
            if retry:
                attempt = 6 - retry
                sleep_time = min(2 ** attempt, 30)
                logger.warning(f"Retry: {retry} (wait {sleep_time}s)")
                time.sleep(sleep_time)
                return self._make_request(url, method, retry=(retry - 1), **kwargs)
            raise

        except requests.exceptions.HTTPError as e:
            if retry and e.response is not None and e.response.status_code >= 500:
                attempt = 6 - retry
                sleep_time = min(2 ** attempt, 30)
                logger.error(f"Server error {e.response.status_code}: {e}")
                logger.warning(f"Retry: {retry} (wait {sleep_time}s)")
                time.sleep(sleep_time)
                return self._make_request(url, method, retry=(retry - 1), **kwargs)
            logger.error(f"HTTP error (no retry): {e}")
            raise

        except requests.RequestException as e:
            logger.error(f"API request failed: {e}")
            raise

    def get(self, url: str, **kwargs) -> Dict[Any, Any]:
        return self._make_request(url, "GET", **kwargs)

    def post(self, url: str, **kwargs) -> Dict[Any, Any]:
        return self._make_request(url, "POST", **kwargs)
