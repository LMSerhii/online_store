from dataclasses import dataclass, field
from typing import Dict
from headers_cookies import cookies, headers

from dotenv import dotenv_values


@dataclass
class AppConfig:
    np_api_key: str
    np_phone: str
    np_base_url: str
    ukr_bearer_token: str
    ukr_base_url: str

    @classmethod
    def from_env(cls) -> 'AppConfig':
        config = dotenv_values('.env')
        return cls(
            np_api_key=config.get('NP_API_KEY'),
            np_phone=config.get('NP_PHONE'),
            np_base_url=config.get('NP_BASE_URL'),
            ukr_bearer_token=config.get('PRODUCTION_BEARER_StatusTracking'),
            ukr_base_url=config.get('UKR_BASE_URL')
        )


@dataclass
class PromExportConfig:
    custom_status_id: int
    month: str
    current_course: float
    prom_base_url: str
    np_delivery_data_url: str
    ukr_delivery_data_url: str
    shipments_statuses_url: str
    status: bool = False
    api_token: str = ''
    cookies: Dict[str, str] = field(default_factory=dict)
    headers: Dict[str, str] = field(default_factory=dict)

    @classmethod
    def from_env(cls) -> 'PromExportConfig':
        config = dotenv_values('.env')
        return cls(
            custom_status_id=config.get('CUSTOM_STATUS_ID'),
            month=config.get('MONTH'),
            current_course=config.get('CURRENT_COURSE'),
            status=config.get('STATUS'),
            prom_base_url=config.get('PROM_BASE_URL'),
            np_delivery_data_url=config.get('NP_DELIVERY_DATA_URL'),
            ukr_delivery_data_url=config.get('UKR_DELIVERY_DATA_URL'),
            shipments_statuses_url=config.get('SHIPMENTS_STATUSES_URL'),
            api_token=config.get('API_TOKEN'),
            cookies=cookies,
            headers=headers,
        )
