from dataclasses import dataclass, field
from typing import Dict

from dotenv import dotenv_values


@dataclass
class AppConfig:
    np_api_key: str
    np_phone: str
    ukr_bearer_token: str
    base_url: str

    @classmethod
    def from_env(cls) -> 'AppConfig':
        config = dotenv_values('.env')
        return cls(
            np_api_key=config.get('NP_API_KEY'),
            np_phone=config.get('NP_PHONE'),
            ukr_bearer_token=config.get('PRODUCTION_BEARER_StatusTracking'),
            base_url=config.get('BASE_URL')
        )


@dataclass
class PromExportConfig:
    custom_status_id: int
    month: str
    current_course: float
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
            api_token=config.get('API_TOKEN')
        )
