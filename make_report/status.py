import re
from dataclasses import dataclass
from typing import Optional

from make_report.clients import NovaPoshtaClient, UkrPoshtaClient
from make_report.logger import logger


@dataclass
class StatusConfig:
    np_api_key: str
    np_phone: str
    ukr_bearer_token: str


class StatusService:
    def __init__(self, config: StatusConfig):
        self.np_client = NovaPoshtaClient(config.np_api_key, config.np_phone)
        self.ukr_client = UkrPoshtaClient(config.ukr_bearer_token)

    def get_status(self, barcode: str) -> Optional[str]:
        if not barcode:
            return None

        try:
            if re.match(r'^(204|59)', barcode):
                return self.np_client.get_status(barcode)
            elif re.match(r'^0?50', barcode):
                barcode = f"0{barcode}" if not barcode.startswith(
                    '0') else barcode
                return self.ukr_client.get_status(barcode)
            return ''
        except Exception as e:
            logger.error(f"Failed to get status for barcode {barcode}: {e}")
            return None
