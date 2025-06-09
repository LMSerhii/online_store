import re
from typing import Optional

from clients import NovaPoshtaClient, UkrPoshtaClient, PromClient
from config import AppConfig, PromExportConfig
from logger import logger


class StatusService:
    def __init__(self, config: AppConfig, export_config: PromExportConfig):
        self.np_client = NovaPoshtaClient(
            config.np_api_key, config.np_phone, config.np_base_url)
        self.ukr_client = UkrPoshtaClient(
            config.ukr_bearer_token, config.ukr_base_url)
        self.prom_client = PromClient(export_config)

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
            return None
        except Exception as e:
            logger.error(f"Failed to get status for barcode {barcode}: {e}")
            return None

    def get_status_from_prom(self, barcode: str, delivery_option_id: int) -> Optional[str]:
        if not barcode:
            return None

        match delivery_option_id:
            case 4898969:
                provider = 'nova_poshta'
            case 10119216:
                provider = 'ukrposhta'
            case _:
                provider = 'nova_poshta'

        try:
            data = self.prom_client.get_status(barcode, provider)
            return data.get(f"{provider}", {}).get(f"{barcode}", {}).get('status_text', '')
        except Exception as e:
            logger.error(f"Failed to get status for barcode {barcode}: {e}")
            return None
