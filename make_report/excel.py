from typing import Optional

import pandas as pd
from logger import logger
from status import StatusService


class ExcelProcessor:
    def __init__(self, status_service: StatusService):
        self.status_service = status_service

    def process_file(self, filepath: str, ttn_col: str = 'H', status_col: str = 'L') -> Optional[bool]:
        try:
            logger.info(f"Processing file: {filepath}")
            df = pd.read_excel(filepath)

            # Обробка статусів
            df[status_col] = df[ttn_col].apply(
                lambda x: self.status_service.get_status(
                    str(x)) if pd.notna(x) else ''
            )

            # Збереження результатів
            df.to_excel(filepath, index=False)
            logger.info(f"Successfully processed file: {filepath}")
            return True

        except Exception as e:
            logger.error(f"Failed to process file {filepath}: {e}")
            return False
