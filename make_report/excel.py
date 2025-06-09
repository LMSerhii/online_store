from typing import Optional

import pandas as pd
from logger import logger
from status import StatusService
from enum import Enum
from typing import List
from dataclasses import dataclass


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


class ColumnAlignment(Enum):
    CENTER = 'center'
    LEFT = 'left'
    RIGHT = 'right'


@dataclass
class Column:
    letter: str
    alignment: ColumnAlignment = ColumnAlignment.LEFT
    width: Optional[int] = None


class ExcelConfig:
    COLUMNS = {
        'A': Column('A', ColumnAlignment.CENTER),
        'D': Column('D', ColumnAlignment.CENTER),
        'G': Column('G', ColumnAlignment.CENTER),
        'H': Column('H', ColumnAlignment.CENTER),
        'I': Column('I', ColumnAlignment.CENTER),
        'J': Column('J', ColumnAlignment.CENTER),
        'K': Column('K', ColumnAlignment.CENTER),
        'L': Column('L', ColumnAlignment.CENTER),
        'M': Column('M', ColumnAlignment.CENTER),
        'N': Column('N', ColumnAlignment.CENTER),
        'O': Column('O', ColumnAlignment.CENTER),
        'R': Column('R', ColumnAlignment.CENTER),
        'S': Column('S', ColumnAlignment.CENTER),
    }

    @classmethod
    def get_columns_by_alignment(cls, alignment: ColumnAlignment) -> List[str]:
        return [col.letter for col in cls.COLUMNS.values() if col.alignment == alignment]
