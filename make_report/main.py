from config import AppConfig, PromExportConfig
from excel import ExcelProcessor
from export import ExportProm
from logger import logger
from status import StatusService


def main():
    try:
        # Ініціалізація конфігурації
        config = AppConfig.from_env()
        export_config = PromExportConfig.from_env()
        logger.info("Configuration loaded successfully")

        # Ініціалізація сервісів
        status_service = StatusService(config)
        logger.info("Status service initialized successfully")

        excel_processor = ExcelProcessor(status_service)
        logger.info("Excel processor initialized successfully")

        # Виконання експорту
        exporter = ExportProm(export_config, status_service)
        data = exporter.get_data()

        if data:
            # Експорт даних в Excel
            exporter.export_to_excel(data, "export_results.xlsx")
            logger.info("Export completed successfully")
        else:
            logger.error("No data to export")

    except Exception as e:
        logger.error(f"An error occurred: {e}")
        raise


if __name__ == '__main__':
    main()
