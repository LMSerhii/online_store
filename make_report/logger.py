import logging

import colorlog


def setup_logger():
    handler = colorlog.StreamHandler()
    handler.setFormatter(colorlog.ColoredFormatter(
        "%(log_color)s%(levelname)s:%(name)s:%(message)s",
        log_colors={
            'DEBUG': 'cyan',
            'INFO': 'green',
            'WARNING': 'yellow',
            'ERROR': 'red',
            'CRITICAL': 'bold_red',
        }
    ))

    logger = colorlog.getLogger('Prom.ua report status')
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)

    logging.getLogger().handlers.clear()

    return logger


logger = setup_logger()
