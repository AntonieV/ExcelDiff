import logging
import sys
from logging.handlers import RotatingFileHandler


def init_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.WARNING)
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s',
                                  '%d-%m-%Y %H:%M:%S')

    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setLevel(logging.DEBUG)
    stdout_handler.setFormatter(formatter)

    file_handler = RotatingFileHandler('logs.log', maxBytes=500 * 1024, backupCount=1)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stdout_handler)
    return logging.getLogger()
