# -*- coding: utf-8 -*-
import logging
import logging.handlers
import os

from config import CONFIG_DIR

LOG_DIR = os.path.join(CONFIG_DIR, "logs")
LOG_PATH = os.path.join(LOG_DIR, "tagflow.log")


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("tagflow")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = logging.handlers.RotatingFileHandler(
            LOG_PATH, maxBytes=2 * 1024 * 1024, backupCount=2, encoding="utf-8"
        )
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(handler)
    return logger
