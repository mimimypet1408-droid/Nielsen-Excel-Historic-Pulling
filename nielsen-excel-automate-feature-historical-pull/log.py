import logging
import os
from datetime import datetime
from typing import Optional

import config
from constants import ROOT_DIR

_initialized = False
_main_log_filepath: Optional[str] = None


def get_logs_dir():
    return os.path.join(ROOT_DIR, "logs")


def get_main_log_filepath() -> Optional[str]:
    if not _initialized or not _main_log_filepath:
        return None
    return _main_log_filepath


def setup_logging():
    global _initialized, _main_log_filepath
    if _initialized:
        return
    _initialized = True

    # Create logs directory if it doesn't exist
    log_dir = get_logs_dir()
    os.makedirs(log_dir, exist_ok=True)

    # Get timestamp for log files
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    # Configure root logger
    root_logger = logging.getLogger()
    default_log_level = logging.DEBUG if config.is_debug() else logging.INFO
    root_logger.setLevel(default_log_level)

    # Create file handler for main logs
    _main_log_filepath = os.path.join(log_dir, f"main_{timestamp}.log")
    file_handler = logging.FileHandler(_main_log_filepath)
    file_handler.setLevel(default_log_level)

    # Create formatter and add it to the handlers
    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
    file_handler.setFormatter(formatter)

    # Add file handler to root logger
    root_logger.addHandler(file_handler)

    # If DEBUG is enabled, also log to console
    if config.is_debug():
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
    else:
        # Remove default stream handler
        for handler in list(root_logger.handlers):
            if isinstance(handler, logging.StreamHandler):
                root_logger.removeHandler(handler)


setup_logging()
