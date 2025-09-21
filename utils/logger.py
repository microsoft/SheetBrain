# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

"""Logging utilities for SheetBrain."""

import logging
import sys
from typing import Optional


def setup_logger(
    name: str = "sheetbrain",
    level: int = logging.INFO,
    format_string: Optional[str] = None
) -> logging.Logger:
    """
    Set up a logger with consistent formatting.

    Args:
        name: Logger name
        level: Logging level
        format_string: Custom format string

    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)

    if logger.handlers:
        return logger

    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"

    formatter = logging.Formatter(format_string)

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(formatter)

    logger.addHandler(handler)
    logger.setLevel(level)

    return logger