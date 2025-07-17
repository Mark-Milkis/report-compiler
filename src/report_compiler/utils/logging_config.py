"""
Centralized logging configuration for the report compiler.
"""

import logging
import sys
from typing import Optional


class ColoredFormatter(logging.Formatter):
    """Custom formatter with color support for console output."""

    # Color codes for different log levels
    COLORS = {
        'DEBUG': '\033[36m',      # Cyan
        'INFO': '\033[32m',       # Green
        'WARNING': '\033[33m',    # Yellow
        'ERROR': '\033[31m',      # Red
        'CRITICAL': '\033[35m',   # Magenta
        'RESET': '\033[0m'        # Reset
    }

    def format(self, record):
        # Add color to the log level name
        if record.levelname in self.COLORS:
            record.levelname = f"{self.COLORS[record.levelname]}{record.levelname}{self.COLORS['RESET']}"
        return super().format(record)


class ReportCompilerLogger:
    """Centralized logger for the report compiler with structured output."""

    _instance = None
    _logger = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if self._logger is None:
            self._setup_logger()

    def _setup_logger(self, level: str = "INFO", log_file: Optional[str] = None):
        """Setup the logger with console and optionally file output."""
        self._logger = logging.getLogger("report_compiler")
        self._logger.setLevel(logging.DEBUG)  # Set logger to lowest level

        # Clear any existing handlers
        if self._logger.hasHandlers():
            self._logger.handlers.clear()

        # Create formatters based on the desired level
        log_format = '%(message)s'
        if level.upper() == 'DEBUG':
            log_format = '[%(name)s] %(message)s'

        console_formatter = ColoredFormatter(f'%(levelname)s: {log_format}')

        file_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s() - %(message)s'
        )

        # Console handler with color
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(getattr(logging, level.upper()))
        console_handler.setFormatter(console_formatter)
        self._logger.addHandler(console_handler)

        # File handler (if specified)
        if log_file:
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(file_formatter)
            self._logger.addHandler(file_handler)

        self._logger.propagate = False

    def set_level(self, level: str):
        """Set the logging level for the console handler."""
        upper_level = level.upper()
        log_level = getattr(logging, upper_level)

        for handler in self._logger.handlers:
            if isinstance(handler, logging.StreamHandler) and handler.stream == sys.stdout:
                handler.setLevel(log_level)
                log_format = '%(message)s'
                if upper_level == 'DEBUG':
                    log_format = '[%(name)s] %(message)s'
                formatter = ColoredFormatter(f'%(levelname)s: {log_format}')
                handler.setFormatter(formatter)

    def add_file_logging(self, log_file: str):
        """Add file logging to an existing logger."""
        file_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(funcName)s() - %(message)s'
        )
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(file_formatter)
        self._logger.addHandler(file_handler)

    @property
    def logger(self):
        """Get the logger instance."""
        return self._logger


# Global logger instance
_logger_instance = ReportCompilerLogger()
logger = _logger_instance.logger


def setup_logging(log_file: Optional[str] = None, verbose: bool = False):
    """
    Setup logging for the report compiler.

    Args:
        log_file: Optional file to log to
        verbose: Enable verbose logging (DEBUG level)
    """
    level = "DEBUG" if verbose else "INFO"
    _logger_instance._setup_logger(level, log_file)
    logger.debug(f"Logging level set to: {level}")
    if log_file:
        logger.debug(f"Log file enabled at: {log_file}")


def get_logger() -> logging.Logger:
    """Get the root report compiler logger."""
    return logger


# Module-specific logger factory
def get_module_logger(name: str) -> logging.Logger:
    """
    Get a module-specific logger that inherits from the main logger.

    Args:
        name: Module name (e.g., __name__)

    Returns:
        Module-specific logger instance
    """
    module_name = name.split('.')[-1] if '.' in name else name
    return logging.getLogger(f"report_compiler.{module_name}")


# Pre-configured loggers for each module
def get_compiler_logger():
    """Get logger for compiler module."""
    return get_module_logger('compiler')

def get_docx_logger():
    """Get logger for DOCX processing."""
    return get_module_logger('docx_processor')

def get_pdf_logger():
    """Get logger for PDF processing."""
    return get_module_logger('pdf_processor')

def get_merge_logger():
    """Get logger for merge processing."""
    return get_module_logger('merge_processor')

def get_overlay_logger():
    """Get logger for overlay processing."""
    return get_module_logger('overlay_processor')

def get_word_logger():
    """Get logger for Word conversion."""
    return get_module_logger('word_converter')

def get_file_logger():
    """Get logger for file operations."""
    return get_module_logger('file_manager')

def get_validation_logger():
    """Get logger for validation."""
    return get_module_logger('validators')
