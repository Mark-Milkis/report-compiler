"""
Configuration and constants for the report compiler.
"""

import os
import re
import tempfile
from typing import Optional


class Config:
    """Configuration class containing all constants and settings for the report compiler."""
    
    # Regex patterns for placeholder detection
    OVERLAY_REGEX = re.compile(r"\[\[OVERLAY:\s*([^,\]]+?)(?:,\s*(.+?))?\s*\]\]", re.IGNORECASE)
    INSERT_REGEX = re.compile(r"\[\[INSERT:\s*(.+?)(?::([^:\\\/\]]+))?\s*\]\]", re.IGNORECASE)
    IMAGE_REGEX = re.compile(r"\[\[IMAGE:\s*([^,\]]+?)(?:,\s*(.+?))?\s*\]\]", re.IGNORECASE)
    
    # Marker patterns for PDF processing
    OVERLAY_MARKER_PREFIX = "%%OVERLAY_START_"
    MERGE_MARKER_PREFIX = "%%MERGE_START_"
    PAGE_MARKER_SUFFIX = "_PAGE_"
    
    # PDF processing defaults
    DEFAULT_PADDING = 32  # points
    DEFAULT_CROP_ENABLED = False

    # Marker stored in the AltText of in-document overlay-preview images so they can be
    # found and stripped again (by the live toggle and the compile-time normalizer).
    OVERLAY_PREVIEW_MARKER = "RCPREVIEW"
    
    # File handling
    TEMP_FILE_PREFIX = "~temp_"

    # Temporary working files are kept in the OS temp directory by default rather
    # than alongside the source document. Storing them next to the working files
    # is problematic on cloud-synced folders (OneDrive/SharePoint), where sync
    # locks and Files-On-Demand placeholders cause intermittent locking and
    # "file not found" errors during conversion. Both locations can be overridden
    # via CLI flags or these environment variables.
    TEMP_DIR_ENV = "REPORT_COMPILER_TEMP_DIR"
    CACHE_DIR_ENV = "REPORT_COMPILER_CACHE_DIR"
    # Sub-directory name used under the temp base for a single compile run.
    RUN_DIR_PREFIX = "report_compiler_run_"
    # Persistent cache of compiled sub-document PDFs, so a re-run (e.g. after a
    # late-stage failure) can reuse appendices that already compiled instead of
    # re-invoking Word on each one.
    CACHE_SUBDIR = "report_compiler_cache"
    # Cached artifacts untouched for longer than this are pruned on startup.
    CACHE_TTL_DAYS = 7
    SUPPORTED_PDF_EXTENSIONS = ['.pdf']
    SUPPORTED_DOCX_EXTENSIONS = ['.docx']
    SUPPORTED_IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp', '.heic', '.heif' , '.emf', '.wmf']
    
    # Word automation settings
    WORD_EXPORT_FORMAT = 17  # PDF format in Word
    
    # Rendering engine selection: 'word' or 'libreoffice'
    DOCX_RENDER_ENGINE = 'word'  # Options: 'word', 'libreoffice'
    LIBREOFFICE_EXECUTABLE = 'libreoffice'  # Path to LibreOffice executable
    
    # Logging settings
    LOG_ICONS = {
        'search': '🔍',
        'table': '📋', 
        'paragraph': '📄',
        'success': '✅',
        'warning': '⚠️',
        'error': '❌',
        'processing': '🔧',
        'overlay': '📌',
        'merge': '📥',
        'fire': '🔥',
        'target': '🎯',
        'dimensions': '📐',
        'note': '📝',
        'position': '📍',
        'ruler': '📏',
        'package': '📦'
    }
    
    @classmethod
    def get_overlay_marker(cls, table_index: int, page_num: Optional[int] = None) -> str:
        """Generate overlay marker string."""
        if page_num is None:
            return f"{cls.OVERLAY_MARKER_PREFIX}{table_index:02d}%%"
        else:
            return f"{cls.OVERLAY_MARKER_PREFIX}{table_index:02d}{cls.PAGE_MARKER_SUFFIX}{page_num:02d}%%"
    
    @classmethod
    def get_merge_marker(cls, merge_index: int) -> str:
        """Generate merge marker string."""
        return f"{cls.MERGE_MARKER_PREFIX}{merge_index}%%"
    
    @classmethod
    def get_temp_filename(cls, base_name: str, timestamp: int) -> str:
        """Generate temporary filename."""
        return f"{cls.TEMP_FILE_PREFIX}{base_name}_{timestamp}"

    @classmethod
    def get_temp_base_dir(cls, override: Optional[str] = None) -> str:
        """Resolve the base directory for per-run temporary files.

        Precedence: explicit override (CLI flag) > environment variable >
        OS temp directory.
        """
        return override or os.environ.get(cls.TEMP_DIR_ENV) or tempfile.gettempdir()

    @classmethod
    def get_cache_dir(cls, override: Optional[str] = None) -> str:
        """Resolve the directory for the persistent compiled-document cache.

        Precedence: explicit override (CLI flag) > environment variable >
        a stable sub-directory of the OS temp directory.
        """
        return (
            override
            or os.environ.get(cls.CACHE_DIR_ENV)
            or os.path.join(tempfile.gettempdir(), cls.CACHE_SUBDIR)
        )
