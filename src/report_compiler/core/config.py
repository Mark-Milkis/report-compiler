"""
Configuration and constants for the report compiler.
"""

import re
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
