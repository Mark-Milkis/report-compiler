"""
Report Compiler - A Python-based DOCX+PDF report compiler for engineering teams.

This package provides functionality to compile Word documents with embedded PDF placeholders
into professional PDF reports with precise overlay positioning and merged appendices.
"""

__version__ = "2.0.0"
__author__ = "Report Compiler Team"

from .core.compiler import ReportCompiler
from .core.config import Config

__all__ = ['ReportCompiler', 'Config']
