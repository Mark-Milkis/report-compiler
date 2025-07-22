"""
Report Compiler - A Python-based DOCX+PDF report compiler for engineering teams.

This package provides functionality to compile Word documents with embedded PDF placeholders
into professional PDF reports with precise overlay positioning and merged appendices.
"""


# Try to import the version from the auto-generated file by setuptools-scm
try:
    from ._version import version as __version__
except ImportError:
    # Fallback for when the package is not installed
    __version__ = "0.0.0-dev"
__author__ = "Report Compiler Team"

# from .core.compiler import ReportCompiler  # Temporarily commented
from .core.config import Config

__all__ = ['Config']  # 'ReportCompiler'
