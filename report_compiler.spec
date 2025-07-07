# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Report Compiler - Windows Single File Build

This creates a single executable file that includes all dependencies.
Run with: pyinstaller report_compiler.spec
"""

import sys
import os
from pathlib import Path

# Get the project root directory
project_root = os.path.dirname(os.path.abspath(SPEC))

# Define paths
main_script = os.path.join(project_root, 'main.py')
icon_path = os.path.join(project_root, 'word_integration', 'icons', 'compile-report.ico')

# Data files to include
datas = [
    # Include the entire report_compiler package
    (os.path.join(project_root, 'report_compiler'), 'report_compiler'),
    # Include Word integration files if they exist
    (os.path.join(project_root, 'word_integration'), 'word_integration'),
]

# Hidden imports - modules that PyInstaller might miss
hiddenimports = [
    # Core modules
    'report_compiler',
    'report_compiler.core',
    'report_compiler.core.compiler',
    'report_compiler.core.config',
    'report_compiler.document',
    'report_compiler.document.docx_processor',
    'report_compiler.document.placeholder_parser',
    'report_compiler.document.word_converter',
    'report_compiler.document.libreoffice_converter',
    'report_compiler.pdf',
    'report_compiler.pdf.content_analyzer',
    'report_compiler.pdf.marker_remover',
    'report_compiler.pdf.merge_processor',
    'report_compiler.pdf.overlay_processor',
    'report_compiler.utils',
    'report_compiler.utils.conversions',
    'report_compiler.utils.file_manager',
    'report_compiler.utils.logging_config',
    'report_compiler.utils.page_selector',
    'report_compiler.utils.pdf_to_svg',
    'report_compiler.utils.validators',
    
    # Windows-specific modules
    'win32com',
    'win32com.client',
    'win32com.client.gencache',
    'comtypes',
    'comtypes.client',
    
    # PDF processing
    'fitz',
    'pymupdf',
    
    # DOCX processing
    'docx',
    'python_docx',
    'docx.shared',
    'docx.enum',
    'docx.enum.text',
    
    # Standard library modules that might be missed
    'argparse',
    'pathlib',
    'shutil',
    'tempfile',
    'subprocess',
    'time',
    'datetime',
    'logging',
    'logging.handlers',
    're',
    'json',
    'zipfile',
    'xml',
    'xml.etree',
    'xml.etree.ElementTree',
    
    # Additional dependencies
    'tqdm',
]

# Binaries to exclude (reduce size)
excludes = [
    'tkinter',
    'matplotlib',
    'numpy',
    'pandas',
    'scipy',
    'PIL',
    'unittest',
    'test',
    'tests',
]

a = Analysis(
    [main_script],
    pathex=[project_root],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

# Remove duplicate entries
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='report-compiler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Keep console for CLI tool
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if os.path.exists(icon_path) else None,
    version_file=None,
)
