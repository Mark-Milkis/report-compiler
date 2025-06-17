"""
LibreOffice automation for DOCX to PDF conversion.
"""

import os
import subprocess
from typing import Optional
from ..core.config import Config

class LibreOfficeConverter:
    """Handles DOCX to PDF conversion using headless LibreOffice."""
    @staticmethod
    def convert_to_pdf(docx_path: str, pdf_path: str) -> bool:
        """
        Convert DOCX to PDF using headless LibreOffice.
        """
        try:
            print("    Converting DOCX to PDF using LibreOffice...")
            print(f"    Input: {docx_path}")
            print(f"    Output: {pdf_path}")
            os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
            cmd = [
                Config.LIBREOFFICE_EXECUTABLE,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', os.path.dirname(pdf_path),
                docx_path
            ]
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode != 0:
                print(f"    ❌ LibreOffice conversion failed: {result.stderr.decode().strip()}")
                return False
            expected_pdf = os.path.join(os.path.dirname(pdf_path), os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
            if expected_pdf != pdf_path:
                if os.path.exists(expected_pdf):
                    os.rename(expected_pdf, pdf_path)
            print(f"    ✓ Successfully converted '{os.path.basename(docx_path)}' to PDF with LibreOffice")
            return True
        except Exception as e:
            print(f"    ❌ Error converting with LibreOffice: {e}")
            return False
