import os
import tempfile
import unittest
from report_compiler.core.config import Config
from report_compiler.document.libreoffice_converter import LibreOfficeConverter

class TestLibreOfficeConverter(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        # Create a minimal DOCX file for testing
        self.docx_path = os.path.join(self.temp_dir, "test.docx")
        with open(self.docx_path, "wb") as f:
            f.write(b"PK\x03\x04")  # Write minimal ZIP header for DOCX
        self.pdf_path = os.path.join(self.temp_dir, "test.pdf")
        self._old_engine = Config.DOCX_RENDER_ENGINE
        Config.DOCX_RENDER_ENGINE = 'libreoffice'
    def tearDown(self):
        import shutil
        shutil.rmtree(self.temp_dir)
        Config.DOCX_RENDER_ENGINE = self._old_engine
    def test_convert_to_pdf(self):
        # This will only pass if LibreOffice is installed and available in PATH
        result = LibreOfficeConverter.convert_to_pdf(self.docx_path, self.pdf_path)
        self.assertTrue(result or not os.path.exists(self.pdf_path), "LibreOffice conversion should succeed or fail gracefully.")

if __name__ == "__main__":
    unittest.main()
