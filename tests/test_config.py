"""
Test configuration and utilities for the report compiler test suite.
"""

import os
import tempfile
import unittest
from unittest.mock import MagicMock
from pathlib import Path


class TestConfig:
    """Configuration constants for tests."""
    
    # Test data directory
    TEST_DATA_DIR = Path(__file__).parent / "data"
    
    # Sample content for test files
    SAMPLE_DOCX_CONTENT = "Test document content"
    SAMPLE_PDF_CONTENT = b"%PDF-1.4 test content"
    
    # Common test placeholders
    OVERLAY_PLACEHOLDER = "[[OVERLAY: test.pdf, page=1]]"
    INSERT_PLACEHOLDER = "[[INSERT: test.pdf:1-3]]"


class TestUtils:
    """Utility functions for tests."""
    
    @staticmethod
    def create_temp_file(suffix=".txt", content="test content"):
        """Create a temporary file with given content."""
        temp_file = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        if isinstance(content, str):
            temp_file.write(content.encode())
        else:
            temp_file.write(content)
        temp_file.close()
        return temp_file.name
    
    @staticmethod
    def create_mock_docx_document():
        """Create a mock docx Document object."""
        mock_doc = MagicMock()
        mock_doc.paragraphs = []
        mock_doc.tables = []
        return mock_doc
    
    @staticmethod
    def create_mock_table(rows=1, cells_per_row=1, cell_texts=None):
        """Create a mock docx table."""
        mock_table = MagicMock()
        mock_rows = []
        
        for i in range(rows):
            mock_row = MagicMock()
            mock_cells = []
            
            for j in range(cells_per_row):
                mock_cell = MagicMock()
                if cell_texts and i < len(cell_texts) and j < len(cell_texts[i]):
                    mock_cell.text = cell_texts[i][j]
                else:
                    mock_cell.text = f"Cell {i},{j}"
                mock_cells.append(mock_cell)
            
            mock_row.cells = mock_cells
            mock_rows.append(mock_row)
        
        mock_table.rows = mock_rows
        return mock_table
    
    @staticmethod
    def create_mock_paragraph(text="Test paragraph"):
        """Create a mock docx paragraph."""
        mock_para = MagicMock()
        mock_para.text = text
        return mock_para


class BaseTestCase(unittest.TestCase):
    """Base test case with common functionality."""
    
    def setUp(self):
        """Set up common test fixtures."""
        self.temp_files = []
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Clean up test fixtures."""
        # Clean up temporary files
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except OSError:
                pass
        
        # Clean up temp directory
        try:
            if os.path.exists(self.temp_dir):
                os.rmdir(self.temp_dir)
        except OSError:
            pass
    
    def create_temp_file(self, suffix=".txt", content="test content"):
        """Create a temporary file and register it for cleanup."""
        temp_file = TestUtils.create_temp_file(suffix, content)
        self.temp_files.append(temp_file)
        return temp_file
    
    def assertFileExists(self, filepath):
        """Assert that a file exists."""
        self.assertTrue(os.path.exists(filepath), f"File does not exist: {filepath}")
    
    def assertFileNotExists(self, filepath):
        """Assert that a file does not exist."""
        self.assertFalse(os.path.exists(filepath), f"File exists but shouldn't: {filepath}")


def discover_and_run_tests():
    """Discover and run all tests in the tests directory."""
    # Discover tests
    test_dir = Path(__file__).parent
    loader = unittest.TestLoader()
    suite = loader.discover(test_dir, pattern='test_*.py')
    
    # Run tests
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    return result.wasSuccessful()


if __name__ == '__main__':
    discover_and_run_tests()
