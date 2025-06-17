"""
Unit tests for the Validators utility class.
"""

import unittest
import tempfile
import os
from unittest.mock import patch, MagicMock
import fitz  # PyMuPDF

from report_compiler.utils.validators import Validators


class TestValidators(unittest.TestCase):
    """Test cases for Validators class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.validators = Validators()
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Clean up test fixtures."""
        # Clean up temp directory if it exists
        if os.path.exists(self.temp_dir):
            os.rmdir(self.temp_dir)
    
    def create_temp_file(self, filename, content="test content"):
        """Helper to create temporary files."""
        filepath = os.path.join(self.temp_dir, filename)
        with open(filepath, 'w') as f:
            f.write(content)
        return filepath
    
    def test_validate_input_file_exists(self):
        """Test input file validation when file exists."""
        test_file = self.create_temp_file("test.docx")
        
        try:
            self.validators.validate_input_file(test_file)
            # Should not raise an exception
        except Exception as e:
            self.fail(f"validate_input_file raised an exception: {e}")
        finally:
            os.remove(test_file)
    
    def test_validate_input_file_not_exists(self):
        """Test input file validation when file doesn't exist."""
        nonexistent_file = os.path.join(self.temp_dir, "nonexistent.docx")
        
        with self.assertRaises(FileNotFoundError):
            self.validators.validate_input_file(nonexistent_file)
    
    def test_validate_input_file_wrong_extension(self):
        """Test input file validation with wrong extension."""
        test_file = self.create_temp_file("test.txt")
        
        try:
            with self.assertRaises(ValueError):
                self.validators.validate_input_file(test_file)
        finally:
            os.remove(test_file)
    
    def test_validate_output_path_valid_directory(self):
        """Test output path validation with valid directory."""
        output_path = os.path.join(self.temp_dir, "output.pdf")
        
        try:
            self.validators.validate_output_path(output_path)
            # Should not raise an exception
        except Exception as e:
            self.fail(f"validate_output_path raised an exception: {e}")
    
    def test_validate_output_path_invalid_directory(self):
        """Test output path validation with invalid directory."""
        output_path = os.path.join("nonexistent_dir", "output.pdf")
        
        with self.assertRaises(ValueError):
            self.validators.validate_output_path(output_path)
    
    def test_validate_output_path_wrong_extension(self):
        """Test output path validation with wrong extension."""
        output_path = os.path.join(self.temp_dir, "output.docx")
        
        with self.assertRaises(ValueError):
            self.validators.validate_output_path(output_path)
    
    @patch('fitz.open')
    def test_validate_pdf_file_success(self, mock_fitz_open):
        """Test PDF file validation when file is valid."""
        mock_doc = MagicMock()
        mock_doc.page_count = 5
        mock_doc.__enter__ = MagicMock(return_value=mock_doc)
        mock_doc.__exit__ = MagicMock(return_value=None)
        mock_fitz_open.return_value = mock_doc
        
        result = self.validators.validate_pdf_file("test.pdf")
        self.assertEqual(result, 5)
        mock_fitz_open.assert_called_once_with("test.pdf")
    
    @patch('fitz.open')
    def test_validate_pdf_file_not_found(self, mock_fitz_open):
        """Test PDF file validation when file is not found."""
        mock_fitz_open.side_effect = FileNotFoundError("File not found")
        
        with self.assertRaises(FileNotFoundError):
            self.validators.validate_pdf_file("nonexistent.pdf")
    
    @patch('fitz.open')
    def test_validate_pdf_file_invalid_pdf(self, mock_fitz_open):
        """Test PDF file validation when file is not a valid PDF."""
        mock_fitz_open.side_effect = Exception("Invalid PDF")
        
        with self.assertRaises(ValueError):
            self.validators.validate_pdf_file("invalid.pdf")
    
    @patch('fitz.open')
    def test_validate_pdf_file_zero_pages(self, mock_fitz_open):
        """Test PDF file validation when PDF has zero pages."""
        mock_doc = MagicMock()
        mock_doc.page_count = 0
        mock_doc.__enter__ = MagicMock(return_value=mock_doc)
        mock_doc.__exit__ = MagicMock(return_value=None)
        mock_fitz_open.return_value = mock_doc
        
        with self.assertRaises(ValueError):
            self.validators.validate_pdf_file("empty.pdf")
    
    def test_validate_pdf_paths_all_valid(self):
        """Test PDF paths validation when all paths are valid."""
        pdf_paths = ["test1.pdf", "test2.pdf", "test3.pdf"]
        
        with patch.object(self.validators, 'validate_pdf_file') as mock_validate:
            mock_validate.side_effect = [3, 5, 2]  # Page counts
            
            result = self.validators.validate_pdf_paths(pdf_paths)
            expected = {
                "test1.pdf": 3,
                "test2.pdf": 5,
                "test3.pdf": 2
            }
            self.assertEqual(result, expected)
            self.assertEqual(mock_validate.call_count, 3)
    
    def test_validate_pdf_paths_some_invalid(self):
        """Test PDF paths validation when some paths are invalid."""
        pdf_paths = ["valid.pdf", "invalid.pdf"]
        
        with patch.object(self.validators, 'validate_pdf_file') as mock_validate:
            mock_validate.side_effect = [3, ValueError("Invalid PDF")]
            
            with self.assertRaises(ValueError):
                self.validators.validate_pdf_paths(pdf_paths)
    
    def test_validate_pdf_paths_empty_list(self):
        """Test PDF paths validation with empty list."""
        result = self.validators.validate_pdf_paths([])
        self.assertEqual(result, {})
    
    def test_validate_pdf_paths_duplicates(self):
        """Test PDF paths validation with duplicate paths."""
        pdf_paths = ["test.pdf", "test.pdf", "other.pdf"]
        
        with patch.object(self.validators, 'validate_pdf_file') as mock_validate:
            mock_validate.side_effect = [3, 5]  # Only called twice due to deduplication
            
            result = self.validators.validate_pdf_paths(pdf_paths)
            expected = {
                "test.pdf": 3,
                "other.pdf": 5
            }
            self.assertEqual(result, expected)
            self.assertEqual(mock_validate.call_count, 2)


if __name__ == '__main__':
    unittest.main()
