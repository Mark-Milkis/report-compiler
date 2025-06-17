"""
Unit tests for the FileManager utility class.
"""

import unittest
import tempfile
import os
from pathlib import Path
from unittest.mock import patch, MagicMock

from report_compiler.utils.file_manager import FileManager


class TestFileManager(unittest.TestCase):
    """Test cases for FileManager class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test.txt")
        with open(self.test_file, 'w') as f:
            f.write("test content")
    
    def tearDown(self):
        """Clean up test fixtures."""
        # Clean up manually created files
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        if os.path.exists(self.temp_dir):
            os.rmdir(self.temp_dir)
    
    def test_init_with_keep_temp_false(self):
        """Test FileManager initialization with keep_temp=False."""
        fm = FileManager(keep_temp=False)
        self.assertFalse(fm.keep_temp)
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_init_with_keep_temp_true(self):
        """Test FileManager initialization with keep_temp=True."""
        fm = FileManager(keep_temp=True)
        self.assertTrue(fm.keep_temp)
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_create_temp_filename(self):
        """Test temporary filename creation."""
        fm = FileManager()
        original_path = "test_document.docx"
        temp_filename = fm.create_temp_filename(original_path, "modified")
        
        # Should contain timestamp and suffix
        self.assertIn("test_document_modified_", temp_filename)
        self.assertTrue(temp_filename.endswith(".docx"))
    
    def test_create_temp_filename_with_different_extension(self):
        """Test temporary filename creation with different extension."""
        fm = FileManager()
        original_path = "test_document.docx"
        temp_filename = fm.create_temp_filename(original_path, "converted", ".pdf")
        
        # Should use new extension
        self.assertIn("test_document_converted_", temp_filename)
        self.assertTrue(temp_filename.endswith(".pdf"))
    
    def test_register_temp_file(self):
        """Test temporary file registration."""
        fm = FileManager()
        fm.register_temp_file(self.test_file)
        
        self.assertIn(self.test_file, fm.temp_files)
        self.assertEqual(len(fm.temp_files), 1)
    
    def test_cleanup_with_keep_temp_false(self):
        """Test cleanup when keep_temp is False."""
        fm = FileManager(keep_temp=False)
        fm.register_temp_file(self.test_file)
        
        # File should exist before cleanup
        self.assertTrue(os.path.exists(self.test_file))
        
        fm.cleanup()
        
        # File should be deleted after cleanup
        self.assertFalse(os.path.exists(self.test_file))
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_cleanup_with_keep_temp_true(self):
        """Test cleanup when keep_temp is True."""
        fm = FileManager(keep_temp=True)
        fm.register_temp_file(self.test_file)
        
        # File should exist before cleanup
        self.assertTrue(os.path.exists(self.test_file))
        
        fm.cleanup()
        
        # File should still exist after cleanup
        self.assertTrue(os.path.exists(self.test_file))
        # But temp_files list should be cleared
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_cleanup_nonexistent_file(self):
        """Test cleanup handles nonexistent files gracefully."""
        fm = FileManager(keep_temp=False)
        nonexistent_file = os.path.join(self.temp_dir, "nonexistent.txt")
        fm.register_temp_file(nonexistent_file)
        
        # Should not raise an exception
        fm.cleanup()
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_context_manager_cleanup_false(self):
        """Test FileManager as context manager with keep_temp=False."""
        temp_file = os.path.join(self.temp_dir, "context_test.txt")
        with open(temp_file, 'w') as f:
            f.write("test")
        
        with FileManager(keep_temp=False) as fm:
            fm.register_temp_file(temp_file)
            self.assertTrue(os.path.exists(temp_file))
        
        # File should be cleaned up automatically
        self.assertFalse(os.path.exists(temp_file))
    
    def test_context_manager_cleanup_true(self):
        """Test FileManager as context manager with keep_temp=True."""
        temp_file = os.path.join(self.temp_dir, "context_test2.txt")
        with open(temp_file, 'w') as f:
            f.write("test")
        
        with FileManager(keep_temp=True) as fm:
            fm.register_temp_file(temp_file)
            self.assertTrue(os.path.exists(temp_file))
        
        # File should still exist
        self.assertTrue(os.path.exists(temp_file))
        # Clean up manually
        os.remove(temp_file)


if __name__ == '__main__':
    unittest.main()
