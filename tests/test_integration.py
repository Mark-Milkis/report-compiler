#!/usr/bin/env python3
"""
Integration tests for the report compiler.
These tests focus on the main functionality rather than individual unit tests.
"""

import os
import sys
import tempfile
import unittest
from pathlib import Path

# Add the project root to the Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from report_compiler.utils.file_manager import FileManager
from report_compiler.utils.page_selector import PageSelector
from report_compiler.utils.validators import Validators
from report_compiler.document.placeholder_parser import PlaceholderParser


class TestIntegration(unittest.TestCase):
    """Integration tests for key functionality."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.temp_dir = tempfile.mkdtemp()
    
    def tearDown(self):
        """Clean up test fixtures."""
        import shutil
        if os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir)
    
    def test_file_manager_basic_functionality(self):
        """Test FileManager basic operations."""
        # Test initialization
        fm = FileManager(keep_temp=False)
        self.assertFalse(fm.keep_temp)
        self.assertEqual(len(fm.temp_files), 0)
        
        # Test temp path generation
        base_path = os.path.join(self.temp_dir, "test.docx")
        temp_path = fm.generate_temp_path(base_path, "modified")
        
        self.assertIn("test_modified_", temp_path)
        self.assertTrue(temp_path.endswith(".docx"))
        self.assertIn(temp_path, fm.temp_files)
        
        # Test cleanup
        # Create a real temp file for testing
        with open(temp_path, 'w') as f:
            f.write("test content")
        
        self.assertTrue(os.path.exists(temp_path))
        fm.cleanup()
        self.assertFalse(os.path.exists(temp_path))
        self.assertEqual(len(fm.temp_files), 0)
    
    def test_file_manager_context_manager(self):
        """Test FileManager as context manager."""
        temp_file = os.path.join(self.temp_dir, "context_test.txt")
        
        with FileManager(keep_temp=False) as fm:
            path = fm.generate_temp_path(temp_file, "test")
            # Create the file
            with open(path, 'w') as f:
                f.write("test")
            self.assertTrue(os.path.exists(path))
        
        # File should be cleaned up automatically
        self.assertFalse(os.path.exists(path))
    
    def test_page_selector_parse_specification(self):
        """Test PageSelector parse functionality."""
        ps = PageSelector()
        
        # Test basic parsing
        result = ps.parse_specification("1-3")
        self.assertIn('pages', result)
        self.assertEqual(result['pages'], [1, 2, 3])
        
        # Test complex specification
        result = ps.parse_specification("1,3-5,7")
        expected_pages = [1, 3, 4, 5, 7]
        self.assertEqual(result['pages'], expected_pages)
        
        # Test empty specification
        result = ps.parse_specification("")
        self.assertEqual(result['pages'], [])
    
    def test_validators_static_methods(self):
        """Test Validators static methods."""
        # Create a test DOCX file
        test_docx = os.path.join(self.temp_dir, "test.docx")
        with open(test_docx, 'w') as f:
            f.write("test content")
        
        # Test DOCX validation
        result = Validators.validate_docx_path(test_docx)
        self.assertTrue(result['valid'])
        self.assertEqual(result['path'], test_docx)
        
        # Test invalid DOCX path
        result = Validators.validate_docx_path("nonexistent.docx")
        self.assertFalse(result['valid'])
        self.assertIn('error', result)
        
        # Test output path validation
        output_path = os.path.join(self.temp_dir, "output.pdf")
        result = Validators.validate_output_path(output_path)
        self.assertTrue(result['valid'])
        
        # Test invalid output path (wrong extension)
        result = Validators.validate_output_path(os.path.join(self.temp_dir, "output.docx"))
        self.assertFalse(result['valid'])
    
    def test_placeholder_parser_initialization(self):
        """Test PlaceholderParser initialization."""
        parser = PlaceholderParser()
        self.assertIsNotNone(parser)
        
        # Test with nonexistent file should handle gracefully
        result = parser.find_all_placeholders("nonexistent.docx")
        self.assertIn('error', result)
        self.assertFalse(result['success'])
    
    def test_file_manager_utility_methods(self):
        """Test FileManager utility static methods."""
        # Test path validation
        valid_path = FileManager.validate_path(self.temp_dir)
        self.assertIsNotNone(valid_path)
        
        # Test directory creation
        new_dir = os.path.join(self.temp_dir, "new_subdir", "file.txt")
        result = FileManager.ensure_directory_exists(new_dir)
        self.assertTrue(result)
        self.assertTrue(os.path.exists(os.path.dirname(new_dir)))
        
        # Test file size calculation
        test_file = os.path.join(self.temp_dir, "size_test.txt")
        with open(test_file, 'w') as f:
            f.write("x" * 1000)  # 1000 bytes
        
        size_mb = FileManager.get_file_size_mb(test_file)
        self.assertGreater(size_mb, 0)
        self.assertLess(size_mb, 1)  # Should be less than 1 MB
    
    def test_page_selector_validation(self):
        """Test PageSelector validation functionality."""
        ps = PageSelector()
        
        # Create a test selection
        selection = ps.parse_specification("1-5,10")
        
        # Test validation with sufficient pages
        validated = ps.validate_pages(selection, total_pages=15)
        self.assertTrue(validated['valid'])
        self.assertEqual(validated['effective_pages'], [1, 2, 3, 4, 5, 10])
        
        # Test validation with insufficient pages
        validated = ps.validate_pages(selection, total_pages=3)
        self.assertTrue(validated['valid'])  # Should still be valid but filtered
        self.assertEqual(validated['effective_pages'], [1, 2, 3])
    
    def test_component_integration(self):
        """Test that components work together."""
        # Test that FileManager and PageSelector can work together
        fm = FileManager(keep_temp=True)
        ps = PageSelector()
        
        # Generate a temp path
        base_path = os.path.join(self.temp_dir, "integration_test.pdf")
        temp_path = fm.generate_temp_path(base_path, "processed")
        
        # Parse page specification
        page_selection = ps.parse_specification("1-3")
        
        # Verify both components produced expected results
        self.assertIn("integration_test_processed_", temp_path)
        self.assertEqual(page_selection['pages'], [1, 2, 3])
        
        # Clean up
        fm.cleanup()


def main():
    """Run integration tests."""
    print("Running Report Compiler Integration Tests")
    print("=" * 50)
    
    # Run the tests
    suite = unittest.TestLoader().loadTestsFromTestCase(TestIntegration)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    
    if result.wasSuccessful():
        print("\n" + "=" * 50)
        print("✅ All integration tests passed!")
        return True
    else:
        print("\n" + "=" * 50)
        print("❌ Some integration tests failed!")
        return False


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
