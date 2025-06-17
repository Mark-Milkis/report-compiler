"""
Unit tests for the PageSelector utility class.
"""

import unittest
from report_compiler.utils.page_selector import PageSelector


class TestPageSelector(unittest.TestCase):
    """Test cases for PageSelector class."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.page_selector = PageSelector()
    
    def test_parse_single_page(self):
        """Test parsing single page specification."""
        result = self.page_selector.parse_page_spec("5")
        self.assertEqual(result, [5])
    
    def test_parse_page_range(self):
        """Test parsing page range specification."""
        result = self.page_selector.parse_page_spec("2-5")
        self.assertEqual(result, [2, 3, 4, 5])
    
    def test_parse_open_ended_range(self):
        """Test parsing open-ended range specification."""
        result = self.page_selector.parse_page_spec("3-", max_pages=7)
        self.assertEqual(result, [3, 4, 5, 6, 7])
    
    def test_parse_comma_separated_pages(self):
        """Test parsing comma-separated page specification."""
        result = self.page_selector.parse_page_spec("1,3,5")
        self.assertEqual(result, [1, 3, 5])
    
    def test_parse_mixed_specification(self):
        """Test parsing mixed page specification."""
        result = self.page_selector.parse_page_spec("1-3,7,9-11")
        self.assertEqual(result, [1, 2, 3, 7, 9, 10, 11])
    
    def test_parse_with_spaces(self):
        """Test parsing page specification with spaces."""
        result = self.page_selector.parse_page_spec("1, 3-5, 8")
        self.assertEqual(result, [1, 3, 4, 5, 8])
    
    def test_parse_empty_string(self):
        """Test parsing empty page specification."""
        result = self.page_selector.parse_page_spec("")
        self.assertEqual(result, [])
    
    def test_parse_invalid_format(self):
        """Test parsing invalid page specification."""
        with self.assertRaises(ValueError):
            self.page_selector.parse_page_spec("invalid")
    
    def test_parse_negative_pages(self):
        """Test parsing page specification with negative numbers."""
        with self.assertRaises(ValueError):
            self.page_selector.parse_page_spec("-1")
    
    def test_parse_zero_page(self):
        """Test parsing page specification with zero."""
        with self.assertRaises(ValueError):
            self.page_selector.parse_page_spec("0")
    
    def test_filter_valid_pages_basic(self):
        """Test filtering valid pages with basic input."""
        pages = [1, 2, 3, 15, 20]
        result = self.page_selector.filter_valid_pages(pages, max_pages=10)
        self.assertEqual(result, [1, 2, 3])
    
    def test_filter_valid_pages_all_valid(self):
        """Test filtering when all pages are valid."""
        pages = [1, 3, 5]
        result = self.page_selector.filter_valid_pages(pages, max_pages=10)
        self.assertEqual(result, [1, 3, 5])
    
    def test_filter_valid_pages_all_invalid(self):
        """Test filtering when all pages are invalid."""
        pages = [15, 20, 25]
        result = self.page_selector.filter_valid_pages(pages, max_pages=10)
        self.assertEqual(result, [])
    
    def test_filter_valid_pages_empty_input(self):
        """Test filtering with empty page list."""
        result = self.page_selector.filter_valid_pages([], max_pages=10)
        self.assertEqual(result, [])
    
    def test_filter_valid_pages_duplicates(self):
        """Test filtering removes duplicates and sorts."""
        pages = [3, 1, 5, 3, 1]
        result = self.page_selector.filter_valid_pages(pages, max_pages=10)
        self.assertEqual(result, [1, 3, 5])
    
    def test_get_effective_pages_with_spec(self):
        """Test getting effective pages with page specification."""
        result = self.page_selector.get_effective_pages("2-4,7", total_pages=10)
        self.assertEqual(result, [2, 3, 4, 7])
    
    def test_get_effective_pages_without_spec(self):
        """Test getting effective pages without page specification."""
        result = self.page_selector.get_effective_pages(None, total_pages=5)
        self.assertEqual(result, [1, 2, 3, 4, 5])
    
    def test_get_effective_pages_empty_spec(self):
        """Test getting effective pages with empty specification."""
        result = self.page_selector.get_effective_pages("", total_pages=3)
        self.assertEqual(result, [1, 2, 3])
    
    def test_get_effective_pages_with_filtering(self):
        """Test getting effective pages with invalid pages filtered."""
        result = self.page_selector.get_effective_pages("2-15", total_pages=5)
        self.assertEqual(result, [2, 3, 4, 5])
    
    def test_range_parsing_edge_cases(self):
        """Test edge cases in range parsing."""
        # Single number treated as range
        result = self.page_selector.parse_page_spec("5-5")
        self.assertEqual(result, [5])
        
        # Invalid range order
        with self.assertRaises(ValueError):
            self.page_selector.parse_page_spec("5-2")
    
    def test_complex_mixed_specification(self):
        """Test complex mixed page specification."""
        result = self.page_selector.parse_page_spec("1,3-5,2,8-10,15-", max_pages=20)
        expected = [1, 2, 3, 4, 5, 8, 9, 10, 15, 16, 17, 18, 19, 20]
        self.assertEqual(result, expected)


if __name__ == '__main__':
    unittest.main()
