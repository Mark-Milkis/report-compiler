#!/usr/bin/env python3
"""
Test script to verify IMAGE tag validation functionality.
"""

import os
import sys
import tempfile

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from report_compiler.utils.validators import Validators
from PIL import Image

def create_test_image():
    """Create a simple test image."""
    img = Image.new('RGB', (300, 200), color='green')
    test_image_path = '/tmp/test_validation_image.png'
    img.save(test_image_path)
    return test_image_path

def test_image_validation():
    """Test image path validation."""
    print("Testing image validation...")
    
    validators = Validators()
    test_image_path = create_test_image()
    base_directory = '/tmp'
    
    # Test valid image
    result = validators.validate_image_path('test_validation_image.png', base_directory)
    print(f"  Valid image: {result}")
    
    # Test non-existent image
    result = validators.validate_image_path('nonexistent.png', base_directory)
    print(f"  Non-existent image: {result}")
    
    # Test invalid extension
    result = validators.validate_image_path('test.txt', base_directory)
    print(f"  Invalid extension: {result}")

def test_placeholder_validation():
    """Test placeholder validation with IMAGE placeholders."""
    print("\nTesting placeholder validation...")
    
    validators = Validators()
    test_image_path = create_test_image()
    
    # Create test placeholders
    placeholders = [
        {
            'type': 'table',
            'subtype': 'image',
            'file_path': '/tmp/test_validation_image.png',
            'table_index': 0,
            'is_recursive_docx': False
        },
        {
            'type': 'table',
            'subtype': 'overlay',  
            'file_path': '/tmp/nonexistent.pdf',
            'table_index': 1,
            'is_recursive_docx': False
        },
        {
            'type': 'table',
            'subtype': 'image',
            'file_path': '/tmp/nonexistent_image.png',
            'table_index': 2,
            'is_recursive_docx': False
        }
    ]
    
    result = validators.validate_placeholders(placeholders, '/tmp')
    print(f"  Validation result: {result}")
    
    print(f"  Valid: {result['valid']}")
    print(f"  Errors: {result['errors']}")
    print(f"  Warnings: {result['warnings']}")
    
    # Check that valid image placeholder was processed correctly
    valid_placeholder = placeholders[0]
    if 'resolved_path' in valid_placeholder:
        print(f"  Valid image placeholder resolved path: {valid_placeholder['resolved_path']}")
        print(f"  Image dimensions: {valid_placeholder.get('image_width')}x{valid_placeholder.get('image_height')}")

if __name__ == "__main__":
    print("Testing IMAGE validation functionality...\n")
    
    test_image_validation()
    test_placeholder_validation()
    
    print("\n✓ Validation tests completed!")