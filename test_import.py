#!/usr/bin/env python3
"""Test import of individual modules."""

try:
    print("Testing individual imports...")
    
    print("1. Testing Config import...")
    from report_compiler.core.config import Config
    print("✓ Config imported successfully")
    
    print("2. Testing FileManager import...")
    from report_compiler.utils.file_manager import FileManager
    print("✓ FileManager imported successfully")
    
    print("3. Testing PlaceholderParser import...")
    from report_compiler.document.placeholder_parser import PlaceholderParser
    print("✓ PlaceholderParser imported successfully")
    
    print("4. Testing DocxProcessor import...")
    from report_compiler.document.docx_processor import DocxProcessor
    print("✓ DocxProcessor imported successfully")
    
    print("5. Testing ReportCompiler import...")
    from report_compiler.core.compiler import ReportCompiler
    print("✓ ReportCompiler imported successfully")
    
    print("All imports successful!")
    
except ImportError as e:
    print(f"❌ Import error: {e}")
except Exception as e:
    print(f"❌ Unexpected error: {e}")
