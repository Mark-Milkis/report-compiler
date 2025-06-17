#!/usr/bin/env python3
"""Minimal test to debug compiler import."""

import sys
import traceback

try:
    print("Testing step-by-step import...")
    
    # Test basic imports
    print("1. Testing basic imports...")
    import os
    import time
    from typing import Dict, List, Any, Optional
    from pathlib import Path
    print("✓ Basic imports successful")
    
    # Test relative imports one by one
    print("2. Testing relative imports...")
    
    print("   - FileManager...")
    from report_compiler.utils.file_manager import FileManager
    print("   ✓ FileManager imported")
    
    print("   - Validators...")
    from report_compiler.utils.validators import Validators
    print("   ✓ Validators imported")
    
    print("   - PlaceholderParser...")
    from report_compiler.document.placeholder_parser import PlaceholderParser
    print("   ✓ PlaceholderParser imported")
    
    print("   - DocxProcessor...")
    from report_compiler.document.docx_processor import DocxProcessor
    print("   ✓ DocxProcessor imported")
    
    print("   - WordConverter...")
    from report_compiler.document.word_converter import WordConverter
    print("   ✓ WordConverter imported")
    
    print("   - ContentAnalyzer...")
    from report_compiler.pdf.content_analyzer import ContentAnalyzer
    print("   ✓ ContentAnalyzer imported")
    
    print("   - OverlayProcessor...")
    from report_compiler.pdf.overlay_processor import OverlayProcessor
    print("   ✓ OverlayProcessor imported")
    
    print("   - MergeProcessor...")
    from report_compiler.pdf.merge_processor import MergeProcessor
    print("   ✓ MergeProcessor imported")
    
    print("   - MarkerRemover...")
    from report_compiler.pdf.marker_remover import MarkerRemover
    print("   ✓ MarkerRemover imported")
    
    print("   - Config...")
    from report_compiler.core.config import Config
    print("   ✓ Config imported")
    
    print("3. All imports successful, now creating class...")
    
    # Now manually create the class here to test
    class TestReportCompiler:
        def __init__(self, input_path: str, output_path: str, keep_temp: bool = False):
            self.input_path = input_path
            self.output_path = output_path
            self.keep_temp = keep_temp
            
            # Initialize components
            self.file_manager = FileManager(keep_temp)
            self.validators = Validators()
            self.placeholder_parser = PlaceholderParser()
            self.content_analyzer = ContentAnalyzer()
            self.docx_processor = DocxProcessor(input_path)
            self.word_converter = WordConverter()
            self.overlay_processor = OverlayProcessor()
            self.merge_processor = MergeProcessor()
            self.marker_remover = MarkerRemover()
            
            print("✓ ReportCompiler created successfully")
    
    print("4. Testing class creation...")
    test_compiler = TestReportCompiler("test.docx", "test.pdf")
    print("✓ Class creation successful")
    
except Exception as e:
    print(f"❌ Error: {e}")
    traceback.print_exc()
