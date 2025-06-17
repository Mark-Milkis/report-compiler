#!/usr/bin/env python3
"""Test compiler dependencies."""

import sys
sys.path.insert(0, '.')

print("Testing compiler dependencies...")

try:
    from report_compiler.utils.file_manager import FileManager
    print("✓ FileManager imported")
except Exception as e:
    print(f"❌ FileManager error: {e}")

try:
    from report_compiler.utils.validators import Validators
    print("✓ Validators imported")
except Exception as e:
    print(f"❌ Validators error: {e}")

try:
    from report_compiler.document.placeholder_parser import PlaceholderParser
    print("✓ PlaceholderParser imported")
except Exception as e:
    print(f"❌ PlaceholderParser error: {e}")

try:
    from report_compiler.document.docx_processor import DocxProcessor
    print("✓ DocxProcessor imported")
except Exception as e:
    print(f"❌ DocxProcessor error: {e}")

try:
    from report_compiler.document.word_converter import WordConverter
    print("✓ WordConverter imported")
except Exception as e:
    print(f"❌ WordConverter error: {e}")

try:
    from report_compiler.pdf.content_analyzer import ContentAnalyzer
    print("✓ ContentAnalyzer imported")
except Exception as e:
    print(f"❌ ContentAnalyzer error: {e}")

try:
    from report_compiler.pdf.overlay_processor import OverlayProcessor
    print("✓ OverlayProcessor imported")
except Exception as e:
    print(f"❌ OverlayProcessor error: {e}")

try:
    from report_compiler.pdf.merge_processor import MergeProcessor
    print("✓ MergeProcessor imported")
except Exception as e:
    print(f"❌ MergeProcessor error: {e}")

try:
    from report_compiler.pdf.marker_remover import MarkerRemover
    print("✓ MarkerRemover imported")
except Exception as e:
    print(f"❌ MarkerRemover error: {e}")

try:
    from report_compiler.core.config import Config
    print("✓ Config imported")
except Exception as e:
    print(f"❌ Config error: {e}")

print("\nTesting compiler module import...")
try:
    import report_compiler.core.compiler as compiler_module
    print("✓ Compiler module imported")
    print("Available attributes:", [attr for attr in dir(compiler_module) if not attr.startswith('_')])
except Exception as e:
    print(f"❌ Compiler module error: {e}")
    import traceback
    traceback.print_exc()
