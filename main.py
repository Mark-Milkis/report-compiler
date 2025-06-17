#!/usr/bin/env python3
"""
Report Compiler - Main CLI entry point.

A Python-based DOCX+PDF report compiler for engineering teams.
"""

import argparse
import sys
import os
from pathlib import Path

# Add the package to the path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from report_compiler import ReportCompiler, Config


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Report Compiler v2.0 - Compile DOCX documents with embedded PDF placeholders',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s report.docx final_report.pdf
  %(prog)s report.docx output.pdf --keep-temp

Placeholder Types:
  [[OVERLAY: path/file.pdf]]        - Table-based overlay (precise positioning)
  [[OVERLAY: path/file.pdf, crop=false]]  - Overlay without content cropping
  [[INSERT: path/file.pdf]]         - Paragraph-based merge (full document)
  [[INSERT: path/file.pdf:1-3,7]]   - Insert specific pages only

Features:
  • Content-aware cropping with border preservation
  • Multi-page overlay support with automatic table replication
  • Advanced marker removal without artifacts
  • Annotation baking to preserve PDF markup
  • Comprehensive validation and error reporting
        """)
    
    parser.add_argument('input_file', 
                       help='Input DOCX file path')
    parser.add_argument('output_file', 
                       help='Output PDF file path')
    parser.add_argument('--keep-temp', 
                       action='store_true',
                       help='Keep temporary files for debugging')
    parser.add_argument('--version', 
                       action='version', 
                       version=f'Report Compiler v{Config.__version__ if hasattr(Config, "__version__") else "2.0.0"}')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"❌ Input file not found: {args.input_file}")
        sys.exit(1)
    
    if not input_path.suffix.lower() == '.docx':
        print(f"❌ Input file must be a DOCX document: {args.input_file}")
        sys.exit(1)
    
    # Validate output directory
    output_path = Path(args.output_file)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        print(f"❌ Cannot create output directory: {e}")
        sys.exit(1)
    
    # Run the report compiler
    try:
        compiler = ReportCompiler(
            input_path=str(input_path.absolute()),
            output_path=str(output_path.absolute()),
            keep_temp=args.keep_temp
        )
        
        success = compiler.run()
        
        if success:
            print(f"\\n🎉 Report compilation completed successfully!")
            print(f"📄 Output: {output_path.absolute()}")
            sys.exit(0)
        else:
            print(f"\\n❌ Report compilation failed!")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print(f"\\n⚠️ Report compilation interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\\n❌ Unexpected error: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
