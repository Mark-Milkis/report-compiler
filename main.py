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

# Import from the modular refactored system
from report_compiler.core.compiler import ReportCompiler
from report_compiler.core.config import Config
from report_compiler.utils.logging_config import setup_logging, get_logger


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
  [[INSERT: path/file.docx]]        - Recursively compile and insert a DOCX file

Features:
  • Recursive compilation of DOCX files
  • Content-aware cropping with border preservation
  • Multi-page overlay support with automatic table replication
  • Comprehensive validation and error reporting
        """)

    parser.add_argument('input_file', help='Input DOCX file path')
    parser.add_argument('output_file', help='Output PDF file path')
    parser.add_argument('--keep-temp', action='store_true', help='Keep temporary files for debugging')
    parser.add_argument('--verbose', '-v', '--debug', action='store_true', help='Enable verbose logging (DEBUG level)')
    parser.add_argument('--log-file', help='Log to file in addition to console')
    parser.add_argument('--version', action='version', version=f'Report Compiler v{Config.__version__ if hasattr(Config, "__version__") else "2.0.0"}')

    # Parse arguments
    args = parser.parse_args()

    # Setup logging
    setup_logging(log_file=args.log_file, verbose=args.verbose)

    logger = get_logger()
    logger.info("=" * 60)
    logger.info("Report Compiler v2.0 - Starting compilation")
    logger.info("=" * 60)
    
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {args.input_file}")
        sys.exit(1)
    
    if not input_path.suffix.lower() == '.docx':
        logger.error(f"Input file must be a DOCX document: {args.input_file}")
        sys.exit(1)
    
    logger.info(f"Input DOCX: {input_path.absolute()}")
    
    # Validate output directory
    output_path = Path(args.output_file)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output PDF: {output_path.absolute()}")
        logger.debug(f"Output directory created/verified: {output_path.parent}")
    except Exception as e:
        logger.error(f"Cannot create output directory: {e}", exc_info=True)
        sys.exit(1)
    
    # Run the report compiler
    compiler = None
    try:
        compiler = ReportCompiler(
            input_path=str(input_path.absolute()),
            output_path=str(output_path.absolute()),
            keep_temp=args.keep_temp
        )
        
        success = compiler.run()
        
        if success:
            logger.info("=" * 60)
            logger.info("🎉 Report compilation completed successfully!")
            logger.info(f"📄 Output: {output_path.absolute()}")
            logger.info("=" * 60)
            sys.exit(0)
        else:
            logger.error("=" * 60)
            logger.error("❌ Report compilation failed!")
            logger.error("=" * 60)
            sys.exit(1)
            
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Report compilation interrupted by user.")
        sys.exit(1)
    except Exception as e:
        logger.error(f"\n❌ An unexpected error occurred during compilation: {e}", exc_info=True)
        sys.exit(1)
    finally:
        if compiler and hasattr(compiler, 'word_converter'):
            compiler.word_converter.disconnect()


if __name__ == '__main__':
    main()
