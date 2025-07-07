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
from report_compiler.utils.pdf_to_svg import PdfToSvgConverter


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description='Report Compiler v2.0 - Compile DOCX documents with embedded PDF placeholders',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s report.docx final_report.pdf
  %(prog)s report.docx output.pdf --keep-temp
  %(prog)s --action svg_import input.pdf output.svg --page 3

Placeholder Types:
  [[OVERLAY: path/file.pdf]]        - Table-based overlay (precise positioning)
  [[OVERLAY: path/file.pdf, crop=false]]  - Overlay without content cropping
  [[INSERT: path/file.pdf]]         - Paragraph-based merge (full document)
  [[INSERT: path/file.pdf:1-3,7]]   - Insert specific pages only
  [[INSERT: path/file.docx]]        - Recursively compile and insert a DOCX file

Actions:
  compile    - (Default) Compile DOCX documents with embedded PDF placeholders
  svg_import - Convert single or multiple PDF pages to high-quality SVG format

Features:
  • Recursive compilation of DOCX files
  • Content-aware cropping with border preservation
  • Multi-page overlay support with automatic table replication
  • High-quality PDF to SVG conversion for single or multiple pages
  • Comprehensive validation and error reporting
        """)

    parser.add_argument('input_file', help='Input DOCX file path (or PDF file path when using --action svg_import)')
    parser.add_argument('output_file', help='Output PDF file path (or SVG file path when using --action svg_import)')
    parser.add_argument('--keep-temp', action='store_true', help='Keep temporary files for debugging')
    parser.add_argument('--verbose', '-v', '--debug', action='store_true', help='Enable verbose logging (DEBUG level)')
    parser.add_argument('--log-file', help='Log to file in addition to console')
    parser.add_argument('--action', choices=['compile', 'svg_import'], default='compile', help='Action to perform: compile DOCX to PDF or convert PDF page to SVG')
    parser.add_argument('--page', type=str, default='all', help='Page(s) to convert: single number, range (1-3), list (1,3,5), or "all" (for svg_import action)')
    parser.add_argument('--version', action='version', version=f'Report Compiler v{Config.__version__ if hasattr(Config, "__version__") else "2.0.0"}')

    # Parse arguments
    args = parser.parse_args()

    # Setup logging
    setup_logging(log_file=args.log_file, verbose=args.verbose)

    logger = get_logger()
    logger.info("=" * 60)
    logger.info("Report Compiler v2.0 - Starting compilation")
    logger.info("=" * 60)
    
    # Handle different actions
    if args.action == 'svg_import':
        return handle_svg_import(args, logger)
    else:
        return handle_compilation(args, logger)


def handle_svg_import(args, logger) -> int:
    """Handle PDF to SVG conversion."""
    logger.info("Mode: PDF to SVG conversion")
    
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {args.input_file}")
        return 1
    
    if not input_path.suffix.lower() == '.pdf':
        logger.error(f"Input file must be a PDF document: {args.input_file}")
        return 1
    
    logger.info(f"Input PDF: {input_path.absolute()}")
    
    # Validate output file
    output_path = Path(args.output_file)
    if not output_path.suffix.lower() == '.svg':
        logger.error(f"Output file must have .svg extension: {args.output_file}")
        return 1
    
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output SVG: {output_path.absolute()}")
        logger.debug(f"Output directory created/verified: {output_path.parent}")
    except Exception as e:
        logger.error(f"Cannot create output directory: {e}", exc_info=True)
        return 1
    
    # Initialize converter and validate PDF
    converter = PdfToSvgConverter()
    validation_result = converter.validate_pdf(str(input_path.absolute()))
    
    if not validation_result['valid']:
        logger.error(f"PDF validation failed: {validation_result['error']}")
        return 1
    
    logger.info(f"PDF is valid with {validation_result['page_count']} pages")
    
    # Parse page specification
    try:
        pages_to_convert = parse_page_range(args.page, validation_result['page_count'])
    except ValueError as e:
        logger.error(f"Invalid page specification: {e}")
        return 1
    
    logger.info(f"Converting {len(pages_to_convert)} page(s): {pages_to_convert}")
    
    # Handle multiple pages
    if len(pages_to_convert) == 1:
        # Single page - use the original output path
        page_num = pages_to_convert[0]
        logger.info(f"Converting page {page_num} to SVG...")
        
        success = converter.convert_page_to_svg(
            pdf_path=str(input_path.absolute()),
            page_number=page_num,
            output_svg_path=str(output_path.absolute())
        )
        
        if success:
            logger.info("=" * 60)
            logger.info("🎉 PDF to SVG conversion completed successfully!")
            logger.info(f"📄 Output: {output_path.absolute()}")
            logger.info("=" * 60)
            return 0
        else:
            logger.error("=" * 60)
            logger.error("❌ PDF to SVG conversion failed!")
            logger.error("=" * 60)
            return 1
    else:
        # Multiple pages - create numbered files
        output_dir = output_path.parent
        output_stem = output_path.stem
        
        successful_conversions = 0
        
        for page_num in pages_to_convert:
            # Create filename like "output_page_1.svg", "output_page_2.svg", etc.
            page_output_path = output_dir / f"{output_stem}_page_{page_num}.svg"
            
            logger.info(f"Converting page {page_num} to {page_output_path.name}...")
            
            success = converter.convert_page_to_svg(
                pdf_path=str(input_path.absolute()),
                page_number=page_num,
                output_svg_path=str(page_output_path.absolute())
            )
            
            if success:
                successful_conversions += 1
            else:
                logger.error(f"Failed to convert page {page_num}")
        
        if successful_conversions == len(pages_to_convert):
            logger.info("=" * 60)
            logger.info("🎉 All PDF pages converted successfully!")
            logger.info(f"📄 {successful_conversions} SVG files created in: {output_dir.absolute()}")
            logger.info("=" * 60)
            return 0
        elif successful_conversions > 0:
            logger.warning("=" * 60)
            logger.warning(f"⚠️ Partial success: {successful_conversions}/{len(pages_to_convert)} pages converted")
            logger.warning(f"📄 {successful_conversions} SVG files created in: {output_dir.absolute()}")
            logger.warning("=" * 60)
            return 1
        else:
            logger.error("=" * 60)
            logger.error("❌ All PDF to SVG conversions failed!")
            logger.error("=" * 60)
            return 1


def handle_compilation(args, logger) -> int:
    """Handle the traditional DOCX compilation."""
    logger.info("Mode: DOCX compilation")
    
    # Validate input file
    input_path = Path(args.input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {args.input_file}")
        return 1
    
    if not input_path.suffix.lower() == '.docx':
        logger.error(f"Input file must be a DOCX document: {args.input_file}")
        return 1
    
    logger.info(f"Input DOCX: {input_path.absolute()}")
    
    # Validate output directory
    output_path = Path(args.output_file)
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output PDF: {output_path.absolute()}")
        logger.debug(f"Output directory created/verified: {output_path.parent}")
    except Exception as e:
        logger.error(f"Cannot create output directory: {e}", exc_info=True)
        return 1
    
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
            return 0
        else:
            logger.error("=" * 60)
            logger.error("❌ Report compilation failed!")
            logger.error("=" * 60)
            return 1
            
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Report compilation interrupted by user.")
        return 1
    except Exception as e:
        logger.error(f"\n❌ An unexpected error occurred during compilation: {e}", exc_info=True)
        return 1
    finally:
        if compiler and hasattr(compiler, 'word_converter'):
            compiler.word_converter.disconnect()

def parse_page_range(page_spec: str, total_pages: int) -> list:
    """
    Parse page specification into a list of page numbers.
    
    Args:
        page_spec: Page specification string (e.g., "1", "1-3", "1,3,5", "all")
        total_pages: Total number of pages in the PDF
        
    Returns:
        List of page numbers (1-based indexing)
        
    Raises:
        ValueError: If page specification is invalid
    """
    page_spec = page_spec.strip().lower()
    
    if page_spec == "all":
        return list(range(1, total_pages + 1))
    
    pages = []
    
    # Split by commas to handle lists like "1,3,5"
    for part in page_spec.split(','):
        part = part.strip()
        
        if '-' in part:
            # Handle ranges like "1-3"
            try:
                start, end = part.split('-', 1)
                start = int(start.strip())
                end = int(end.strip())
                
                if start < 1 or end < 1 or start > total_pages or end > total_pages:
                    raise ValueError(f"Page range {start}-{end} is out of bounds (1-{total_pages})")
                if start > end:
                    raise ValueError(f"Invalid range {start}-{end}: start page must be <= end page")
                
                pages.extend(range(start, end + 1))
            except ValueError as e:
                if "invalid literal" in str(e):
                    raise ValueError(f"Invalid page range format: {part}")
                raise
        else:
            # Handle single page numbers
            try:
                page_num = int(part)
                if page_num < 1 or page_num > total_pages:
                    raise ValueError(f"Page {page_num} is out of bounds (1-{total_pages})")
                pages.append(page_num)
            except ValueError as e:
                if "invalid literal" in str(e):
                    raise ValueError(f"Invalid page number: {part}")
                raise
    
    # Remove duplicates and sort
    return sorted(list(set(pages)))

if __name__ == '__main__':
    main()
