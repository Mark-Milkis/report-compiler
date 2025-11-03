#!/usr/bin/env python3
"""
Report Compiler - CLI logic.
"""

import sys
import os
from pathlib import Path
import typer

from report_compiler.core.compiler import ReportCompiler
from report_compiler.core.config import Config
from report_compiler.utils.logging_config import setup_logging, get_logger
from report_compiler.utils.pdf_to_svg import PdfToSvgConverter
from report_compiler.utils.word_integration_manager import WordIntegrationManager
from report_compiler._version import __version__

app = typer.Typer(
    help=f"""
Report Compiler v{__version__} - Compile DOCX documents with embedded PDF placeholders

Examples:
  report-compiler report.docx final_report.pdf
  report-compiler report.docx output.pdf --keep-temp
  report-compiler svg-import input.pdf output.svg --page 3
  report-compiler word-integration install

Placeholder Types:
  [[OVERLAY: path/file.pdf]]        - Table-based overlay (precise positioning)
  [[OVERLAY: path/file.pdf, crop=false]]  - Overlay without content cropping
  [[IMAGE: path/image.png]]         - Direct image insertion into tables
  [[IMAGE: image.jpg, width=2in]]   - Image with size parameters
  [[INSERT: path/file.pdf]]         - Paragraph-based merge (full document)
  [[INSERT: path/file.pdf:1-3,7]]   - Insert specific pages only
  [[INSERT: path/file.docx]]        - Recursively compile and insert a DOCX file

Word Integration Commands:
  word-integration install          - Install Word template for ribbon buttons
  word-integration remove           - Remove Word template  
  word-integration update           - Update Word template to latest version
  word-integration status           - Show Word integration status

Features:
  • Recursive compilation of DOCX files
  • Content-aware cropping with border preservation
  • Multi-page overlay support with automatic table replication
  • High-quality PDF to SVG conversion for single or multiple pages
  • Comprehensive validation and error reporting
  • Automated Word integration management via uvx
    """
)

def version_callback(value: bool):
    if value:
        typer.echo(f"Report Compiler v{__version__}")
        raise typer.Exit()

@app.command("compile")
def compile_docx(
    input_file: str = typer.Argument(..., help="Input DOCX file path"),
    output_file: str = typer.Argument(..., help="Output PDF file path"),
    keep_temp: bool = typer.Option(False, help="Keep temporary files for debugging"),
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console"),
    version: bool = typer.Option(False, "--version", callback=version_callback, is_eager=True, help="Show version and exit")
):
    """Compile DOCX to PDF."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    logger.info("=" * 60)
    logger.info(f"Report Compiler v{__version__} - Starting compilation")
    logger.info("=" * 60)
    return handle_compilation(input_file, output_file, keep_temp, logger)

@app.command("svg-import")
def svg_import(
    input_file: str = typer.Argument(..., help="Input PDF file path"),
    output_file: str = typer.Argument(..., help="Output SVG file path"),
    page: str = typer.Option("all", help="Page(s) to convert: single number, range (1-3), list (1,3,5), or 'all'"),
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console"),
    version: bool = typer.Option(False, "--version", callback=version_callback, is_eager=True, help="Show version and exit")
):
    """Convert PDF page(s) to SVG format."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    logger.info("=" * 60)
    logger.info(f"Report Compiler v{__version__} - Starting PDF to SVG conversion")
    logger.info("=" * 60)
    return handle_svg_import(input_file, output_file, page, logger)

# Create a subcommand app for word-integration commands
word_app = typer.Typer(
    help="Manage Word integration template installation and updates",
    name="word-integration"
)

@word_app.command("install")
def install_word_template(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Install Word integration template to startup folder."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    
    logger.info("=" * 60)
    logger.info("Installing Word Integration Template")
    logger.info("=" * 60)
    
    manager = WordIntegrationManager()
    success, message = manager.install_template()
    
    if success:
        logger.info(f"✅ {message}")
        logger.info("")
        logger.info("Next steps:")
        logger.info("1. Restart Microsoft Word")
        logger.info("2. Look for 'Report Compiler' ribbon buttons")
        logger.info("3. Use the buttons to insert placeholders and compile reports")
        raise typer.Exit(0)
    else:
        logger.error(f"❌ {message}")
        raise typer.Exit(1)

@word_app.command("remove")
def remove_word_template(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Remove Word integration template from startup folder."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    
    logger.info("=" * 60)
    logger.info("Removing Word Integration Template")
    logger.info("=" * 60)
    
    manager = WordIntegrationManager()
    success, message = manager.remove_template()
    
    if success:
        logger.info(f"✅ {message}")
        logger.info("")
        logger.info("The Word integration has been removed.")
        logger.info("Restart Microsoft Word to complete the removal.")
        raise typer.Exit(0)
    else:
        logger.error(f"❌ {message}")
        raise typer.Exit(1)

@word_app.command("update")
def update_word_template(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Update Word integration template to latest version."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    
    logger.info("=" * 60)
    logger.info("Updating Word Integration Template")
    logger.info("=" * 60)
    
    manager = WordIntegrationManager()
    success, message = manager.update_template()
    
    if success:
        logger.info(f"✅ {message}")
        logger.info("")
        logger.info("The Word integration has been updated.")
        logger.info("Restart Microsoft Word to use the updated template.")
        raise typer.Exit(0)
    else:
        logger.error(f"❌ {message}")
        raise typer.Exit(1)

@word_app.command("status")
def word_integration_detailed_status(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Show detailed Word integration status."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    
    manager = WordIntegrationManager()
    status = manager.get_status()
    
    logger.info("=" * 60)
    logger.info("Word Integration Detailed Status")
    logger.info("=" * 60)
    logger.info(f"Platform: {status['platform']}")
    logger.info(f"Platform Supported: {'Yes' if status['supported'] else 'No'}")
    
    if status['supported']:
        logger.info(f"Word Startup Folder: {status['startup_folder']}")
        logger.info(f"Startup Folder Exists: {'Yes' if status['startup_folder'] and Path(status['startup_folder']).exists() else 'No'}")
        logger.info(f"Template Installed: {'Yes' if status['template_installed'] else 'No'}")
        if status['template_installed']:
            logger.info(f"Installed Template Path: {status['template_path']}")
    else:
        logger.info("Word integration is not supported on this platform.")
        logger.info("Supported platforms: Windows, macOS")
    
    logger.info(f"Source Template Available: {'Yes' if status['source_template_exists'] else 'No'}")
    logger.info(f"Source Template Path: {status['source_template_path']}")
    
    if status['supported'] and not status['template_installed']:
        logger.info("")
        logger.info("💡 To install Word integration, run:")
        logger.info("   uvx report-compiler word-integration install")
    elif status['supported'] and status['template_installed']:
        logger.info("")
        logger.info("💡 To update Word integration, run:")
        logger.info("   uvx report-compiler word-integration update")
        logger.info("💡 To remove Word integration, run:")
        logger.info("   uvx report-compiler word-integration remove")
    
    logger.info("=" * 60)
    
    raise typer.Exit(0 if status['supported'] else 1)

from report_compiler import interactive_menu

@app.command("interactive")
def interactive_mode():
    """Start an interactive shell session."""
    interactive_menu.main()

# Add the word-integration subcommand app to the main app
app.add_typer(word_app, name="word-integration")

def main():
    if len(sys.argv) == 1:
        interactive_menu.main()
    else:
        app()

def handle_svg_import(input_file, output_file, page, logger) -> int:
    """Handle PDF to SVG conversion."""
    logger.info("Mode: PDF to SVG conversion")
    
    # Validate input file
    input_path = Path(input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {input_file}")
        return 1
    
    if not input_path.suffix.lower() == '.pdf':
        logger.error(f"Input file must be a PDF document: {input_file}")
        return 1
    
    logger.info(f"Input PDF: {input_path.absolute()}")
    
    # Validate output file
    output_path = Path(output_file)
    if not output_path.suffix.lower() == '.svg':
        logger.error(f"Output file must have .svg extension: {output_file}")
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
        pages_to_convert = parse_page_range(page, validation_result['page_count'])
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

def handle_compilation(input_file, output_file, keep_temp, logger) -> int:
    """Handle the traditional DOCX compilation."""
    logger.info("Mode: DOCX compilation")
    
    # Validate input file
    input_path = Path(input_file)
    if not input_path.exists():
        logger.error(f"Input file not found: {input_file}")
        return 1
    
    if not input_path.suffix.lower() == '.docx':
        logger.error(f"Input file must be a DOCX document: {input_file}")
        return 1
    
    logger.info(f"Input DOCX: {input_path.absolute()}")
    
    # Validate output directory
    output_path = Path(output_file)
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
            keep_temp=keep_temp
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

if __name__ == "__main__":
    main()
