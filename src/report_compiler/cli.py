#!/usr/bin/env python3
"""
Report Compiler - CLI logic.
"""

import sys
from pathlib import Path
import typer

from report_compiler.core.compiler import ReportCompiler
from report_compiler.utils.logging_config import setup_logging, get_logger
from report_compiler.utils.progress import ProgressReporter
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
    no_progress: bool = typer.Option(False, "--no-progress", help="Disable the live progress indicator"),
    temp_dir: str = typer.Option(None, "--temp-dir", help="Directory for temporary files (default: OS temp folder). Avoids OneDrive/SharePoint sync issues."),
    cache_dir: str = typer.Option(None, "--cache-dir", help="Directory for the compiled-document cache (default: under OS temp folder)."),
    no_cache: bool = typer.Option(False, "--no-cache", help="Disable reusing/storing compiled sub-document PDFs across runs."),
    version: bool = typer.Option(False, "--version", callback=version_callback, is_eager=True, help="Show version and exit")
):
    """Compile DOCX to PDF."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    logger.info("=" * 60)
    logger.info(f"Report Compiler v{__version__} - Starting compilation")
    logger.info("=" * 60)
    return handle_compilation(
        input_file, output_file, keep_temp, logger,
        show_progress=not no_progress,
        temp_dir=temp_dir, cache_dir=cache_dir, use_cache=not no_cache,
    )

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

    built, build_msg = _ensure_template_built(logger)
    if not built:
        logger.error(f"❌ {build_msg}")
        raise typer.Exit(1)

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

    built, build_msg = _ensure_template_built(logger, force=True)
    if not built:
        logger.error(f"❌ {build_msg}")
        raise typer.Exit(1)

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

def _template_sources():
    """Resolve the tracked template-source paths (only present in a source checkout)."""
    manager = WordIntegrationManager()
    output_dotm = manager.get_template_source_path()
    root = output_dotm.parent
    return {
        "root": root,
        "skeleton": root / "skeleton",
        "customui": root / "report_compiler_UI.xml",
        "icons": root / "icons",
        "vbaproject_bin": root / "vbaProject.bin",
        # Static empty macro-enabled template, used only as the carrier to compile the
        # .bas modules into. Shipped with the package so install can build from a PyPI
        # install (the real ReportCompilerTemplate.dotm is a build artifact).
        "seed": root / "seed.dotm",
        "output_dotm": output_dotm,
    }


def _ensure_template_built(logger, force: bool = False) -> tuple:
    """Build ReportCompilerTemplate.dotm from the shipped sources if it's missing.

    Compiles the .bas into the static seed.dotm carrier (needs Word), then assembles
    the final .dotm from skeleton + ribbon + icons + the compiled VBA. Returns
    (success, message).
    """
    from report_compiler.utils.template_builder import build_vba_bin
    from report_compiler.utils.template_packager import package_template

    src = _template_sources()
    if src["output_dotm"].exists() and not force:
        return True, "Template already built."
    if not src["seed"].exists():
        return False, (
            f"Seed template not found: {src['seed']}. The package is missing seed.dotm "
            "(needed to compile the Word macros)."
        )
    logger.info("  > Building Word template (compiling macros + packaging)...")
    ok, msg = build_vba_bin(src["root"], src["seed"], src["vbaproject_bin"], logger)
    if not ok:
        return False, msg
    return package_template(
        src["skeleton"], src["customui"], src["icons"], src["vbaproject_bin"],
        src["output_dotm"], logger,
    )


@word_app.command("build-vba")
def build_word_vba(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Compile the .bas modules into vbaProject.bin (maintenance; needs Word)."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler.utils.template_builder import build_vba_bin

    logger.info("=" * 60)
    logger.info("Compiling VBA modules -> vbaProject.bin")
    logger.info("=" * 60)
    src = _template_sources()
    success, message = build_vba_bin(src["root"], src["seed"], src["vbaproject_bin"], logger)
    if success:
        logger.info(f"✅ {message}")
        raise typer.Exit(0)
    logger.error(f"❌ {message}")
    raise typer.Exit(1)


@word_app.command("package")
def package_word_template(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Assemble the .dotm from tracked sources (ribbon + icons + vbaProject.bin). No Word."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler.utils.template_packager import package_template

    logger.info("=" * 60)
    logger.info("Packaging .dotm from sources")
    logger.info("=" * 60)
    src = _template_sources()
    success, message = package_template(
        src["skeleton"], src["customui"], src["icons"], src["vbaproject_bin"],
        src["output_dotm"], logger,
    )
    if success:
        logger.info(f"✅ {message}")
        logger.info("")
        logger.info("Install/refresh it in Word with: word-integration update")
        raise typer.Exit(0)
    logger.error(f"❌ {message}")
    raise typer.Exit(1)


@word_app.command("build-template")
def build_word_template(
    skip_vba: bool = typer.Option(False, "--skip-vba", help="Reuse existing vbaProject.bin; only re-package"),
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Full rebuild: compile VBA (needs Word) then package the .dotm from sources."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler.utils.template_builder import build_vba_bin
    from report_compiler.utils.template_packager import package_template

    logger.info("=" * 60)
    logger.info("Rebuilding Word Template (VBA + package)")
    logger.info("=" * 60)
    src = _template_sources()

    if not skip_vba:
        ok, msg = build_vba_bin(src["root"], src["seed"], src["vbaproject_bin"], logger)
        if not ok:
            logger.error(f"❌ {msg}")
            raise typer.Exit(1)
        logger.info(f"✅ {msg}")

    ok, msg = package_template(
        src["skeleton"], src["customui"], src["icons"], src["vbaproject_bin"],
        src["output_dotm"], logger,
    )
    if ok:
        logger.info(f"✅ {msg}")
        logger.info("")
        logger.info("Install/refresh it in Word with: word-integration update")
        raise typer.Exit(0)
    logger.error(f"❌ {msg}")
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

# Create a subcommand app for COM server commands
com_app = typer.Typer(
    help="Manage the Report Compiler COM server (Word integration)",
    name="com-server"
)


@com_app.command("register")
def com_server_register(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Register the COM server (per-user, no admin) so Word can drive compilation."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler import com_server

    logger.info("=" * 60)
    logger.info("Registering Report Compiler COM Server")
    logger.info("=" * 60)
    try:
        command = com_server.bootstrap_register()
        logger.info(f"✅ Registered '{com_server.PROGID}'")
        logger.info(f"   LocalServer32: {command}")
        logger.info("")
        logger.info("Word can now use 'CreateObject(\"ReportCompiler.Application\")'.")
    except typer.Exit:
        raise
    except Exception as e:
        logger.error(f"❌ Registration failed: {e}", exc_info=verbose)
        raise typer.Exit(1)
    raise typer.Exit(0)


@com_app.command("unregister")
def com_server_unregister(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Remove the per-user COM server registration."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler import com_server

    logger.info("=" * 60)
    logger.info("Unregistering Report Compiler COM Server")
    logger.info("=" * 60)
    try:
        com_server.unregister_self()
        logger.info(f"✅ Unregistered '{com_server.PROGID}'")
    except typer.Exit:
        raise
    except Exception as e:
        logger.error(f"❌ Unregistration failed: {e}", exc_info=verbose)
        raise typer.Exit(1)
    raise typer.Exit(0)


@com_app.command("status")
def com_server_status(
    verbose: bool = typer.Option(False, "-v", "--verbose", "--debug", help="Enable verbose logging (DEBUG level)"),
    log_file: str = typer.Option(None, help="Log to file in addition to console")
):
    """Show whether the COM server is registered."""
    setup_logging(log_file=log_file, verbose=verbose)
    logger = get_logger()
    from report_compiler import com_server

    logger.info("=" * 60)
    logger.info("Report Compiler COM Server Status")
    logger.info("=" * 60)
    try:
        info = com_server.status()
    except Exception as e:
        logger.error(f"❌ {e}")
        raise typer.Exit(1)

    logger.info(f"ProgID: {info['progid']}")
    logger.info(f"CLSID:  {info['clsid']}")
    logger.info(f"Registered: {'Yes' if info['registered'] else 'No'}")
    if info['registered']:
        logger.info(f"LocalServer32: {info['local_server']}")
        logger.info(f"Class spec:    {info['class_spec']}")
    else:
        logger.info("")
        logger.info("💡 To register, run: uvx report-compiler com-server register")
    logger.info("=" * 60)
    raise typer.Exit(0 if info['registered'] else 1)


@com_app.command("_register-self", hidden=True)
def com_server_register_self():
    """Internal: write the registration for the currently running interpreter.

    Invoked by 'register' from inside the stable uv-tool install so the registry
    captures a path that survives uvx's ephemeral environments.
    """
    from report_compiler import com_server

    command = com_server.register_self()
    typer.echo(f"Registered {com_server.PROGID}: {command}")
    raise typer.Exit(0)


from report_compiler import interactive_menu

@app.command("interactive")
def interactive_mode():
    """Start an interactive shell session."""
    interactive_menu.main()

# Add the word-integration subcommand app to the main app
app.add_typer(word_app, name="word-integration")
# Add the com-server subcommand app to the main app
app.add_typer(com_app, name="com-server")

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

def handle_compilation(input_file, output_file, keep_temp, logger, show_progress: bool = True,
                       temp_dir: str = None, cache_dir: str = None, use_cache: bool = True) -> int:
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
    # Ensure the output ends in .pdf. Word/LibreOffice always write a PDF and,
    # when given a name without an extension, Word silently appends ".pdf" to the
    # file on disk while the pipeline keeps tracking the extension-less name. That
    # mismatch surfaces later as a confusing "no such file" when the base PDF is
    # re-opened, so normalize it up front (mirrors the .svg check in svg-import).
    if output_path.suffix.lower() != ".pdf":
        normalized = output_path.parent / (output_path.name + ".pdf")
        logger.warning(
            f"Output path '{output_path.name}' has no .pdf extension; using '{normalized.name}' instead."
        )
        output_path = normalized
    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        logger.info(f"Output PDF: {output_path.absolute()}")
        logger.debug(f"Output directory created/verified: {output_path.parent}")
    except Exception as e:
        logger.error(f"Cannot create output directory: {e}", exc_info=True)
        return 1
    
    # Run the report compiler
    compiler = None
    # Only drive the live indicator on an interactive terminal; otherwise it
    # would emit control sequences into piped/redirected output.
    progress_enabled = show_progress and sys.stdout.isatty()
    try:
        with ProgressReporter(enabled=progress_enabled) as progress:
            compiler = ReportCompiler(
                input_path=str(input_path.absolute()),
                output_path=str(output_path.absolute()),
                keep_temp=keep_temp,
                progress=progress,
                temp_dir=temp_dir,
                cache_dir=cache_dir,
                use_cache=use_cache,
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
