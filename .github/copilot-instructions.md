# Report Compiler - Development Instructions

Always reference these instructions first and fallback to search or bash commands only when you encounter unexpected information that does not match the info here.

## Project Overview

Report Compiler is a Python-based tool for engineering teams to create professional PDF reports by embedding PDF content directly into Word documents. The system uses Word automation (Windows) or LibreOffice (cross-platform) to convert DOCX to PDF, then uses PyMuPDF to overlay/merge additional PDF content based on placeholders.

**Architecture**: Modular pipeline-based system: DOCX → Placeholder Detection → Document Modification → PDF Conversion → PDF Processing → Final PDF

## Working Effectively

### Prerequisites and Environment Setup

- **Python 3.7+** (tested with 3.12.3)
- **Windows recommended** (for Word automation), Linux compatible (LibreOffice fallback)
- **Microsoft Word** (Windows) or **LibreOffice** (cross-platform)

### Bootstrap and Installation

```bash
# Clone and setup (first time)
git clone <repository-url>
cd report-compiler

# Method 1: Development setup (recommended for coding)
pip install -e .

# Method 2: Using uv (fastest for users)
pip install uv
uvx install report-compiler
# or direct run: uvx report-compiler@latest compile input.docx output.pdf

# Method 3: From PyPI
pip install report-compiler
```

### Build and Package

```bash
# Install uv if not available
pip install uv

# Build packages - NEVER CANCEL: Takes 3 seconds, set timeout to 60+ seconds
uv build

# Outputs:
# dist/report_compiler-VERSION-py3-none-any.whl
# dist/report_compiler-VERSION.tar.gz
```

### Testing the Installation

```bash
# Verify CLI works
report-compiler --help
report-compiler compile --help
report-compiler svg-import --help

# Test Python import
python -c "import report_compiler; print('Import successful')"

# Test with PYTHONPATH (development)
PYTHONPATH=src python -m report_compiler.cli --help
```

## CLI Commands and Usage

### Basic Compilation

```bash
# Standard compilation
report-compiler compile input.docx output.pdf

# Debug mode (keeps temporary files)
report-compiler compile input.docx output.pdf --keep-temp --verbose

# With custom log file
report-compiler compile input.docx output.pdf --log-file compilation.log
```

### PDF to SVG Conversion

```bash
# Convert specific page
report-compiler svg-import input.pdf output.svg --page 3

# Convert page range
report-compiler svg-import input.pdf output.svg --page 1-5

# Convert multiple pages
report-compiler svg-import input.pdf output.svg --page 1,3,5

# Convert all pages
report-compiler svg-import input.pdf output.svg --page all
```

## Dependencies and Limitations

### Core Dependencies

- **comtypes>=1.2.1** - COM automation interface
- **Pillow>=10.2.0** - Image processing
- **python-docx>=1.2.0** - DOCX document manipulation
- **PyMuPDF>=1.26.3** - PDF processing and manipulation
- **typer>=0.9.0** - CLI framework
- **pywin32** - Windows COM automation (Windows only)

### Platform Limitations

- **Word Automation**: Works only on Windows with Microsoft Word installed
- **LibreOffice Alternative**: Cross-platform but with slightly different formatting
- **PDF Processing**: Works on all platforms via PyMuPDF
- **Word Integration (VBA)**: Windows-only Word add-in features

## Validation and Testing

### Manual Testing Workflow

Since there are no automated tests, always manually validate changes:

1. **Build and install** your changes
2. **Create test documents** with placeholders:
   - Simple DOCX with `[[INSERT: test.pdf]]`
   - Table with `[[OVERLAY: test.pdf, page=1]]`
3. **Test compilation** with debug mode:
   ```bash
   report-compiler compile test.docx output.pdf --keep-temp --verbose
   ```
4. **Verify outputs** by opening generated PDFs
5. **Test error conditions**:
   - Missing PDF files
   - Invalid page selections
   - Corrupted input files

### Expected Timing

- **Build time**: ~3 seconds (uv build)
- **Installation**: 5-15 seconds (depending on method)
- **Simple compilation**: 5-30 seconds (depends on document complexity)
- **Word automation**: Can be slow (30+ seconds for complex documents)

## Word Integration

### Installation (Windows Only)

```bash
# Copy Word template to startup folder
# Windows: %APPDATA%\Microsoft\Word\STARTUP\
cp word_integration/ReportCompilerTemplate.dotm "%APPDATA%\Microsoft\Word\STARTUP\"

# Restart Word to see "Report Compiler" ribbon buttons
```

### Available Buttons

- **Insert Appendix** - Adds INSERT placeholder with file browser
- **Insert Overlay** - Adds OVERLAY placeholder in table with options
- **PDF to SVG** - Converts PDF pages to SVG for direct insertion
- **Compile Report** - Runs compiler directly from Word

## Troubleshooting Common Issues

### Installation Problems

- **"Document must be saved first"**: Save Word document before compilation
- **"PDF file not found"**: Check paths are relative to Word document location
- **"Word automation failed"**: Ensure Word is installed and not blocked by other instances
- **Network timeouts during pip install**: Use `--timeout 600` or try uv instead

### Development Issues

- **Import errors**: Use `PYTHONPATH=src python -m report_compiler.cli` for development
- **COM errors (Windows)**: Close all Word instances before testing
- **PDF processing errors**: Ensure PyMuPDF is properly installed
- **Missing dependencies**: Run `pip install -e .` to install all requirements

### Debug Mode

Always use debug mode when troubleshooting:

```bash
report-compiler compile input.docx output.pdf --keep-temp --verbose
```

This preserves temporary files and shows detailed processing steps.

## Codebase Navigation

### Key Source Locations

- **Entry point**: `src/report_compiler/cli.py` - Main CLI interface
- **Core orchestration**: `src/report_compiler/core/compiler.py` - Main ReportCompiler class
- **Document processing**: `src/report_compiler/document/` - DOCX handling, Word automation
- **PDF operations**: `src/report_compiler/pdf/` - PDF overlay, merge, content analysis
- **Utilities**: `src/report_compiler/utils/` - File management, validation, page selection
- **Word integration**: `word_integration/` - VBA macros and Word template
- **Configuration**: `src/report_compiler/core/config.py` - Constants and regex patterns

### Processing Pipeline (12 Stages)

1. **Initialization** - Setup logging and file management
2. **Path Validation** - Validate input/output paths
3. **Copy Input** - Copy DOCX to temp location
4. **Find Placeholders** - Scan for INSERT/OVERLAY patterns
5. **Recursive DOCX Resolution** - Handle nested DOCX files
6. **Validate Placeholders** - Check PDF paths and page selections
7. **DOCX Modification** - Insert markers, replicate tables
8. **PDF Conversion** - Convert DOCX to base PDF
9. **Content Analysis** - Find markers in PDF, analyze structure
10. **Overlay Processing** - Process table-based overlays
11. **Merge Processing** - Process paragraph-based inserts
12. **Finalization** - Cleanup and validation

### Making Changes

- **Always test compilation** after changes to core modules
- **Check both Windows and LibreOffice** conversion paths when modifying document processing
- **Validate PDF operations** with various PDF types and page selections
- **Test recursive DOCX compilation** when modifying placeholder detection
- **Update documentation** if changing CLI interface or placeholder syntax

## Common Development Tasks

### Adding New Placeholder Types

1. Update regex patterns in `Config` class
2. Extend `PlaceholderParser` to detect new patterns
3. Update `DocxProcessor` for new marker types
4. Create new processor in `pdf/` module
5. Update `ReportCompiler` pipeline

### Debugging Processing Issues

1. Use `--keep-temp --verbose` flags
2. Examine intermediate files in temp directory
3. Check marker placement in base PDF
4. Verify coordinate calculations and overlay positioning
5. Test with minimal test cases

### Performance Optimization

- **Memory usage**: Large PDFs consume significant memory
- **Processing time**: Word automation is typically slowest step
- **File I/O**: Minimize temporary file operations
- **PDF operations**: PyMuPDF performance scales with document complexity

Always measure timing changes and document performance implications.