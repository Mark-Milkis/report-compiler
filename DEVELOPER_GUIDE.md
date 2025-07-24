# Report Compiler - Developer Guide

This document provides technical details about the Report Compiler's architecture, implementation, and development workflow for contributors and developers who want to understand or extend the codebase.

## Architecture Overview

The Report Compiler uses a modular, pipeline-based architecture with clear separation of concerns:

```text
Input DOCX → Placeholder Detection → Document Modification → PDF Conversion → PDF Processing → Final PDF
     ↓              ↓                     ↓                     ↓              ↓              ↓
User Document → Parse Tags → Insert Markers → Word Automation → Overlay/Merge → Clean Output
```

### Core Modules

#### `report_compiler.core`
Main orchestration and configuration management.

- **`ReportCompiler`** - Primary orchestrator that manages the entire compilation pipeline
- **`Config`** - Central configuration class containing constants, regex patterns, and settings

#### `report_compiler.document`
Word document processing and manipulation.

- **`PlaceholderParser`** - Detects and parses PDF placeholders using regex patterns
- **`DocxProcessor`** - Modifies DOCX files (inserts markers, replicates table cells, handles page breaks)
- **`WordConverter`** - Converts DOCX to PDF using Word automation (win32com)
- **`LibreOfficeConverter`** - Alternative converter using LibreOffice (cross-platform fallback)

#### `report_compiler.pdf`
PDF processing and manipulation using PyMuPDF.

- **`ContentAnalyzer`** - Analyzes PDF structure, finds markers, handles annotation baking
- **`OverlayProcessor`** - Handles table-based PDF overlays with precise positioning
- **`MergeProcessor`** - Handles paragraph-based PDF insertions and TOC management
- **`MarkerRemover`** - Removes placement markers from final PDF using redaction

#### `report_compiler.utils`
Shared utilities and helper functions.

- **`FileManager`** - Temporary file management and cleanup operations
- **`Validators`** - Input validation and PDF verification utilities
- **`PageSelector`** - Page selection parsing and processing logic
- **`PdfToSvgConverter`** - High-quality PDF to SVG conversion utility

## Processing Pipeline

### 1. Initialization and Validation

```python
compiler = ReportCompiler(input_path, output_path, keep_temp, recursion_level)
compiler.run()
```

The `ReportCompiler` performs initial setup:
- Validates input/output paths
- Initializes file manager for temporary file handling
- Sets up logging with appropriate prefixes for nested compilation

### 2. Placeholder Detection

The `PlaceholderParser` scans the DOCX document structure:

```python
placeholders = parser.find_all_placeholders(docx_path)
# Returns: {'table': [...], 'paragraph': [...], 'total': int}
```

**Table Placeholders (OVERLAY):**
- Searches through `document.tables`
- Identifies single-cell tables containing `[[OVERLAY:...]]` patterns
- Extracts table metadata (dimensions, position) for precise overlay calculation

**Paragraph Placeholders (INSERT):**
- Searches through `document.paragraphs`
- Identifies standalone paragraphs containing `[[INSERT:...]]` patterns
- Records paragraph position for marker placement

**Recursive DOCX Handling:**
- Detects `[[INSERT: file.docx]]` patterns
- Triggers recursive compilation of nested DOCX files
- Manages circular dependency detection

### 3. Path Resolution and Validation

The `Validators` class resolves and validates all referenced files:

```python
validation_result = validators.validate_placeholders(placeholders, base_directory)
```

- Resolves relative paths based on the Word document's location
- Validates PDF file existence and readability
- Counts pages and validates page selection syntax
- Reports detailed error information for troubleshooting

### 4. Document Modification

The `DocxProcessor` creates a modified version of the input document:

```python
table_metadata = processor.create_modified_docx(input_path, placeholders, output_path)
```

**For Table Placeholders:**
- Replaces placeholder text with visible red markers (`%%OVERLAY_START_N%%`)
- Replicates table rows for multi-page PDF selections
- Extracts precise table dimensions (width, height) in points
- Returns metadata dictionary mapping table indices to dimensions

**For Paragraph Placeholders:**
- Replaces placeholder text with merge markers (`%%MERGE_START_N%%`)
- Inserts page breaks for proper PDF pagination
- Preserves original paragraph formatting

### 5. PDF Conversion

Two conversion engines are available:

**Word Automation (Default):**
```python
with WordConverter() as converter:
    success = converter.update_fields_and_save_as_pdf(docx_path, pdf_path)
```

**LibreOffice (Alternative):**
```python
converter = LibreOfficeConverter()
success = converter.convert_to_pdf(docx_path, pdf_path)
```

The Word converter provides better formatting fidelity but requires Windows and installed Word. LibreOffice provides cross-platform compatibility with slightly different output formatting.

### 6. PDF Content Analysis

The `ContentAnalyzer` processes the base PDF:

```python
content_map, toc_pages = analyzer.analyze(pdf_path, placeholders, table_metadata)
```

**Marker Detection:**
- Searches for all marker text in the PDF using text extraction
- Records exact position coordinates for each marker
- Maps markers to their corresponding placeholder metadata

**Annotation Processing:**
- Automatically detects and bakes PDF annotations into content
- Ensures annotations are preserved during overlay operations
- Uses `doc.bake()` for permanent annotation integration

**TOC Detection:**
- Identifies Table of Contents pages based on content patterns
- Used for intelligent TOC merging in INSERT operations

### 7. PDF Processing

**Overlay Processing (Table-based):**
```python
overlay_processor.process_overlays(base_pdf_path, content_map, output_path)
```

- Calculates precise overlay rectangles using marker position + table dimensions
- Applies content-aware cropping to source PDFs (unless `crop=false`)
- Overlays selected pages sequentially for multi-page selections
- Removes markers using redaction (white fill)

**Merge Processing (Paragraph-based):**
```python
merge_processor.process_merges(base_pdf_path, content_map, toc_pages, output_path)
```

- Inserts complete PDF pages at marker positions
- Maintains document flow and pagination
- Merges Table of Contents entries hierarchically
- Handles page numbering and reference updates

### 8. Finalization

- Removes all temporary files (unless `--keep-temp` specified)
- Validates final PDF integrity
- Reports success/failure with detailed logging

## Key Technical Details

### Coordinate Systems

The system uses multiple coordinate systems:

**Word (EMU):** English Metric Units (1 point = 12700 EMU)
**PyMuPDF (Points):** PostScript points (72 points = 1 inch)
**Overlay Calculations:** Points for consistency

Conversion utilities handle transformations:
```python
from report_compiler.utils.conversions import emu_to_points, points_to_inches
```

### Multi-page Handling

**Table Overlays:**
- Replicates table rows automatically for each selected page
- Creates unique markers for each cell (`%%OVERLAY_START_00_PAGE_02%%`)
- Maintains consistent table formatting across replicated cells

**Paragraph Inserts:**
- Inserts pages sequentially at marker position
- Preserves original document pagination
- Updates page references in TOC if present

### Error Handling

The system uses comprehensive error handling with structured logging:

```python
logger = get_module_logger(__name__)
logger.error(f"Specific error description: {details}", exc_info=True)
```

**Error Categories:**
- **Validation Errors:** File not found, invalid paths, corrupted PDFs
- **Processing Errors:** Word automation failures, PDF corruption
- **Logic Errors:** Invalid page selections, circular dependencies

### Debugging Support

**Debug Mode (`--keep-temp`):**
- Retains all temporary files with timestamped names
- Enables detailed step-by-step inspection
- Preserves intermediate PDFs for troubleshooting

**Verbose Logging (`--verbose`):**
- Shows detailed processing steps with timing information
- Reports coordinate calculations and dimension measurements
- Displays PDF processing statistics (pages, annotations, markers)

## Development Setup

### Prerequisites

- Python 3.7+
- Microsoft Word (for Windows development)
- LibreOffice (for cross-platform testing)

### Installation

```bash
git clone <repository-url>
cd report-compiler
pip install -e .
```

### Code Style

The project follows PEP 8 with some specific conventions:

- Use type hints for public methods
- Comprehensive docstrings for all classes and public methods
- Structured logging with appropriate log levels
- Error handling with specific exception types

### Adding New Features

#### Adding a New Placeholder Type

1. **Extend regex patterns** in `Config`:
   ```python
   NEW_REGEX = re.compile(r"\[\[NEWTYPE:\s*(.+?)\s*\]\]", re.IGNORECASE)
   ```

2. **Update PlaceholderParser** to detect new patterns
3. **Extend DocxProcessor** to handle new marker types
4. **Create new processor** in `pdf/` module for specific handling
5. **Update ReportCompiler** pipeline to include new processing step

#### Adding New Conversion Engines

1. **Create new converter** class in `document/` module
2. **Implement standard interface** (`convert_to_pdf` method)
3. **Update Config** to include new engine option
4. **Add engine selection logic** in ReportCompiler

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with appropriate tests
4. Ensure all tests pass
5. Submit a pull request with detailed description

### Commit Guidelines

- Use descriptive commit messages
- Include test changes with feature changes
- Update documentation for user-facing changes
- Follow semantic versioning for releases

## Security Considerations

- **File Access:** Validate all file paths to prevent directory traversal
- **Command Injection:** Sanitize inputs to LibreOffice subprocess calls
- **Temporary Files:** Ensure proper cleanup to prevent information leakage
- **Word Automation:** Handle COM object lifecycle properly to prevent crashes

## Performance Considerations

- **Memory Usage:** Large PDFs can consume significant memory during processing
- **Processing Time:** Word automation can be slow for complex documents
- **File I/O:** Minimize temporary file creation and cleanup overhead
- **PDF Operations:** PyMuPDF operations scale with document complexity