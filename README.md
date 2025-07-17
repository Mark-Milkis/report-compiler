# Report Compiler

A Python-based automated DOCX and PDF report compiler for engineering teams. This tool allows engineers to write reports in Word, use placeholders to insert external PDFs, and compile everything into a professional PDF with a single command.

## Overview

The Report Compiler automates the creation of comprehensive PDF reports by:

1. **Finding PDF placeholders** in Word documents using two types of tags:
   - `[[OVERLAY: path/to/file.pdf, page=5]]` for table-based overlays
   - `[[INSERT: path/to/file.pdf]]` for paragraph-based insertions
2. **Modifying the Word document** to create markers and page breaks
3. **Converting to PDF** using Word automation (win32com)
4. **Processing PDF insertions** with overlays and merges using PyMuPDF

## Features

- ✅ **Two insertion types** - Table-based overlays and paragraph-based merges
- ✅ **Relative path support** - PDF paths resolved relative to the input Word document
- ✅ **Page selection support** - Specify which pages to include from source PDFs using flexible syntax
- ✅ **Multi-page PDF support** - Automatic cell replication for multi-page table overlays
- ✅ **Annotation preservation** - PDF annotations automatically baked into content during processing
- ✅ **Marker removal** - Automatic removal of placement markers from final PDF
- ✅ **Robust page breaks** - Proper page breaks for paragraph-based insertions
- ✅ **Error handling** - Comprehensive error reporting and validation
- ✅ **Debug support** - `--keep-temp` flag to retain temporary files for debugging
- ✅ **Table-based overlay** - Precise PDF placement using table dimensions and marker positioning
- ✅ **Cell replication** - Multi-page PDFs create consecutive table cells automatically
- ✅ **Intelligent positioning** - Uses table properties for automatic overlay rectangle calculation
- ✅ **Modular architecture** - Clean separation of concerns with focused classes and modules

## Architecture

The Report Compiler uses a modular architecture with clear separation of responsibilities:

### Core Modules

- **`report_compiler.core`** - Main orchestration and configuration
  - `ReportCompiler` - Main orchestrator class
  - `Config` - Configuration management and constants

- **`report_compiler.document`** - Word document processing
  - `PlaceholderParser` - Detects and parses PDF placeholders
  - `DocxProcessor` - Modifies DOCX files (markers, page breaks, cell replication)
  - `WordConverter` - Converts DOCX to PDF using Word automation

- **`report_compiler.pdf`** - PDF processing and manipulation
  - `ContentAnalyzer` - Analyzes PDF content and structure
  - `OverlayProcessor` - Handles table-based PDF overlays
  - `MergeProcessor` - Handles paragraph-based PDF merges
  - `MarkerRemover` - Removes placement markers from final PDF

- **`report_compiler.utils`** - Utility classes and helpers
  - `FileManager` - Temporary file management and cleanup
  - `Validators` - Input validation and PDF verification
  - `PageSelector` - Page selection parsing and processing

### Usage as a Library

```python
from report_compiler.core.compiler import ReportCompiler

# Basic usage
compiler = ReportCompiler("input.docx", "output.pdf")
compiler.compile()

# With debug mode
compiler = ReportCompiler("input.docx", "output.pdf", keep_temp=True)
compiler.compile()
```

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage

```bash
report-compiler compile input_report.docx output_report.pdf
```

### Debug Mode (with temp files)

```bash
report-compiler compile input_report.docx output_report.pdf --keep-temp
```

## Placeholder Format

The Report Compiler supports two types of PDF insertion placeholders:

### Table-based Overlays (OVERLAY tags)

For inserting PDFs as overlays onto existing pages, preserving the main document's content and layout. Place these in **single-cell (1x1) tables**:

```text
[[OVERLAY: appendices/sketch.pdf]]
[[OVERLAY: calculations/diagram.pdf, page=2]]
[[OVERLAY: C:\Shared\drawing.pdf, page=1-3]]
[[OVERLAY: diagrams/full_page.pdf, crop=false]]
[[OVERLAY: sketches/detail.pdf, page=2, crop=false]]
```

**OVERLAY Parameters:**

- `page=` - Page selection (same format as INSERT)
- `crop=` - Content cropping control:
  - `crop=true` (default): Automatically crops to content bounding box, removing excess whitespace
  - `crop=false`: Uses the full page dimensions without cropping

### Paragraph-based Merges (INSERT tags)

For inserting entire PDF pages after a marker position. The original paragraph content is preserved, and PDF pages are inserted immediately after it. Place these in **standalone paragraphs**:

```text
[[INSERT: appendices/structural_analysis.pdf]]
[[INSERT: calculations/load_analysis.pdf:1-5]]
[[INSERT: C:\Shared\external_report.pdf]]
```

### Page Selection

Both OVERLAY and INSERT tags support page selection:

**OVERLAY page selection (using `page=` parameter):**

```text
[[OVERLAY: appendices/report.pdf, page=5]]        # Page 5 only
[[OVERLAY: appendices/report.pdf, page=1-3]]      # Pages 1, 2, and 3
[[OVERLAY: appendices/report.pdf, page=1,3,5]]    # Pages 1, 3, and 5
[[OVERLAY: appendices/report.pdf, page=2-]]       # Pages 2 to end
```

**INSERT page selection (using `:` separator):**

```text
[[INSERT: appendices/report.pdf:1-3]]      # Pages 1, 2, and 3
[[INSERT: appendices/report.pdf:5]]        # Page 5 only
[[INSERT: appendices/report.pdf:1,3,5]]    # Pages 1, 3, and 5
[[INSERT: appendices/report.pdf:2-]]       # Pages 2 to end
[[INSERT: appendices/report.pdf:1-3,7,9-]] # Mixed: pages 1-3, 7, and 9 to end
```

**Page Selection Formats:**

- `5` - Single page (page 5)
- `1-3` - Range of pages (pages 1, 2, 3)
- `2-` - Open-ended range (pages 2 to end of document)
- `1,3,5` - Specific pages (pages 1, 3, and 5)
- `1-3,7,9-12` - Combined specifications

**Note:** Page numbers are 1-indexed (first page = 1). Invalid page numbers are automatically filtered out.

**Multi-page PDFs**: Automatically handled via cell replication (table-based overlays) or sequential page insertion (paragraph-based merges)

**Note**: Relative paths are resolved relative to the Word document's location.

## How It Works

### 1. Placeholder Detection

- **Table scanning** - Identifies `[[OVERLAY: ...]]` tags in single-cell tables
- **Paragraph scanning** - Identifies `[[INSERT: ...]]` tags in standalone paragraphs  
- **Path resolution** - Resolves relative paths relative to Word document location
- **Page parsing** - Parses page selection syntax (e.g., `:1-3`, `,page=5`)
- **PDF validation** - Validates that referenced PDF files exist and are readable
- **Page counting** - Counts effective pages after applying page selection filters
- **Layout detection** - Identifies single-cell tables vs standalone paragraphs

### 2. Document Modification

- **Table placeholders** - Replaces with visible red markers (`%%OVERLAY_START_N%%`)
- **Cell replication** - Creates additional table cells for multi-page selections
- **Paragraph placeholders** - Replaces with merge markers and page breaks (`%%MERGE_START_N%%`)  
- **Marker placement** - Places markers first, then page breaks for correct timing
- **Temporary document** - Saves modified document for PDF conversion

### 3. PDF Conversion

- Converts modified Word document to PDF using Word automation
- Preserves formatting and creates base PDF with markers

### 4. PDF Processing

#### Paragraph-based Merges (INSERT)

- **Marker location** - Finds merge markers in the base PDF
- **Marker removal** - Removes markers using redaction (white fill)
- **Page insertion** - Inserts PDF pages immediately after marker position
- **Content preservation** - Original document content remains intact

#### Table-based Overlays (OVERLAY)

- **Page selection** - Processes only the specified pages from source PDFs
- **Annotation preservation** - Automatically bakes PDF annotations into content using `Document.bake()`
- **Multi-page support** - Creates additional table cells for multi-page selections
- **Precise positioning** - Searches for overlay markers in the base PDF
- **Rectangle calculation** - Uses the marker position as the top-left corner of the overlay area
- **Marker removal** - Removes markers using redaction (white fill)
- **Sequential overlay** - Overlays each selected page onto calculated rectangles
- **Final assembly** - Saves completed PDF with all appendices integrated

## Table-Based Overlay System

The Report Compiler uses a precise approach for PDF overlay placement with full support for multi-page PDFs and annotation preservation:

### Single-Page PDF Overlay

1. **Table Detection** - Identifies single-cell tables containing `[[OVERLAY: path.pdf]]` placeholders
2. **Page Selection** - Parses page specifications like `,page=1-3` or `,page=5` if provided
3. **Dimension Extraction** - Extracts exact table dimensions from Word document metadata  
4. **Marker Placement** - Places a red marker at the top-left of the table cell
5. **Rectangle Calculation** - Uses marker position + table dimensions = overlay area
6. **Annotation Preservation** - Bakes PDF annotations into content before overlay
7. **Precise Overlay** - Places selected PDF pages exactly within the calculated rectangle

### Multi-Page PDF Overlay

For multi-page PDFs or page selections, the system automatically replicates table cells:

1. **Page Detection** - Identifies PDFs with multiple pages or page selections
2. **Cell Replication** - Adds consecutive table rows for each selected page
3. **Marker Generation** - Creates unique markers for each cell (`%%OVERLAY_START_00_PAGE_02%%`)
4. **Sequential Overlay** - Overlays selected pages into consecutive table cells
5. **Unified Layout** - All selected PDF pages appear together in the same table area

### Page Selection Examples

```text
[[OVERLAY: report.pdf, page=1-3]]     → 3 table cells with pages 1, 2, 3
[[OVERLAY: report.pdf, page=2,5,7]]   → 3 table cells with pages 2, 5, 7  
[[OVERLAY: report.pdf, page=3-]]      → Multiple cells with pages 3 to end
```

### Example Output

```text
Single Table → Page Selection:
┌─────────────────┐
│ PDF Page 2      │ ← Only page 2 (from [[OVERLAY: doc.pdf, page=2]])
└─────────────────┘

Single Table → Multi-Page Selection:  
┌─────────────────┐
│ PDF Page 1      │ ← From [[OVERLAY: doc.pdf, page=1,3,5]]
├─────────────────┤
│ PDF Page 3      │ ← Replicated cell  
├─────────────────┤
│ PDF Page 5      │ ← Replicated cell
└─────────────────┘
```

### Example Debug Output

```text
📋 Table found: 7.50 x 4.00 inches
📍 Marker at: (0.50, 1.59) inches  
📐 Overlay: (0.50, 1.59) to (8.00, 5.59) inches
🔥 Baking annotations: 12 found
✅ PDF positioned perfectly
```

### Key Benefits

- **Simple & Reliable** - Single marker approach with cell replication
- **Flexible Page Selection** - Extract exactly the pages you need from large PDFs
- **Multi-page Support** - Automatic handling of PDFs with any number of pages
- **Annotation Preservation** - PDF annotations automatically preserved during overlay
- **Accurate** - Uses Word's own measurements
- **Easy to Debug** - Clear inch measurements and detailed logging with page selection info
- **Consistent** - Predictable placement and unified layout

## Example Workflow

```text
Input: bridge_report.docx containing [[INSERT: appendices/analysis.pdf:2-4,7]]
↓
Step 1: Find placeholder and validate analysis.pdf (10 pages)
       Parse page spec "2-4,7" → pages 2, 3, 4, 7 (4 pages selected)
↓
Step 2: Replace placeholder with marker + replicate table cells for 4 pages
↓
Step 3: Convert modified DOCX to PDF (creates base PDF with 4 table cells)
↓
Step 4: Bake annotations, find markers, overlay pages 2,3,4,7 sequentially
↓
Output: bridge_report.pdf with selected pages integrated in consecutive cells
```

## Requirements

- **Windows** (for Word automation via win32com)
- **Microsoft Word** installed and accessible
- **Python 3.7+**
- **Dependencies**: `python-docx`, `pywin32`, `PyMuPDF`

## VS Code Debugging

The project includes comprehensive VS Code launch configurations:

- **Debug Report Compiler - Example File** - Basic debugging with example file
- **Debug Report Compiler - Example File (Keep Temp)** - Debug with temp files retained
- **Debug Report Compiler - Custom Input** - Interactive file input debugging
- **Debug Report Compiler - Step Into All Code** - Detailed debugging with all code
- **Debug Report Compiler - Error Testing** - Test error handling scenarios

## License

This project is licensed under the MIT License - see the LICENSE file for details.
