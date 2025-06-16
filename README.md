# Report Compiler

A Python-based automated DOCX and PDF report compiler for engineering teams. This tool allows engineers to write reports in Word, use placeholders to insert external PDFs, and compile everything into a professional PDF with a single command.

## Overview

The Report Compiler automates the creation of comprehensive PDF reports by:

1. **Finding PDF placeholders** in Word documents (`[[INSERT: path/to/file.pdf]]`)
2. **Modifying the Word document** to create blank pages with visible markers
3. **Converting to PDF** using Word automation (win32com)
4. **Overlaying appendix PDFs** onto the blank pages using PyMuPDF

## Features

- ✅ **Relative path support** - PDF paths resolved relative to the input Word document
- ✅ **Page selection support** - Specify which pages to include from source PDFs using flexible syntax
- ✅ **Multi-page PDF support** - Automatic cell replication for multi-page table overlays
- ✅ **Annotation preservation** - PDF annotations automatically baked into content during overlay
- ✅ **Robust page breaks** - Proper page breaks using `WD_BREAK.PAGE`
- ✅ **Visible markers** - Red markers that are automatically removed during overlay
- ✅ **Error handling** - Comprehensive error reporting and validation
- ✅ **Debug support** - `--keep-temp` flag to retain temporary files for debugging
- ✅ **VS Code integration** - Complete debugger launch configurations
- ✅ **Table-based overlay** - Precise PDF placement using table dimensions and marker positioning
- ✅ **Cell replication** - Multi-page PDFs create consecutive table cells automatically
- ✅ **Intelligent positioning** - Uses table properties for automatic overlay rectangle calculation

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage

```bash
python report_compiler.py input_report.docx output_report.pdf
```

### Debug Mode (with temp files)

```bash
python report_compiler.py input_report.docx output_report.pdf --keep-temp
```

### Disable Annotation Preservation (for faster processing)

```bash
python report_compiler.py input_report.docx output_report.pdf --no-annotations
```

## Placeholder Format

In your Word document, use the following format to insert PDF appendices:

```text
[[INSERT: appendices/structural_analysis.pdf]]
[[INSERT: calculations/load_analysis.pdf]]
[[INSERT: C:\Shared\external_report.pdf]]
```

### Page Selection (New!)

You can now specify which pages to include from the source PDF:

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

**Supported Formats:**

- **Table-based overlays**: Single-cell tables containing `[[INSERT: path.pdf]]` for precise placement
- **Paragraph-based merges**: Regular paragraphs containing `[[INSERT: path.pdf]]` for full-page insertion

**Multi-page PDFs**: Automatically handled via cell replication (table-based) or sequential page insertion (paragraph-based)

**Note**: Relative paths are resolved relative to the Word document's location.

## How It Works

### 1. Document Analysis

- **Scans Word document** for `[[INSERT: path]]` and `[[INSERT: path:pages]]` placeholders
- **Page specification parsing** - Handles single pages, ranges, and complex combinations
- **PDF validation** - Verifies that referenced PDF files exist and page selections are valid
- **Page counting** - Counts effective pages after applying page selection filters
- **Layout detection** - Identifies table-based overlays vs paragraph-based merges

### 2. Document Modification

- **Table placeholders** - Replaces with visible red markers (`%%OVERLAY_START_N%%`)
- **Cell replication** - Creates additional table cells for multi-page selections
- **Paragraph placeholders** - Replaces with merge markers (`%%MERGE_START_N%%`)  
- **Page break insertion** - Inserts proper page breaks for each selected page
- **Temporary document** - Saves modified document for PDF conversion

### 3. PDF Conversion

- Converts modified Word document to PDF using Word automation
- Preserves formatting and creates separate pages for overlays

### 4. PDF Overlay

- **Page selection** - Processes only the specified pages from source PDFs
- **Annotation preservation** - Automatically bakes PDF annotations into content using `Document.bake()`
- **Multi-page support** - Creates additional table cells for multi-page selections
- **Precise positioning** - Searches for overlay markers in the base PDF
- **Rectangle calculation** - Uses the marker position as the top-left corner of the overlay area
- **Marker removal** - Removes markers using redaction (white fill)
- **Sequential overlay** - Overlays each selected page onto calculated rectangles
- **Final assembly** - Saves completed PDF with all appendices integrated

## Table-Based Overlay System

The Report Compiler uses a simple but precise approach for PDF overlay placement with full support for multi-page PDFs and annotation preservation:

### Single-Page PDF Overlay

1. **Table Detection** - Identifies single-cell tables containing `[[INSERT: path.pdf]]` placeholders
2. **Page Selection** - Parses page specifications like `:1-3` or `:5` if provided
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
[[INSERT: report.pdf:1-3]]     → 3 table cells with pages 1, 2, 3
[[INSERT: report.pdf:2,5,7]]   → 3 table cells with pages 2, 5, 7  
[[INSERT: report.pdf:3-]]      → Multiple cells with pages 3 to end
```

### Example Output

```text
Single Table → Page Selection:
┌─────────────────┐
│ PDF Page 2      │ ← Only page 2 (from [[INSERT: doc.pdf:2]])
└─────────────────┘

Single Table → Multi-Page Selection:  
┌─────────────────┐
│ PDF Page 1      │ ← From [[INSERT: doc.pdf:1,3,5]]
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
