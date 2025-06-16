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
- ✅ **Robust page breaks** - Proper page breaks using `WD_BREAK.PAGE`
- ✅ **Visible markers** - Red markers that are automatically removed during overlay
- ✅ **Error handling** - Comprehensive error reporting and validation
- ✅ **Debug support** - `--keep-temp` flag to retain temporary files for debugging
- ✅ **VS Code integration** - Complete debugger launch configurations
- ✅ **Table-based overlay** - Precise PDF placement within table cell boundaries
- ✅ **Advanced table detection** - Uses PyMuPDF table detection with text analysis fallback
- ✅ **Intelligent positioning** - Automatically detects table structure for precise overlay placement

## Quick Start

### Installation

```bash
pip install -r requirements.txt
```

### Basic Usage

```bash
python report_compiler.py input_report.docx output_report.pdf
```

### Debug Mode

```bash
python report_compiler.py input_report.docx output_report.pdf --keep-temp
```

## Placeholder Format

In your Word document, use the following format to insert PDF appendices:

```
[[INSERT: appendices/structural_analysis.pdf]]
[[INSERT: calculations/load_analysis.pdf]]
[[INSERT: C:\Shared\external_report.pdf]]
```

**Note**: Relative paths are resolved relative to the Word document's location.

## VS Code Debugging

The project includes comprehensive VS Code launch configurations:

- **Debug Report Compiler - Example File** - Basic debugging with example file
- **Debug Report Compiler - Example File (Keep Temp)** - Debug with temp files retained
- **Debug Report Compiler - Custom Input** - Interactive file input debugging
- **Debug Report Compiler - Step Into All Code** - Detailed debugging with all code
- **Debug Report Compiler - Error Testing** - Test error handling scenarios

## How It Works

### 1. Document Analysis

- Scans Word document for `[[INSERT: path]]` placeholders
- Validates that referenced PDF files exist
- Counts pages in each appendix PDF

### 2. Document Modification

- Replaces placeholders with visible red markers (`%%APPENDIX_START_N%%`)
- Inserts proper page breaks for each page in the appendix
- Saves modified document as temporary DOCX

### 3. PDF Conversion

- Converts modified Word document to PDF using Word automation
- Preserves formatting and creates separate pages for overlays

### 4. PDF Overlay

- Searches for visible markers in the base PDF
- Removes markers using redaction (white fill)
- Overlays each appendix page onto corresponding blank pages
- Saves final compiled PDF

## Requirements

- **Windows** (for Word automation via win32com)
- **Microsoft Word** installed and accessible
- **Python 3.7+**
- **Dependencies**: `python-docx`, `pywin32`, `PyMuPDF`

## Example Workflow

```
Input: bridge_report.docx containing [[INSERT: appendices/structural_analysis.pdf]]
↓
Step 1: Find placeholder and validate structural_analysis.pdf (26 pages)
↓
Step 2: Replace placeholder with marker + 26 page breaks
↓
Step 3: Convert modified DOCX to PDF (creates 28-page base PDF)
↓
Step 4: Find marker on page 2, overlay 26 pages of structural_analysis.pdf
↓
Output: bridge_report.pdf with integrated appendices
```

## TODO - Future Improvements

### Core Functionality

- [ ] **Multiple placeholder support per document** - Handle multiple appendices in order
- [ ] **Recursive references** - Use recursion to handle references of nested Word documents
- [ ] **Cross-platform support** - LibreOffice integration for Linux/Mac
- [ ] **Batch processing** - Process multiple documents in one command
- [ ] **Template support** - Predefined document templates with placeholders

### Advanced Features

- [ ] **Smart page sizing** - Auto-scale appendix pages to match document format
- [ ] **Bookmark generation** - Auto-create PDF bookmarks for each appendix
- [ ] **Table of contents** - Auto-update TOC with appendix page numbers
- [ ] **Watermark support** - Add watermarks or headers to appendix pages
- [ ] **Page numbering** - Continuous page numbering across main document and appendices

### User Experience

- [ ] **GUI interface** - Desktop application with drag-and-drop
- [ ] **Configuration files** - JSON/YAML config for default settings
- [ ] **Progress indicators** - Real-time progress for large documents
- [ ] **Preview mode** - Preview final layout before compilation
- [ ] **Undo/rollback** - Ability to revert to previous versions

## Advanced Table Detection

The Report Compiler includes sophisticated coordinate mapping for precise PDF overlay placement using **Word-to-PDF Coordinate Mapping**:

### Word-to-PDF Coordinate Mapping

The system captures table positioning during Word document processing and converts it to PDF coordinates:

1. **Word Coordinate Extraction** - Captures table dimensions, position, and margins during document analysis
2. **Metadata Storage** - Stores coordinate information for each table with placeholders
3. **PDF Coordinate Conversion** - Converts Word measurements to PDF points during overlay
4. **Precise Positioning** - Uses calculated coordinates for exact overlay placement

### Detection Process

```text
� Stored coordinate metadata for table 0
🎯 Using Word-to-PDF coordinate mapping for table 0
🔄 Converting Word coordinates to PDF coordinates...
� Word table: 7.50" x 4.00" = 540.0 x 288.0 points
📍 Calculated position: margin(90.0) + indent(0.0) = x:90.0
� Y position: marker_top(114.2) - padding(10) = y:104.2
✅ Converted Word coordinates to PDF coordinates
```

### Coordinate Conversion Features

- **Dimension Mapping** - Converts Word table dimensions (inches) to PDF points (1" = 72 points)
- **Position Calculation** - Accounts for Word document margins and table indentation
- **Absolute Positioning** - Supports tables with explicit positioning in Word
- **Margin Awareness** - Uses standard Word margins (1.25" left/right) for accurate placement
- **Fallback Support** - Falls back to table detection methods if coordinate mapping fails

### Benefits

- **Pixel-Perfect Placement** - Overlays positioned exactly where the table appears in Word
- **Consistent Scaling** - PDF content scaled to match Word table dimensions precisely
- **Document Fidelity** - Maintains the exact layout intended in the original Word document
- **Robust Conversion** - Handles various Word table positioning scenarios

## License

This project is licensed under the MIT License - see the LICENSE file for details.
