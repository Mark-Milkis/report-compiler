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
- ✅ **Table-based overlay** - Precise PDF placement using table dimensions and marker positioning
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

## How It Works

### 1. Document Analysis

- Scans Word document for `[[INSERT: path]]` placeholders
- Validates that referenced PDF files exist
- Counts pages in each appendix PDF
- Identifies table-based overlays vs paragraph-based merges

### 2. Document Modification

- Replaces table placeholders with visible red markers (`%%OVERLAY_START_N%%`)
- Replaces paragraph placeholders with merge markers (`%%MERGE_START_N%%`)
- Inserts proper page breaks for each page in the appendix
- Saves modified document as temporary DOCX

### 3. PDF Conversion

- Converts modified Word document to PDF using Word automation
- Preserves formatting and creates separate pages for overlays

### 4. PDF Overlay

- Searches for overlay markers in the base PDF
- Uses the marker position as the top-left corner of the overlay area
- Calculates the overlay rectangle using the table's actual dimensions
- Removes markers using redaction (white fill)
- Overlays each appendix page onto the calculated rectangle
- Saves final compiled PDF

## Table-Based Overlay System

The Report Compiler uses a simple but precise approach for PDF overlay placement:

### Overlay Process

1. **Table Detection** - Identifies single-cell tables containing `[[INSERT: path.pdf]]` placeholders
2. **Dimension Extraction** - Extracts exact table dimensions from Word document metadata  
3. **Marker Placement** - Places a red marker at the top-left of the table cell
4. **Rectangle Calculation** - Uses marker position + table dimensions = overlay area
5. **Precise Overlay** - Places PDF content exactly within the calculated rectangle

### Example

```text
📋 Table found: 7.50 x 4.00 inches
📍 Marker at: (0.50, 1.59) inches  
📐 Overlay: (0.50, 1.59) to (8.00, 5.59) inches
✅ PDF positioned perfectly
```

### Key Benefits

- **Simple & Reliable** - Single marker approach
- **Accurate** - Uses Word's own measurements
- **Easy to Debug** - Clear inch measurements
- **Consistent** - Predictable placement

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
