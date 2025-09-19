# Report Compiler

A powerful automated tool for engineering teams to create professional PDF reports by embedding PDF content directly into Word documents. Write your reports in Word, add simple placeholders for external PDFs, and compile everything into a polished final report with a single command.

## What It Does

Transform your Word documents into comprehensive PDF reports by:

- **Adding PDF placeholders** in your Word document using simple tags
- **Automatically inserting** external PDF content (drawings, calculations, appendices)
- **Precisely positioning** PDF overlays within tables or merging full pages
- **Compiling everything** into a single professional PDF report

Perfect for engineering reports, technical documentation, and any workflow where you need to combine Word content with external PDF files.

## Quick Example

1. **Write your report in Word** with placeholders:

   ```text
   This is my engineering report. 
   
   [[INSERT: calculations/load_analysis.pdf:1-3]]
   
   The table below shows the design sketch:
   ┌─────────────────────────────────────┐
   │ [[OVERLAY: drawings/sketch.pdf]]    │
   └─────────────────────────────────────┘
   
   Here's the company logo:
   ┌─────────────────────────────────────┐
   │ [[IMAGE: assets/logo.png]]          │
   └─────────────────────────────────────┘
   ```

2. **Run the compiler**:

   ```bash
   report-compiler report.docx final_report.pdf
   ```

3. **Get a professional PDF** with all content seamlessly integrated!

## Installation

### Option 1: Using uvx (Recommended)

The easiest way to install and run Report Compiler is with [uvx](https://docs.astral.sh/uv/guides/tools/):

```bash
# Install uv if not already installed. Required version of python will be installed automatically by uv.
winget install --id=astral-sh.uv  -e

# Install and run directly (no permanent installation)
uvx report-compiler@latest compile report.docx output.pdf

# Or install globally for repeated use
uvx install report-compiler
report-compiler compile report.docx output.pdf
```

### Option 2: Traditional Installation

```bash
# Install from PyPI
pip install report-compiler

# Or install from source
git clone https://github.com/your-repo/report-compiler.git
cd report-compiler
pip install -e .
```

### Word Integration (Optional)

For enhanced productivity, install the Word add-in that provides buttons to insert placeholders and compile reports:

1. **Copy the template file** to your Word templates folder:

   ```bash
   # Download and copy ReportCompilerTemplate.dotm to:
   # Windows: %APPDATA%\Microsoft\Word\STARTUP\
   ```

2. **Restart Word** - you'll see new "Report Compiler" ribbon buttons

3. **Use the buttons** to insert placeholders instead of typing them manually

## Requirements

- **Windows** (for Word automation)
- **Microsoft Word** installed
- **Python 3.7+**

## Usage

### Command Line Interface

#### Basic Compilation

```bash
# Compile a Word document to PDF
report-compiler compile input.docx output.pdf

# Enable debug mode (keeps temporary files)
report-compiler compile input.docx output.pdf --keep-temp

# Verbose logging
report-compiler compile input.docx output.pdf --verbose
```

#### PDF to SVG Conversion

```bash
# Convert single page
report-compiler svg-import input.pdf output.svg --page 3

# Convert multiple pages
report-compiler svg-import input.pdf output.svg --page 1-5

# Convert all pages
report-compiler svg-import input.pdf output.svg --page all
```

### Using Placeholders in Word

There are two types of placeholders you can use:

#### 1. INSERT Placeholders (Full Page Merging)

Use `INSERT` placeholders to add complete PDF pages into your document. Place these in standalone paragraphs:

```text
[[INSERT: appendices/structural_analysis.pdf]]
[[INSERT: calculations/load_analysis.pdf:1-5]]
[[INSERT: external/report.pdf:2,4,6]]
```

**Page Selection Options:**

- `[[INSERT: file.pdf]]` - All pages
- `[[INSERT: file.pdf:5]]` - Page 5 only
- `[[INSERT: file.pdf:1-3]]` - Pages 1, 2, and 3
- `[[INSERT: file.pdf:1,3,5]]` - Pages 1, 3, and 5
- `[[INSERT: file.pdf:2-]]` - Pages 2 to end
- `[[INSERT: file.pdf:1-3,7,9-]]` - Combined: pages 1-3, 7, and 9 to end

#### 2. OVERLAY Placeholders (Table-Based Positioning)

Use `OVERLAY` placeholders to position PDF content precisely within tables. Place these inside single-cell tables:

```text
[[OVERLAY: drawings/sketch.pdf]]
[[OVERLAY: diagrams/detail.pdf, page=2]]
[[OVERLAY: sketches/plan.pdf, page=1-3]]
[[OVERLAY: drawings/full_page.pdf, crop=false]]
```

**Parameters:**

- `page=` - Specify which pages to overlay (same format as INSERT)
- `crop=` - Control content cropping:
  - `crop=true` (default): Auto-crop to remove whitespace
  - `crop=false`: Use full page dimensions

#### 3. IMAGE Placeholders (Direct Image Insertion)

Use `IMAGE` placeholders to insert image files directly into Word documents. Place these inside single-cell tables:

```text
[[IMAGE: photos/diagram.png]]
[[IMAGE: charts/graph.jpg, width=3in]]
[[IMAGE: icons/logo.png, width=2in, height=1in]]
```

**Supported formats:** PNG, JPG, JPEG, GIF, BMP, TIFF, WEBP

**Parameters:**

- `width=` - Set image width (e.g., `width=2in`, `width=100px`)
- `height=` - Set image height (e.g., `height=1.5in`, `height=200px`)
- If no dimensions specified: Auto-fits to table size while maintaining aspect ratio
- If only width or height specified: Maintains aspect ratio

#### Path Resolution

All paths are resolved relative to your Word document's location:

```text
# If your Word document is in C:\Reports\
[[INSERT: appendices/data.pdf]]        → C:\Reports\appendices\data.pdf
[[OVERLAY: ..\shared\drawing.pdf]]     → C:\shared\drawing.pdf
[[IMAGE: images/chart.png]]            → C:\Reports\images\chart.png
[[INSERT: C:\absolute\path\file.pdf]]  → C:\absolute\path\file.pdf
```

### Word Integration Buttons

If you installed the Word template, you'll have these ribbon buttons:

- **Insert Appendix** - Adds an INSERT placeholder with file browser
- **Insert Overlay** - Adds an OVERLAY placeholder in a table with options
- **PDF to SVG** - Converts PDF pages to SVG for direct insertion
- **Compile Report** - Runs the compiler directly from Word

### Example Workflow

1. **Create your report structure** in Word with regular text and formatting
2. **Add placeholders** where you want PDF content:
   - Use tables with OVERLAY placeholders for positioned content
   - Use paragraphs with INSERT placeholders for full-page appendices
3. **Save your document** (required for relative path resolution)
4. **Run the compiler** from command line or Word button
5. **Get your final PDF** with all content integrated seamlessly

## Troubleshooting

### Common Issues

#### "Document must be saved first"

- Save your Word document before compilation to enable relative path resolution

#### "PDF file not found"

- Check that PDF paths are correct relative to your Word document's location
- Use the Word integration buttons for automatic relative path creation

#### "Word automation failed"

- Ensure Microsoft Word is installed and can be opened
- Close any open Word documents that might interfere

#### "Page selection invalid"

- Check page numbers exist in the source PDF
- Use 1-based page numbering (first page = 1)

### Debug Mode

Enable debug mode to troubleshoot issues:

```bash
report-compiler compile report.docx output.pdf --keep-temp --verbose
```

This will:

- Keep all temporary files for inspection
- Show detailed processing logs
- Help identify where issues occur

## System Requirements

- **Windows** (for Word automation)
- **Microsoft Word** installed
- **Python 3.7+**

## License

This project is licensed under the MIT License.
