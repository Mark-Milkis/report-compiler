# High-Efficiency DOCX and PDF Report Compiler

This tool allows engineers to write reports in Microsoft Word (.docx) and compile them into professional PDFs with seamlessly integrated external PDF appendices.

## Files Created

1. **`compile_report.bat`** - Main controller script
2. **`insert_pdf.lua`** - Pandoc Lua filter for PDF insertion
3. **`template.tex`** - Professional LaTeX template for clean business reports
4. **`modern_template.tex`** - Enhanced modern template with vibrant colors and creative design
5. **`creative_template.tex`** - Ultra-futuristic template with neon colors and creative styling
6. **`pdfpages_header.tex`** - LaTeX header that adds pdfpages package support

## How to Use

### 1. Write Your Report
- Create your report in Microsoft Word (.docx format)
- Save it in any location

### 2. Add PDF Placeholders
To insert external PDFs (like calculation outputs), add a placeholder on its own line:
```
[[INSERT: path/to/your/file.pdf]]
```

**Path Resolution:**
- **Relative paths**: Resolved relative to the Word document's location
  - `[[INSERT: appendices/calculations.pdf]]` → looks in `appendices/` folder next to your Word document
- **Absolute paths**: Used as-is
  - `[[INSERT: C:\Reports\SharedFiles\analysis.pdf]]` → exact path specified

Example in your Word document:
```
This is the main report content.

[[INSERT: appendices/structural analysis.pdf]]

The analysis above shows...
```

### 3. Compile the Report

You have two simple options for compiling your report:

#### Option A: Drag and Drop (Default Template)
- Simply drag your `.docx` file onto the `compile_report.bat` script
- Uses Pandoc's default template for clean, professional output

#### Option B: Command Line with Custom Template
```cmd
compile_report.bat --template "C:\path\to\your\custom.tex" your_report.docx
```

#### Option C: Command Line with Default Template
```cmd
compile_report.bat your_report.docx
```

The script will automatically:
- Convert your Word document to PDF
- Insert external PDFs at the specified locations
- Apply the specified styling option (custom template or Pandoc default)
- Create a final PDF with the same name as your Word document

### 4. Output
- A professional PDF will be created in the same folder as your Word document
- The command window will show progress and any errors

## Requirements

Before using this tool, ensure you have:
- **Pandoc** installed and available in your PATH
- **LaTeX distribution** (like MiXTeX or TeX Live) with `pdflatex`
- **pdfpages** LaTeX package (usually included in full LaTeX distributions)

## Template Options

You have flexible control over template usage with **three beautiful template choices**:

### Available Templates

1. **Default Template (Pandoc Built-in)**
   - Clean, professional, academic style
   - Automatic when no template is specified

2. **Professional Template (`template.tex`)**
   - Business-focused design with company branding support
   - Professional color scheme and typography
   ```cmd
   compile_report.bat --template "template.tex" report.docx
   ```

3. **Modern Template (`modern_template.tex`)**
   - Enhanced modern design with vibrant colors
   - Creative headers/footers and colorful styling
   - Perfect for contemporary engineering reports
   ```cmd
   compile_report.bat --template "modern_template.tex" report.docx
   ```

4. **Creative Template (`creative_template.tex`)**
   - Ultra-futuristic design with neon colors
   - Cutting-edge styling with creative visual elements
   - Perfect for innovative tech projects and presentations
   ```cmd
   compile_report.bat --template "creative_template.tex" report.docx
   ```

### Simple Two-Argument Design
- **Required**: Input `.docx` file
- **Optional**: `--template PATH` to specify a custom template file
- **Automatic PDF Support**: When using Pandoc's default template, the `pdfpages` package is automatically included to support PDF insertion

### Custom Template Usage
Use the `--template` flag to specify any template file:
```cmd
compile_report.bat --template "C:\MyCompany\report_template.tex" report.docx
```

### Default Template (Pandoc Built-in)
Both drag-and-drop and command-line without the template flag use Pandoc's default template:
```cmd
compile_report.bat your_report.docx
```

### Template Customization
All included templates can be:
- Used as-is for their specific styling
- Customized for your company branding
- Copied and modified for different report types
- Enhanced with metadata from your Word document:
  - Title, Author, Date, Company, Department, etc.

## Troubleshooting

If compilation fails, check:
1. Pandoc is installed: `pandoc --version`
2. LaTeX is installed: `pdflatex --version`
3. All referenced PDF files exist and are accessible
4. Your Word document is saved and not corrupted
5. File paths in placeholders use the correct format

## Example Usage with Sample Files

Your workspace contains sample files. To test:
1. Ensure the sample Word document contains a placeholder like:
   ```
   [[INSERT: C:\Users\p005452g\Source\report-compiler\examples\simple-report\appendices\structural analysis.pdf]]
   ```
2. Drag `bridge_report.docx` onto `compile_report.bat`
3. Check the generated `bridge_report.pdf`
