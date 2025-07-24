# Word Integration Setup Guide

This guide explains how to install and use the Word integration features for Report Compiler.

## Overview

The Word integration provides:
- **Ribbon buttons** to insert placeholders automatically
- **File browser dialogs** for selecting PDF files with automatic relative path creation
- **One-click compilation** directly from Word
- **PDF to SVG conversion** for direct image insertion

## Installation Steps

### 1. Install Report Compiler

First, ensure Report Compiler is installed and accessible from the command line:

```bash
# Test that it's working
report-compiler --version
```

### 2. Install Word Template

1. **Download the template file**: `ReportCompilerTemplate.dotm` from the `word_integration/` folder

2. **Copy to Word startup folder**:
   ```
   Windows: %APPDATA%\Microsoft\Word\STARTUP\
   ```
   
   You can access this folder by:
   - Press `Win+R`, type `%APPDATA%\Microsoft\Word\STARTUP\`, press Enter
   - Or open File Explorer and paste the path in the address bar

3. **Restart Microsoft Word**

### 3. Verify Installation

After restarting Word, you should see a new "Report Compiler" group in the ribbon with these buttons:

- **Insert Appendix** - Adds INSERT placeholders for full-page PDF content
- **Insert Overlay** - Adds OVERLAY placeholders in tables for positioned content  
- **PDF to SVG** - Converts PDF pages to SVG images for direct insertion
- **Compile Report** - Runs the report compiler directly from Word

## Using the Word Integration

### Insert Appendix Button

1. Place cursor where you want to insert PDF pages
2. Click "Insert Appendix" 
3. Browse and select a PDF file
4. Optional: Enter page selection (e.g., "1-3,7")
5. The placeholder will be inserted: `[[INSERT: relative/path/file.pdf:1-3,7]]`

### Insert Overlay Button

1. Place cursor where you want positioned PDF content
2. Click "Insert Overlay"
3. Browse and select a PDF file
4. Optional: Enter page selection
5. Optional: Choose cropping behavior
6. A table with the placeholder will be created: `[[OVERLAY: relative/path/file.pdf, page=1-3]]`

### PDF to SVG Button

1. Click "PDF to SVG"
2. Select a PDF file
3. Enter page numbers to convert (e.g., "1", "1-3", "1,3,5", or "all")
4. SVG images will be created and can be inserted as regular images

### Compile Report Button

1. Save your Word document first
2. Click "Compile Report"
3. Choose output PDF location
4. The report will be compiled automatically

## Troubleshooting

### "Macros are disabled"
- Go to File → Options → Trust Center → Trust Center Settings → Macro Settings
- Select "Enable all macros" or "Disable all macros with notification"
- Restart Word

### "Report Compiler not found"
- Ensure `report-compiler` command works in Command Prompt
- Check that Python and Report Compiler are properly installed
- Verify PATH environment variable includes Python Scripts folder

### "Document must be saved first"
- Save your Word document before using any integration features
- This is required for relative path resolution

### Buttons not appearing
- Verify the template file is in the correct STARTUP folder
- Restart Word completely
- Check if template is blocked by security settings

## Features Explanation

### Automatic Relative Paths
The Word integration automatically creates relative paths based on your document's location:
- If document is in `C:\Reports\project.docx`
- And you select `C:\Reports\pdfs\data.pdf`  
- The placeholder will use `pdfs\data.pdf`

### Content Controls
Placeholders are wrapped in Word Content Controls to:
- Prevent accidental editing
- Provide visual distinction
- Enable easy selection and modification

### Error Handling
The integration includes error handling for:
- Missing files
- Invalid path formats
- Compilation failures
- Word automation issues

## Customization

### Modifying the Template

To customize the Word integration:

1. Open `ReportCompilerTemplate.dotm` in Word
2. Go to Developer tab → Visual Basic (or press Alt+F11)
3. Modify the `ReportingTools` module
4. Save and restart Word

### Adding Custom Buttons

You can add new ribbon buttons by:
1. Modifying the `report_compiler_UI.xml` file
2. Adding corresponding VBA procedures
3. Updating the template file

## Best Practices

1. **Always save your document** before inserting placeholders
2. **Use descriptive folder structures** for better organization
3. **Test compilation frequently** during document development
4. **Keep PDF files close** to your Word document for shorter relative paths
5. **Use consistent naming conventions** for easier management

## Advanced Usage

### Batch Processing
The Word integration can be extended to process multiple documents:
- Create a macro that iterates through document folders
- Use the compilation functions programmatically
- Automate report generation workflows

### Custom Page Selection
Take advantage of flexible page selection:
- `1-3,7,10-` for complex page ranges
- Test page selections before final compilation
- Use PDF viewers to identify correct page numbers

### Integration with Document Management
The Word integration works well with:
- SharePoint document libraries
- Version control systems
- Automated document workflows
- Template-based report systems
