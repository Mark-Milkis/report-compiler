# Building Report Compiler Executable

This directory contains the configuration files needed to build a single-file Windows executable of the Report Compiler using PyInstaller.

## Files

- `report_compiler.spec` - PyInstaller specification file with all configuration
- `build.ps1` - PowerShell script to build the executable  
- `requirements-build.txt` - Additional Python packages needed for building

## Prerequisites

1. **Python 3.8+** with pip installed
2. **All runtime dependencies** installed:
   ```bash
   pip install -r requirements.txt
   ```

## Building the Executable


### Option 1: Using PowerShell Script
```powershell
.\build.ps1
```

### Option 2: Manual Build
```bash
# Install build requirements
pip install -r requirements-build.txt

# Clean previous builds (optional)
rmdir /s build dist 2>nul

# Build executable
pyinstaller report_compiler.spec
```

## Output

After successful build:
- Single executable file: `report-compiler.exe` (copied to project root)
- Full build output in `dist/` directory
- Build artifacts in `build/` directory

## Executable Features

The built executable includes:
- ✅ All Python dependencies bundled
- ✅ Report Compiler modules and utilities
- ✅ Windows COM support for Word automation
- ✅ PyMuPDF for PDF processing
- ✅ python-docx for DOCX handling
- ✅ Application icon (if available)
- ✅ Console interface for CLI usage

## Usage

The executable can be used exactly like the Python script:

```bash
# Compile a report
report-compiler.exe input.docx output.pdf

# Convert PDF to SVG
report-compiler.exe --action svg_import input.pdf output.svg --page 3

# Enable verbose logging
report-compiler.exe input.docx output.pdf --verbose

# Keep temporary files for debugging
report-compiler.exe input.docx output.pdf --keep-temp
```

## Distribution

The `report-compiler.exe` file is completely self-contained and can be distributed without requiring Python or any dependencies to be installed on the target machine.

**System Requirements for End Users:**
- Windows 7/8/10/11 (64-bit)
- Microsoft Word (for DOCX to PDF conversion) OR LibreOffice (alternative)
- No Python installation required

## Troubleshooting

### Build Issues

1. **Missing dependencies**: Ensure all packages in `requirements.txt` are installed
2. **PyInstaller errors**: Try updating PyInstaller: `pip install --upgrade pyinstaller`
3. **Icon conversion fails**: Install Pillow: `pip install pillow`

### Runtime Issues

1. **Word automation fails**: Ensure Microsoft Word is installed and can be automated
2. **PDF processing errors**: The executable includes PyMuPDF, but very large PDFs may cause memory issues
3. **Antivirus false positives**: Some antivirus software may flag PyInstaller executables - this is normal

### Size Optimization

The executable size can be reduced by:

- Removing unused modules from `hiddenimports` in the spec file
- Adding more modules to the `excludes` list
- Using UPX compression (enabled by default)

## Advanced Configuration

Edit `report_compiler.spec` to customize:

- **Icon**: Change `icon` parameter to use different icon file
- **Console**: Set `console=False` for windowed mode (not recommended for CLI tool)
- **Debug**: Set `debug=True` for debugging builds
- **Compression**: Modify UPX settings
- **Dependencies**: Add/remove modules from `hiddenimports` and `excludes`
