# Word Integration Setup Guide

This guide explains how to install and use the Word integration features for Report Compiler.

## Overview

The Word integration provides:
- **Ribbon buttons** to insert placeholders automatically
- **File browser dialogs** for selecting PDF files with automatic relative path creation
- **One-click compilation** directly from Word
- **PDF page insertion** as high-quality images

### Architecture

Instead of spawning console processes, the ribbon talks to Report Compiler through a
local **COM server** (`ProgID ReportCompiler.Application`). The **Compile Report** and
**Insert PDF Page (Image)** buttons call the server, which runs the work on a background
thread and reports status back — so Word stays responsive and shows a real
success/failure message. This server must be registered once per user (no admin
required); see [Register the COM server](#3-register-the-com-server).

### Platform Support

- **Windows**: Full support with Microsoft Word
- **macOS**: Full support with Microsoft Word
- **Linux**: Word integration not available (Word not supported on Linux)

The `uvx report-compiler word-integration` commands will automatically detect your platform and provide appropriate error messages on unsupported platforms.

## Installation Steps

### 1. Install Report Compiler

First, ensure Report Compiler is installed and accessible from the command line:

```bash
# Test that it's working
report-compiler --version
```

### 2. Install Word Template

> **Note:** `ReportCompilerTemplate.dotm` is a build artifact, not tracked in git. If you
> are working from a source checkout and the file is missing, build it first (requires
> Word — see [Building the template](#building-the-template-from-sources)):
> ```bash
> uvx report-compiler word-integration build-template
> ```

**Option 1: Using uvx (Recommended)**

```bash
# Install Word integration template automatically
uvx report-compiler word-integration install
```

This will:
- Automatically detect your Word startup folder
- Copy the template file to the correct location
- Provide instructions for next steps

**Option 2: Manual Installation**

1. **Download the template file**: `ReportCompilerTemplate.dotm` (build it from a source checkout with `word-integration build-template`, or grab it from a release)

2. **Copy to Word startup folder**:
   ```
   Windows: %APPDATA%\Microsoft\Word\STARTUP\
   ```
   
   You can access this folder by:
   - Press `Win+R`, type `%APPDATA%\Microsoft\Word\STARTUP\`, press Enter
   - Or open File Explorer and paste the path in the address bar

3. **Restart Microsoft Word**

### 3. Register the COM server

The **Compile Report** and **Insert PDF Page (Image)** buttons talk to the
`ReportCompiler.Application` COM server, which must be registered once per user:

```bash
uvx report-compiler com-server register
```

This will:
- Ensure a stable install of Report Compiler exists (via `uv tool install`)
- Register the server per-user under `HKCU\Software\Classes` (no administrator rights)

Check it any time with:

```bash
uvx report-compiler com-server status      # "Registered: Yes"
```

To remove it later: `uvx report-compiler com-server unregister`.

> The placeholder buttons (Insert Appendix / Overlay / Image) do **not** need the COM
> server — they only insert text into the document.

### 4. Verify Installation

After installation, restart Microsoft Word and look for the "Report Compiler" tab in the ribbon. If you don't see it:

1. **Check the installation status**:
   ```bash
   uvx report-compiler word-integration status
   ```

2. **Verify the template is in the correct location**
3. **Make sure macros are enabled** (see Troubleshooting section)

You should see these buttons on the **Reporting Tools** tab:

- **Insert Appendix** - Adds INSERT placeholders for full-page PDF content
- **Insert Overlay** - Adds OVERLAY placeholders in tables for positioned content
- **Insert Image** - Adds IMAGE placeholders in tables for image files
- **Insert PDF Page (Image)** - Converts PDF page(s) to SVG and inserts them as images (via the COM server)
- **Compile Report** - Compiles the report directly from Word (via the COM server)

## Managing Word Integration

### Check Installation Status

```bash
# Get detailed status of Word integration
uvx report-compiler word-integration status
```

This will show:
- Platform support status
- Word startup folder location
- Template installation status
- Source template availability

### Update Word Integration

```bash
# Update to latest Word integration template
uvx report-compiler word-integration update
```

Use this when:
- A new version of Report Compiler is released
- You want to ensure you have the latest template features
- The integration stops working after an update

### Remove Word Integration

```bash
# Remove Word integration template
uvx report-compiler word-integration remove
```

This will:
- Remove the template from your Word startup folder
- Clean up the integration completely
- Require Word restart to complete removal

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

### Insert Image Button

1. Place cursor where you want an image
2. Click "Insert Image"
3. Browse and select an image file
4. A table with the placeholder will be created: `[[IMAGE: relative/path/image.png]]`

### Insert PDF Page (Image) Button

1. Save your document first (a temporary folder is created next to it)
2. Click "Insert PDF Page (Image)"
3. Select a PDF file
4. Enter page numbers to convert (e.g., "1", "1-3", "1,3,5", or "all")
5. The COM server converts the page(s) to SVG; Word inserts them as images at the cursor

### Compile Report Button

1. Save your Word document first
2. Click "Compile Report"
3. The COM server compiles the document; the PDF is written next to it (same name, `.pdf`)
4. Word stays responsive and shows a success or failure message when finished

> Both buttons require the COM server to be registered (see
> [Register the COM server](#3-register-the-com-server)). If it isn't, the button shows
> a "COM Server Not Registered" message with the command to run.

## Troubleshooting

### "Macros are disabled"
- Go to File → Options → Trust Center → Trust Center Settings → Macro Settings
- Select "Enable all macros" or "Disable all macros with notification"
- Restart Word

### "COM Server Not Registered"
- The Compile Report / Insert PDF Page buttons couldn't reach the COM server
- Register it: `uvx report-compiler com-server register`
- Verify with `uvx report-compiler com-server status` (should say "Registered: Yes")

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

The template is **built from plain-text sources** — no Office RibbonX Editor or manual
copy-paste required. The tracked sources live in `src/report_compiler/word_integration/` (shipped with the package):

| Source | Purpose |
|--------|---------|
| `*.bas` (`ReportingTools.bas`, `LibFileTools.bas`) | VBA macro code |
| `report_compiler_UI.xml` | Ribbon definition (becomes `customUI/customUI14.xml`) |
| `icons/*.png` | Ribbon button images |
| `skeleton/` | Static OPC parts (document/styles/theme/rels/content-types) |

`ReportCompilerTemplate.dotm` and `vbaProject.bin` are **build artifacts** (git-ignored),
produced from those sources.

### Building the template from sources

A `.dotm` is an OPC ZIP; everything except `word/vbaProject.bin` is plain text/PNG, so
the package is assembled with pure Python. `vbaProject.bin` is the compiled VBA blob —
Word is the only reliable way to author it.

```bash
# Full rebuild: compile the .bas modules (needs Word) then package the .dotm
uvx report-compiler word-integration build-template

# Or run the stages individually:
uvx report-compiler word-integration build-vba    # .bas -> vbaProject.bin (needs Word)
uvx report-compiler word-integration package      # assemble the .dotm (no Word)

# Then push the rebuilt template into Word's STARTUP folder
uvx report-compiler word-integration update
```

`build-vba` requires Word's "Trust access to the VBA project object model" setting; the
command enables it for you (`HKCU`). After any change to a `.bas` file or the ribbon XML,
re-run `build-template` (or `package` if only the ribbon/icons changed) and `update`.

### Modifying the macros

1. Edit the `.bas` files in `src/report_compiler/word_integration/` (e.g. `ReportingTools.bas`)
2. Run `build-template`, then `update`
3. Restart Word

### Adding Custom Buttons

1. Add the button to `report_compiler_UI.xml` (referencing an `image=` id and an
   `onAction=` callback)
2. Add the matching image to `icons/` and a relationship in
   `skeleton/customUI/_rels/customUI14.xml.rels`
3. Add the corresponding VBA procedure to `ReportingTools.bas`
4. Run `build-template`, then `update`

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
