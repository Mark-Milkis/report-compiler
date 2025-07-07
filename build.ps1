# PowerShell build script for Report Compiler Windows executable
# This script creates a single-file executable using PyInstaller

Write-Host "Starting Report Compiler build process..." -ForegroundColor Green
Write-Host ""

# Check if PyInstaller is installed
try {
    python -c "import PyInstaller" 2>$null
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller not found"
    }
}
catch {
    Write-Host "PyInstaller not found. Installing build requirements..." -ForegroundColor Yellow
    pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Failed to install build requirements." -ForegroundColor Red
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Clean previous builds
Write-Host "Cleaning previous builds..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item "build" -Recurse -Force }
if (Test-Path "dist") { Remove-Item "dist" -Recurse -Force }
if (Test-Path "report-compiler.exe") { Remove-Item "report-compiler.exe" -Force }

# Convert PNG icon to ICO if Pillow is available and ICO doesn't exist
if (-not (Test-Path "word_integration\icons\compile-report.ico")) {
    Write-Host "Converting icon from PNG to ICO..." -ForegroundColor Yellow
    python -c @"
try:
    from PIL import Image
    img = Image.open('word_integration/icons/compile-report.png')
    img.save('word_integration/icons/compile-report.ico', format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128)])
    print('Icon converted successfully.')
except ImportError:
    print('Pillow not available, skipping icon conversion.')
except Exception as e:
    print(f'Icon conversion failed: {e}')
"@
}

# Build the executable
Write-Host "Building executable with PyInstaller..." -ForegroundColor Green
pyinstaller report_compiler.spec

if ($LASTEXITCODE -ne 0) {
    Write-Host "Build failed!" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Copy executable to root directory for convenience
if (Test-Path "dist\report-compiler.exe") {
    Copy-Item "dist\report-compiler.exe" "report-compiler.exe"
    Write-Host ""
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "Executable created: report-compiler.exe" -ForegroundColor Cyan
    
    # Show file size
    $fileInfo = Get-Item "report-compiler.exe"
    $sizeInMB = [math]::Round($fileInfo.Length / 1MB, 2)
    Write-Host "File size: $sizeInMB MB" -ForegroundColor Cyan
} else {
    Write-Host "Build completed but executable not found in expected location." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "You can now distribute the single file: report-compiler.exe" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
