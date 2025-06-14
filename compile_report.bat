@echo off
setlocal enabledelayedexpansion

REM Simple argument parsing
set "INPUT_FILE="
set "TEMPLATE_FILE="

REM Check for --template flag
if /i "%1"=="--template" (
    set "TEMPLATE_FILE=%2"
    set "INPUT_FILE=%3"
) else (
    set "INPUT_FILE=%1"
)

REM Check if a file was provided
if not defined INPUT_FILE goto show_usage
if "!INPUT_FILE!"=="" goto show_usage
goto continue_processing

:show_usage
echo Error: No file provided.
echo.
echo Usage: %~nx0 [--template PATH] input.docx
echo.
echo Arguments:
echo   input.docx          Input Word document (required)
echo.
echo Options:
echo   --template PATH     Use specified template file (optional)
echo.
echo Examples:
echo   %~nx0 report.docx
echo   %~nx0 --template custom.tex report.docx
echo.
echo You can also drag a .docx file onto this script.
pause
exit /b 1

:continue_processing

REM Get the input file path and validate it's a .docx file
REM Extract extension from INPUT_FILE
for %%i in ("!INPUT_FILE!") do set "FILE_EXT=%%~xi"

if /i not "!FILE_EXT!"==".docx" (
    echo Error: Input file must be a .docx file.
    echo Provided file: !INPUT_FILE!
    echo File extension found: '!FILE_EXT!'
    pause
    exit /b 1
)

REM Derive output filename (replace .docx with .pdf)
for %%i in ("!INPUT_FILE!") do set "OUTPUT_FILE=%%~dpni.pdf"

REM Get the directory of the batch script to look for insert_pdf.lua
set "SCRIPT_DIR=%~dp0"

echo =====================================
echo DOCX to PDF Report Compiler
echo =====================================
echo Input file:  !INPUT_FILE!
echo Output file: !OUTPUT_FILE!
echo Script directory: %SCRIPT_DIR%
echo.

REM Determine which template to use
set "PANDOC_TEMPLATE_FLAG="

if defined TEMPLATE_FILE (
    if exist "!TEMPLATE_FILE!" (
        echo Using specified template: !TEMPLATE_FILE!
        set "PANDOC_TEMPLATE_FLAG=--template="!TEMPLATE_FILE!""
    ) else (
        echo Error: Specified template file not found: !TEMPLATE_FILE!
        pause
        exit /b 1
    )
) else (
    echo Using Pandoc default template...
)

echo.
echo Running Pandoc...
echo.

REM Run pandoc with or without template
if defined PANDOC_TEMPLATE_FLAG (
    pandoc "!INPUT_FILE!" ^
           --lua-filter="%SCRIPT_DIR%insert_pdf.lua" ^
           --pdf-engine=pdflatex ^
           !PANDOC_TEMPLATE_FLAG! ^
           -o "!OUTPUT_FILE!"
) else (
    REM When using default template, add pdfpages package for includepdf support
    pandoc "!INPUT_FILE!" ^
           --lua-filter="%SCRIPT_DIR%insert_pdf.lua" ^
           --pdf-engine=pdflatex ^
           -H "%SCRIPT_DIR%pdfpages_header.tex" ^
           -o "!OUTPUT_FILE!"
)

REM Check if the pandoc command was successful
if !ERRORLEVEL! neq 0 (
    echo.
    echo ERROR: Pandoc compilation failed with error code !ERRORLEVEL!
    echo Please check that:
    echo - Pandoc is installed and in your PATH
    echo - The input .docx file is valid
    echo - All referenced PDF files exist
    echo - pdflatex is installed and available
    pause
    exit /b !ERRORLEVEL!
) else (
    echo.
    echo SUCCESS: PDF report compiled successfully!
    echo Output saved to: !OUTPUT_FILE!
)

echo.
echo Compilation complete. This window will close in 10 seconds...
timeout /t 10 /nobreak >nul
