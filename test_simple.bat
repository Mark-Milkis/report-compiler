@echo off
setlocal enabledelayedexpansion

echo DEBUG: First argument: "%1"
echo DEBUG: All arguments: %*

set "INPUT_FILE=%1"
echo DEBUG: INPUT_FILE set to: "!INPUT_FILE!"

if "!INPUT_FILE!"=="" (
    echo ERROR: No input file provided
    pause
    exit /b 1
)

echo SUCCESS: Input file is: !INPUT_FILE!
pause
