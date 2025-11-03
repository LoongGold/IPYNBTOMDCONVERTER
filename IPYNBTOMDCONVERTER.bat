@echo off
chcp 65001 >nul

if "%~1"=="" (
    echo Drag .ipynb files onto this batch file to convert to .md
    pause
    exit /b
)

if not exist "%~1" (
    echo File not found: %~1
    pause
    exit /b
)

if /i not "%~x1"==".ipynb" (
    echo Only .ipynb files are supported
    pause
    exit /b
)

echo Converting %~1 to Markdown...
jupyter nbconvert --to markdown "%~1"

if errorlevel 1 (
    echo Conversion failed! Make sure Jupyter is installed.
    echo Install with: pip install jupyter
) else (
    echo Conversion successful!
)

pause