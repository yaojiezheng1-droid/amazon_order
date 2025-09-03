@echo off
echo JSON Templates to Excel Converter
echo ==================================
echo.
echo This script converts all JSON template files to Excel format
echo using the empty_base_template.xlsx structure.
echo.

cd /d "c:\Users\Cheng\Desktop\amazon_order\order_generation"

if "%~1"=="--help" (
    echo Usage: convert_templates_to_excel.bat [options]
    echo.
    echo Options:
    echo   --help     Show this help message
    echo   --list     List available JSON templates
    echo   --all      Convert all templates ^(default^)
    echo.
    echo Examples:
    echo   convert_templates_to_excel.bat
    echo   convert_templates_to_excel.bat --list
    echo.
    pause
    exit /b 0
)

if "%~1"=="--list" (
    echo Listing available JSON templates...
    C:/Users/Cheng/AppData/Local/Programs/Python/Python310/python.exe json_templates_to_excel.py --list
    echo.
    pause
    exit /b 0
)

echo Starting conversion of all JSON templates...
echo.
C:/Users/Cheng/AppData/Local/Programs/Python/Python310/python.exe json_templates_to_excel.py

echo.
echo Conversion completed!
echo Excel files are saved in the PO_excel directory.
echo.
pause
