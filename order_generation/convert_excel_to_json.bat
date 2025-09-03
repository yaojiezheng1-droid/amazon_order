@echo off
echo Excel to JSON Template Converter
echo ================================
echo.
echo This script converts Excel files in the empty_base_template.xlsx format
echo to JSON template files for use with the order generation system.
echo.

if "%~1"=="" (
    echo Usage: convert_excel_to_json.bat ^<excel_file^>
    echo        convert_excel_to_json.bat *.xlsx
    echo.
    echo Examples:
    echo   convert_excel_to_json.bat my_order.xlsx
    echo   convert_excel_to_json.bat docs\*.xlsx
    echo.
    pause
    exit /b 1
)

cd /d "c:\Users\Cheng\Desktop\amazon_order\order_generation"
C:/Users/Cheng/AppData/Local/Programs/Python/Python310/python.exe excel_to_json_template.py %*

echo.
echo Conversion completed!
pause
