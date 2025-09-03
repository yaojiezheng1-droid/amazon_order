@echo off
REM Batch file to convert JSON files to PO import Excel format
REM Usage: convert_json_to_po_import.bat [order_name]
REM If order_name is provided, it will process files named {order_name}-*.json
REM If no order_name is provided, it will process all JSON files

setlocal enabledelayedexpansion

if "%~1"=="" (
    echo Converting all JSON files to PO import format...
    python fill_po_import.py
) else (
    echo Converting JSON files for order "%~1" to PO import format...
    python fill_po_import.py "%~1"
)

if %errorlevel% equ 0 (
    echo.
    echo PO import Excel file created successfully!
    echo Check the PO_import_filled directory for the output file.
    pause
) else (
    echo.
    echo Error occurred during conversion. Please check the output above.
    pause
)
