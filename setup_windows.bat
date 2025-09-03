@echo off
echo Amazon Order Generation - Quick Setup for Windows
echo ================================================
echo.
echo This will install all required dependencies for:
echo - Product Search GUI
echo - Accessory Mapping Updater GUI  
echo - Advanced Excel to JSON Converter
echo - Direct SKU to JSON Converter
echo - JSON to Excel Converter
echo - Excel to JSON Template Converter
echo.

pause

echo Installing required packages...
python -m pip install --upgrade pip
python -m pip install pyperclip openpyxl pillow

echo.
echo Installation complete!
echo.
echo You can now run:
echo   python order_generation\product_search_gui.py
echo   python order_generation\accessory_mapping_updater_gui.py
echo   python order_generation\direct_sku_to_json.py
echo   python order_generation\json_PO_excel.py
echo   python order_generation\excel_to_json_template.py
echo.

pause
