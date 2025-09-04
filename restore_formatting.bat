@echo off
echo ==========================================
echo Excel Formatting Restoration Tool
echo ==========================================
echo.
echo This will restore Excel formatting while preserving your content changes.
echo.
pause

cd /d "C:\Users\Cheng\Desktop\amazon_order"

echo Running Python restoration script...
"C:/Users/Cheng/AppData/Local/Programs/Python/Python310/python.exe" restore_excel_formatting.py

echo.
echo ==========================================
echo Restoration complete!
echo.
echo Your files with restored formatting are in:
echo C:\Users\Cheng\Desktop\amazon_order\order_generation\PO_excel_restored\
echo ==========================================
pause
