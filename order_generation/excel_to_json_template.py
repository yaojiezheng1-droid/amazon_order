#!/usr/bin/env python3
"""
Excel to JSON Template Generator

This script converts Excel files in the format of empty_base_template.xlsx
to JSON template files suitable for the json_template folder.

The script:
1. Reads Excel files with the standard template format
2. Extracts cell data, product information, and footer data
3. Generates individual JSON template files for each product SKU
4. Saves files to the json_template directory with proper formatting

Usage:
    python excel_to_json_template.py input_file.xlsx
    python excel_to_json_template.py *.xlsx  # Process multiple files
"""

import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:
    print("Error: openpyxl is required. Install it with: pip install openpyxl")
    sys.exit(1)


class ExcelToJsonConverter:
    def __init__(self):
        self.root_dir = Path(__file__).resolve().parent
        self.template_dir = self.root_dir / "json_template"
        self.images_dir = self.root_dir / "images"
        
        # Create template directory if it doesn't exist
        self.template_dir.mkdir(exist_ok=True)
        
        # Load accessory mapping for product names
        self.accessory_map = self._load_accessory_mapping()
        
    def _load_accessory_mapping(self) -> Dict[str, Dict]:
        """Load accessory mapping to get product names"""
        mapping_path = self.root_dir / "docs" / "accessory_mapping.json"
        try:
            with open(mapping_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("products", {})
        except FileNotFoundError:
            print(f"Warning: {mapping_path} not found. Product names may not be populated.")
            return {}
    
    def _find_image_path(self, sku: str) -> Optional[str]:
        """Find image path for the given SKU"""
        for sub in ("products", "accessories"):
            img_path = self.images_dir / sub / f"{sku}.jpg"
            if img_path.exists():
                return f"order_generation/images/{sub}/{sku}.jpg"
        return None
    
    def _get_cell_value(self, ws, row: int, col: int) -> str:
        """Get cell value as string"""
        cell = ws.cell(row=row, column=col)
        return str(cell.value or "").strip()
    
    def _extract_cells_data(self, ws) -> Dict[str, Dict[str, str]]:
        """Extract all cell data with keys and values"""
        cells = {}
        
        # Define the cell mapping based on the template structure
        # These are the standard cells that contain metadata
        cell_mappings = {
            # Row 3
            "B3": {"key": "供货商：", "row": 3, "col": 2},
            "G3": {"key": "订单号", "row": 3, "col": 7},
            
            # Row 4
            "B4": {"key": "电话：", "row": 4, "col": 2},
            "G4": {"key": "日期", "row": 4, "col": 7},
            
            # Row 5
            "B5": {"key": "联系人：", "row": 5, "col": 2},
            "G5": {"key": "订单安排人", "row": 5, "col": 7},
            
            # Row 12-17 (various fields)
            "B12": {"key": "进仓地址：", "row": 12, "col": 2},
            "B13": {"key": "付款方式", "row": 13, "col": 2},
            "B14": {"key": "交货时间", "row": 14, "col": 2},
            "F14": {"key": "色卡", "row": 14, "col": 6},
            "G14": {"key": "Logo", "row": 14, "col": 7},
            "B15": {"key": "箱规", "row": 15, "col": 2},
            "F15": {"key": "色卡1", "row": 15, "col": 6},
            "G15": {"key": "Logo1", "row": 15, "col": 7},
            "B16": {"key": "产前确认样", "row": 16, "col": 2},
            "F16": {"key": "色卡2", "row": 16, "col": 6},
            "G16": {"key": "Logo2", "row": 16, "col": 7},
            "B17": {"key": "出货样", "row": 17, "col": 2},
            "F17": {"key": "色卡3", "row": 17, "col": 6},
            "G17": {"key": "Logo3", "row": 17, "col": 7},
            "F18": {"key": "色卡4", "row": 18, "col": 6},
            "G18": {"key": "Logo4", "row": 18, "col": 7},
        }
        
        # Add note fields (A19-A30)
        for i in range(19, 31):
            if i == 19:
                key = "注意事项：1"
            else:
                key = f"{i-18}："
            cell_mappings[f"A{i}"] = {"key": key, "row": i, "col": 1}
        
        # Extract values for mapped cells
        for addr, info in cell_mappings.items():
            value = self._get_cell_value(ws, info["row"], info["col"])
            cells[addr] = {
                "key": info["key"],
                "value": value
            }
        
        return cells
    
    def _find_product_table_start(self, ws) -> int:
        """Find the row where the product table starts"""
        # Look for the header row containing "产品编号", "数量/个", etc.
        for row in range(1, 20):  # Check first 20 rows
            for col in range(1, 8):  # Check columns A-G
                cell_value = self._get_cell_value(ws, row, col)
                if cell_value in ("产品编号", "型号"):
                    return row
        return 7  # Default to row 7 if not found
    
    def _extract_products(self, ws) -> List[Dict[str, Any]]:
        """Extract product data from the worksheet"""
        products = []
        header_row = self._find_product_table_start(ws)
        
        # Define column mapping for product table
        # Based on standard template: A=产品编号, B=产品图片, C=描述, D=数量/个, E=单价, G=包装方式
        col_mapping = {
            1: "产品编号",     # Column A
            2: "产品图片",     # Column B  
            3: "描述",         # Column C
            4: "数量/个",      # Column D
            5: "单价",         # Column E
            7: "包装方式"      # Column G
        }
        
        # Start from the row after header
        row = header_row + 1
        
        while row <= ws.max_row:
            # Get product code from column A
            sku = self._get_cell_value(ws, row, 1)
            
            # Stop if we hit an empty SKU or total row
            if not sku or sku.upper().startswith("TOTAL") or sku == "总计":
                break
            
            product = {}
            
            # Extract data for each column
            for col, field in col_mapping.items():
                value = self._get_cell_value(ws, row, col)
                
                if field == "产品编号":
                    product[field] = value
                elif field == "产品图片":
                    # Try to find image path, use provided value as fallback
                    img_path = self._find_image_path(sku)
                    product[field] = img_path or value
                elif field == "数量/个":
                    # Convert to integer, default to 0
                    try:
                        product[field] = int(float(value)) if value else 0
                    except ValueError:
                        product[field] = 0
                elif field == "单价":
                    # Convert to float
                    try:
                        product[field] = float(value) if value else 0.0
                    except ValueError:
                        product[field] = 0.0
                else:
                    product[field] = value
            
            # Add product name from accessory mapping if available
            if sku in self.accessory_map:
                product["产品名称"] = self.accessory_map[sku]["name"]
            elif "产品名称" not in product:
                product["产品名称"] = ""  # Default empty name
            
            products.append(product)
            row += 1
        
        return products
    
    def _extract_footer(self, ws) -> Dict[str, str]:
        """Extract footer information (buyer, supplier)"""
        footer = {}
        
        # Look for buyer and supplier info around row 69 (standard template)
        try:
            buyer = self._get_cell_value(ws, 69, 2)  # B69
            supplier = self._get_cell_value(ws, 69, 5)  # E69
            
            if buyer:
                footer["buyer"] = buyer
            if supplier:
                footer["supplier"] = supplier
        except:
            pass
        
        return footer
    
    def convert_excel_to_json(self, excel_path: Path) -> List[Path]:
        """Convert Excel file to JSON template(s)"""
        print(f"Processing: {excel_path}")
        
        try:
            wb = load_workbook(excel_path, data_only=True)
            ws = wb.active
            
            # Extract data
            cells = self._extract_cells_data(ws)
            products = self._extract_products(ws)
            footer = self._extract_footer(ws)
            
            if not products:
                print(f"Warning: No products found in {excel_path}")
                return []
            
            generated_files = []
            
            # Group products by SKU to avoid duplicates
            products_by_sku = {}
            for product in products:
                sku = product.get("产品编号")
                if sku:
                    if sku not in products_by_sku:
                        products_by_sku[sku] = product
                    else:
                        # If duplicate SKU, combine quantities
                        existing = products_by_sku[sku]
                        existing["数量/个"] += product.get("数量/个", 0)
            
            # Generate one JSON file per unique product SKU
            for sku, product in products_by_sku.items():
                # Create JSON structure
                json_data = {
                    "cells": cells,
                    "products": [product],
                    "footer": footer
                }
                
                # Save to json_template directory
                output_path = self.template_dir / f"{sku}.json"
                
                with open(output_path, "w", encoding="utf-8") as f:
                    json.dump(json_data, f, ensure_ascii=False, indent=2)
                
                generated_files.append(output_path)
                print(f"  Generated: {output_path}")
            
            return generated_files
            
        except Exception as e:
            print(f"Error processing {excel_path}: {e}")
            return []
    
    def process_files(self, file_patterns: List[str]) -> None:
        """Process multiple Excel files"""
        total_generated = 0
        
        for pattern in file_patterns:
            # Handle both specific files and glob patterns
            if "*" in pattern:
                files = list(Path(".").glob(pattern))
            else:
                files = [Path(pattern)]
            
            for file_path in files:
                if file_path.suffix.lower() in (".xlsx", ".xls"):
                    generated = self.convert_excel_to_json(file_path)
                    total_generated += len(generated)
        
        print(f"\nTotal JSON templates generated: {total_generated}")
        print(f"Output directory: {self.template_dir}")


def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_to_json_template.py <excel_file1> [excel_file2] ...")
        print("       python excel_to_json_template.py *.xlsx")
        sys.exit(1)
    
    converter = ExcelToJsonConverter()
    converter.process_files(sys.argv[1:])


if __name__ == "__main__":
    main()
