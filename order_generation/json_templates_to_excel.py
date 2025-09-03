#!/usr/bin/env python3
"""
JSON Template to Excel Batch Converter

This script converts all JSON template files from the json_template directory
to Excel files in the format of empty_base_template.xlsx and saves them
under the PO_excel directory.

The script uses the existing json_PO_excel.py functionality to ensure
consistency with the current order generation system.

Usage:
    python json_templates_to_excel.py
    python json_templates_to_excel.py --output-dir custom_output
"""

import argparse
import subprocess
import sys
from pathlib import Path
from typing import List


class JsonTemplateToExcelConverter:
    def __init__(self, output_dir: str = "PO_excel"):
        self.root_dir = Path(__file__).resolve().parent
        self.template_dir = self.root_dir / "json_template"
        self.output_dir = self.root_dir / output_dir
        self.json_po_excel_script = self.root_dir / "json_PO_excel.py"
        
        # Create output directory if it doesn't exist
        self.output_dir.mkdir(exist_ok=True)
        
        # Verify json_PO_excel.py exists
        if not self.json_po_excel_script.exists():
            raise FileNotFoundError(f"Required script not found: {self.json_po_excel_script}")
    
    def get_json_templates(self) -> List[Path]:
        """Get all JSON template files"""
        if not self.template_dir.exists():
            raise FileNotFoundError(f"Template directory not found: {self.template_dir}")
        
        json_files = list(self.template_dir.glob("*.json"))
        return sorted(json_files)
    
    def convert_json_to_excel(self, json_path: Path) -> Path:
        """Convert a single JSON template to Excel using json_PO_excel.py"""
        # Generate output Excel filename based on JSON filename
        excel_filename = json_path.stem + ".xlsx"
        excel_path = self.output_dir / excel_filename
        
        try:
            # Run json_PO_excel.py as subprocess
            result = subprocess.run([
                sys.executable, str(self.json_po_excel_script),
                str(json_path), str(excel_path)
            ], check=True, capture_output=True, text=True)
            
            return excel_path
            
        except subprocess.CalledProcessError as e:
            print(f"Error converting {json_path.name}: {e}")
            print(f"stdout: {e.stdout}")
            print(f"stderr: {e.stderr}")
            raise
        except Exception as e:
            print(f"Unexpected error converting {json_path.name}: {e}")
            raise
    
    def convert_all_templates(self) -> List[Path]:
        """Convert all JSON templates to Excel files"""
        json_files = self.get_json_templates()
        
        if not json_files:
            print(f"No JSON template files found in {self.template_dir}")
            return []
        
        print(f"Found {len(json_files)} JSON template files")
        print(f"Converting to Excel format in: {self.output_dir}")
        print("-" * 50)
        
        converted_files = []
        failed_files = []
        
        for json_path in json_files:
            try:
                excel_path = self.convert_json_to_excel(json_path)
                converted_files.append(excel_path)
                print(f"✓ {json_path.name} -> {excel_path.name}")
                
            except Exception as e:
                failed_files.append(json_path)
                print(f"✗ {json_path.name} -> FAILED: {e}")
        
        # Print summary
        print("-" * 50)
        print(f"Conversion Summary:")
        print(f"  Successfully converted: {len(converted_files)}")
        print(f"  Failed conversions: {len(failed_files)}")
        print(f"  Output directory: {self.output_dir}")
        
        if failed_files:
            print(f"\nFailed files:")
            for failed_file in failed_files:
                print(f"  - {failed_file.name}")
        
        return converted_files
    
    def convert_specific_templates(self, template_names: List[str]) -> List[Path]:
        """Convert specific JSON templates by name"""
        converted_files = []
        
        for template_name in template_names:
            # Add .json extension if not present
            if not template_name.endswith('.json'):
                template_name += '.json'
            
            json_path = self.template_dir / template_name
            
            if not json_path.exists():
                print(f"Template not found: {json_path}")
                continue
            
            try:
                excel_path = self.convert_json_to_excel(json_path)
                converted_files.append(excel_path)
                print(f"✓ {json_path.name} -> {excel_path.name}")
                
            except Exception as e:
                print(f"✗ {json_path.name} -> FAILED: {e}")
        
        return converted_files


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir", "-o",
        default="PO_excel",
        help="Output directory for Excel files (default: PO_excel)"
    )
    parser.add_argument(
        "--templates", "-t",
        nargs="*",
        help="Specific template names to convert (without .json extension). If not provided, converts all templates."
    )
    parser.add_argument(
        "--list", "-l",
        action="store_true",
        help="List available JSON templates and exit"
    )
    
    args = parser.parse_args()
    
    try:
        converter = JsonTemplateToExcelConverter(args.output_dir)
        
        if args.list:
            templates = converter.get_json_templates()
            print(f"Available JSON templates ({len(templates)}):")
            for template in templates:
                print(f"  - {template.stem}")
            return
        
        if args.templates:
            # Convert specific templates
            print(f"Converting specific templates: {', '.join(args.templates)}")
            converted_files = converter.convert_specific_templates(args.templates)
        else:
            # Convert all templates
            print("Converting all JSON templates to Excel format...")
            converted_files = converter.convert_all_templates()
        
        if converted_files:
            print(f"\nSuccessfully converted {len(converted_files)} files!")
        else:
            print("\nNo files were converted.")
    
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\nConversion interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
