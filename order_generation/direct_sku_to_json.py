#!/usr/bin/env python3
"""Generate factory-grouped JSON templates from SKU/quantity pairs.

This script automates steps 1, 2, and 4 of the "Handling Direct SKU Requests"
section of the README. Provide pairs of `<sku> <quantity>` on the command line
and it will:

* look up accessory ratios in ``docs/complete_mapping.json``;
* set the ``数量/个`` field in each product template;
* group templates by supplier (cell ``B3``) and merge items from the same
  factory;
* write the merged JSON files into ``json_exports/``.

Example
-------
```
python direct_sku_to_json.py 48-82P3-QSFG 800 Elasticbrush01 500
```

The resulting JSON files can then be converted to Excel using
``json_PO_excel.py``.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List

from merge_json_templates import merge_json_templates


ROOT = Path(__file__).resolve().parent
TEMPLATE_DIR = ROOT / "json_template"
MAPPING_PATH = ROOT / "docs" / "complete_mapping.json"
OUTPUT_DIR = ROOT / "json_exports"
EXCEL_OUTPUT_DIR = ROOT / "PO_excel_export"


def _load_accessory_mapping() -> Dict[str, List[dict]]:
    with open(MAPPING_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    lookup: Dict[str, List[dict]] = {}
    for parent in data.get("parents", {}).values():
        for child in parent.get("children", []):
            lookup[child["sku"]] = child.get("accessories", [])
    return lookup


def _compute_all_items(requests: Dict[str, int], mapping: Dict[str, List[dict]]) -> Dict[str, int]:
    result: Dict[str, int] = {}
    for sku, qty in requests.items():
        result[sku] = result.get(sku, 0) + qty
        for acc in mapping.get(sku, []):
            try:
                main = int(acc.get("ratio_main", 1))
                accessory = int(acc.get("ratio_accessory", 1))
            except ValueError:
                main = accessory = 1
            acc_qty = qty * accessory // main
            acc_sku = acc.get("sku")
            if acc_sku:
                result[acc_sku] = result.get(acc_sku, 0) + acc_qty
    return result


def _sanitize(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")
    return safe or "factory"


def _run_json_to_excel(json_path: Path) -> Path:
    """Convert JSON file to Excel using json_PO_excel.py"""
    EXCEL_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    excel_filename = json_path.stem + ".xlsx"
    excel_path = EXCEL_OUTPUT_DIR / excel_filename
    
    json_po_excel_script = ROOT / "json_PO_excel.py"
    
    try:
        # Run json_PO_excel.py as a subprocess
        result = subprocess.run([
            sys.executable, str(json_po_excel_script), 
            str(json_path), str(excel_path)
        ], check=True, capture_output=True, text=True)
        
        print(f"Generated Excel file: {excel_path}")
        return excel_path
        
    except subprocess.CalledProcessError as e:
        print(f"Error converting {json_path} to Excel: {e}")
        print(f"stdout: {e.stdout}")
        print(f"stderr: {e.stderr}")
        raise
    except Exception as e:
        print(f"Unexpected error converting {json_path} to Excel: {e}")
        raise


def generate_factory_jsons(pairs: Dict[str, int]) -> List[Path]:
    mapping = _load_accessory_mapping()
    all_items = _compute_all_items(pairs, mapping)
    temp_files: Dict[str, List[Path]] = {}

    for sku, qty in all_items.items():
        template_path = TEMPLATE_DIR / f"{sku}.json"
        if not template_path.exists():
            print(f"warning: template for {sku} not found", flush=True)
            continue
        with open(template_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        for product in data.get("products", []):
            product["数量/个"] = qty
        factory = data.get("cells", {}).get("B3", {}).get("value", "factory")
        tmp = tempfile.NamedTemporaryFile("w", delete=False, encoding="utf-8", suffix=".json")
        json.dump(data, tmp, ensure_ascii=False, indent=2)
        tmp.close()
        temp_files.setdefault(factory, []).append(Path(tmp.name))

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out_paths: List[Path] = []
    excel_paths: List[Path] = []
    factory_counter = 1
    
    for factory, paths in temp_files.items():
        merged = merge_json_templates(paths)
        out_path = OUTPUT_DIR / f"factory{factory_counter}.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        out_paths.append(out_path)
        
        # Automatically convert JSON to Excel
        try:
            excel_path = _run_json_to_excel(out_path)
            excel_paths.append(excel_path)
        except Exception as e:
            print(f"Failed to convert {out_path} to Excel: {e}")
        
        factory_counter += 1
        for p in paths:
            try:
                Path(p).unlink()
            except OSError:
                pass
    
    # Print summary
    print(f"\nGenerated {len(out_paths)} JSON files and {len(excel_paths)} Excel files:")
    for json_path, excel_path in zip(out_paths, excel_paths):
        print(f"  JSON: {json_path}")
        print(f"  Excel: {excel_path}")
    
    return out_paths


def parse_args(argv: List[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("items", nargs="+", help="Pairs of <sku> <quantity>")
    return parser.parse_args(argv)


def main(argv: List[str] | None = None) -> int:
    ns = parse_args(argv)
    if len(ns.items) % 2:
        print("error: expected even number of arguments", flush=True)
        return 1
    requests = {ns.items[i]: int(ns.items[i + 1]) for i in range(0, len(ns.items), 2)}
    paths = generate_factory_jsons(requests)
    for p in paths:
        print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

# Execution code for 1200 pieces EEHB-NBB
if __name__ == "__main__":
    # Generate JSON for 1200 pieces of EEHB-NBB
    sku_quantity_pairs = {"EEHB-NBB": 1200}
    
    print(f"Generating factory JSON files for {sku_quantity_pairs}")
    
    try:
        output_paths = generate_factory_jsons(sku_quantity_pairs)
        
        print("\nGenerated factory JSON files:")
        for path in output_paths:
            print(f"  - {path}")
            
        print(f"\nSuccessfully processed order for 1200 pieces of EEHB-NBB")
        print(f"Output files saved to: {OUTPUT_DIR}")
        
    except Exception as e:
        print(f"Error generating JSON files: {e}")
        raise SystemExit(1)
