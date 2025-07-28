import json
import hashlib
import re
from pathlib import Path
from excel_to_json import parse_order

# Load accessory and parent-child mappings once at startup
with open('docs/accessory_mapping.json', 'r', encoding='utf-8') as f:
    _ACC_MAP = json.load(f)
_SKU_TO_NAME = {sku: info.get('name', '') for sku, info in _ACC_MAP.get('products', {}).items()}

with open('docs/parent_child_mapping.json', 'r', encoding='utf-8') as f:
    _PARENT_MAP = json.load(f)
_SKU_TO_PARENT = {}
for parent, info in _PARENT_MAP.get('parents', {}).items():
    for child in info.get('children', []):
        _SKU_TO_PARENT[child] = parent

OUTPUT_DIR = Path('json_exports')
OUTPUT_DIR.mkdir(exist_ok=True)


def slugify(name: str) -> str:
    slug = re.sub(r'[^A-Za-z0-9_-]+', '_', name)
    digest = hashlib.md5(name.encode('utf-8')).hexdigest()[:6]
    return f"{slug}_{digest}.json"


def _group_products(products):
    groups = {}
    for item in products:
        sku = str(item.get('产品编号', '')).strip()
        parent = _SKU_TO_PARENT.get(sku, sku)
        name = _SKU_TO_NAME.get(sku)
        if name:
            item['产品名称'] = name
        groups.setdefault(parent, []).append(item)
    return groups


def main():
    for path in Path('.').rglob('*.xlsx'):
        raw = parse_order(str(path))
        groups = _group_products(raw.get('products', []))

        if len(groups) <= 1:
            out_name = slugify(Path(path.name).stem)
            out_path = OUTPUT_DIR / out_name
            raw['products'] = next(iter(groups.values())) if groups else []
            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(raw, f, ensure_ascii=False, indent=2)
            print(f"converted {path} -> {out_path}")
        else:
            for parent, items in groups.items():
                data = {
                    'cells': raw.get('cells', {}),
                    'products': items,
                    'footer': raw.get('footer', {}),
                }
                out_name = slugify(f"{Path(path.name).stem}_{parent}")
                out_path = OUTPUT_DIR / out_name
                with open(out_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                print(f"converted {path} ({parent}) -> {out_path}")


if __name__ == '__main__':
    main()
