import json
import hashlib
import re
from pathlib import Path
from legacy.excel_to_json import parse_order

OUTPUT_DIR = Path('json_exports')
OUTPUT_DIR.mkdir(exist_ok=True)


def slugify(filename: str) -> str:
    base = Path(filename).stem
    slug = re.sub(r'[^A-Za-z0-9_-]+', '_', base)
    digest = hashlib.md5(filename.encode('utf-8')).hexdigest()[:6]
    return f"{slug}_{digest}.json"


def main():
    for path in Path('.').rglob('*.xlsx'):
        out_name = slugify(path.name)
        out_path = OUTPUT_DIR / out_name
        data = parse_order(str(path))
        with open(out_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"converted {path} -> {out_path}")


if __name__ == '__main__':
    main()
