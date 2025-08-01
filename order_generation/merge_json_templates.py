import json
import sys
from pathlib import Path
from typing import Any, Dict, List


def _select_cell(existing: Dict[str, Any], new: Dict[str, Any], addr: str) -> Dict[str, Any]:
    """Return preferred cell info between ``existing`` and ``new``.

    Preference is given to the first non-empty value encountered. When both
    values are present and conflict, the original value is kept and a message
    is printed to stderr so the caller can review the decision.
    """
    if not existing or not existing.get("value"):
        return new
    if new.get("value") and new["value"] != existing.get("value"):
        print(
            f"conflicting value for cell {addr}: "
            f"keeping {existing['value']!r} and ignoring {new['value']!r}",
            file=sys.stderr,
        )
    return existing


def merge_json_templates(paths: List[Path]) -> Dict[str, Any]:
    """Merge product JSON templates for a single factory.

    The first file acts as the base. Subsequent files have their ``products``
    appended and their ``cells`` merged using :func:`_select_cell` to choose the
    most appropriate value when conflicts arise.
    """
    if not paths:
        raise ValueError("at least one input path is required")

    with open(paths[0], "r", encoding="utf-8") as f:
        merged: Dict[str, Any] = json.load(f)

    merged.setdefault("products", [])
    merged.setdefault("cells", {})
    merged.setdefault("footer", {})

    for path in paths[1:]:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        merged["products"].extend(data.get("products", []))

        cells = data.get("cells", {})
        for addr, info in cells.items():
            merged["cells"][addr] = _select_cell(merged["cells"].get(addr), info, addr)

        footer = data.get("footer", {})
        for key, value in footer.items():
            if key not in merged["footer"] or not merged["footer"][key]:
                merged["footer"][key] = value

    return merged


def main(argv: List[str]) -> int:
    if len(argv) < 4:
        print(
            "usage: merge_json_templates.py <output.json> <input1.json> <input2.json> [input3.json ...]"
        )
        return 1

    out_path = Path(argv[1])
    in_paths = [Path(p) for p in argv[2:]]

    merged = merge_json_templates(in_paths)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, indent=2)

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
