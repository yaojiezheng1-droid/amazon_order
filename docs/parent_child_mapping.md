# Parent and Child Products

Some products have very similar names and share the same accessories. These relationships were identified from the spreadsheet `导出产品-按SKU-808264991961878528.xlsx` and are documented here so that accessory requirements can be grouped by parent product.

## JSON Structure

```json
{
  "parents": {
    "<parent_sku>": {
      "name": "<parent_name>",
      "children": ["<child_sku1>", "<child_sku2>", ...]
    },
    ...
  }
}
```

For example, the parent SKU `B10-MJB2` groups the two child SKUs `B10-MJB2-BK` and `B10-MJB2-BK2`. They use the same accessory SKU listed in `accessory_mapping.json`.

The actual relationships are stored in `parent_child_mapping.json` in this
directory. This file is generated from the spreadsheet and used by automation
scripts.

