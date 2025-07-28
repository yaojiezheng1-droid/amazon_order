# Accessory Mapping

This document describes the JSON file `accessory_mapping.json` that lists the accessory SKUs linked to product SKUs. The JSON was generated from the spreadsheet `导出产品-按SKU-808264991961878528.xlsx`.

## JSON Structure

```json
{
  "products": {
    "<product_sku>": {
      "name": "<product_name>",
      "accessories": [
        {
          "sku": "<accessory_sku>",
          "name": "<accessory_name>",
          "ratio_main": "<ratio main>",
          "ratio_accessory": "<ratio accessory>"
        },
        ...
      ]
    },
    ...
  }
}
```

For example, product SKU `B10-MJB2-BK` is linked to accessory SKU `B10-MJB2-BK-1` with equal ratios for the main and accessory items.

Use this JSON for programmatically looking up accessory requirements per product.
