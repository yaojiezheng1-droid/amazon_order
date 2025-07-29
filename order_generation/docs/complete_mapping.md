# Complete Product Mapping

The file `complete_mapping.json` merges information from
`full_accessory_mapping.json` and `parent_child_mapping.json`.
Each parent SKU lists all of its child SKUs along with any
accessories associated with those children.

This mapping is generated using `merge_mappings.py`:


python merge_mappings.py order_generation/docs/accessory_mapping.json order_generation/docs/parent_child_mapping.json order_generation/complete_mapping.json

```

The resulting JSON structure looks like:

```json
{
  "parents": {
    "<parent_sku>": {
      "name": "<parent_name>",
      "children": [
        {
          "sku": "<child_sku>",
          "name": "<child_name>",
          "accessories": [ { "sku": "<accessory_sku>", ... } ]
        },
        ...
      ]
    },
    ...
  }
}
```