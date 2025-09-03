#!/usr/bin/env python3
import pandas as pd

# Load the generated PO import file
df = pd.read_excel('PO_import_filled/PO_import_test_order.xlsx', header=1)

print('Shape:', df.shape)
print('\nRequired fields filled:')
required_cols = ['*标识号', '采购单号', '*供应商', '*含税', '*费用分配方式', '*采购币种', '*采购仓库', '*SKU', '*实际采购量', '*含税单价']

for col in required_cols:
    filled = not df[col].isna().all()
    print(f'{col}: {filled}')

print('\nFirst row sample:')
first_row = df[required_cols].iloc[0].to_dict()
for k, v in first_row.items():
    print(f'{k}: {v}')

print(f'\nTotal products: {len(df)}')
print('\nAll SKUs:')
for i, sku in enumerate(df['*SKU']):
    print(f'{i+1}. {sku}')
