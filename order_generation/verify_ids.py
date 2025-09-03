#!/usr/bin/env python3
import pandas as pd

# Read the generated PO import file
df = pd.read_excel('PO_import_filled/PO_import_test_factory_ids.xlsx', header=1)

print('*标识号 and 采购单号 for each product:')
for i, row in df[['*标识号', '采购单号', '*SKU']].iterrows():
    print(f'  {row["*SKU"]}: 标识号={row["*标识号"]}, 采购单号={row["采购单号"]}')

print(f'\nUnique *标识号 values: {sorted(df["*标识号"].unique())}')
print(f'Total products: {len(df)}')
