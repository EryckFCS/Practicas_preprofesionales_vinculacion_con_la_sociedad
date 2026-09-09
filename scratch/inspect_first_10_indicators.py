import pandas as pd
df = pd.read_parquet("data/parquet/35/indicadores.parquet")
for idx, row in df.head(10).iterrows():
    print(f"Row {idx}:")
    for col in df.columns:
        val = row[col]
        if pd.notna(val) and str(val).strip():
            print(f"  {col}: {val}")
