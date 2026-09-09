import pandas as pd
df = pd.read_parquet("data/parquet/35/indicadores.parquet")
for idx, row in df.iterrows():
    v1 = str(row.get('variables_encuesta_a_graduados', '')).strip().lower()
    if "total" in v1:
        print(f"Row {idx}: {v1}")
        for col in df.columns:
            print(f"  {col}: {row[col]}")
