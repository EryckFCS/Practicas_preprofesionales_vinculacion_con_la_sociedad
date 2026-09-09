import pandas as pd
df = pd.read_parquet("data/parquet/35/indicadores.parquet")
print("Columns:", df.columns.tolist())
print(df.iloc[:10, :5].to_string())
