import pandas as pd
df = pd.read_parquet("data/parquet/35/indicadores.parquet")
print(df.head(20).to_string())
