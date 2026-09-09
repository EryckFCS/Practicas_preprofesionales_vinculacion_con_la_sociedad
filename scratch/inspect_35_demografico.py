import pandas as pd
df = pd.read_parquet("data/parquet/35/demografico.parquet")
print(df.to_string())
