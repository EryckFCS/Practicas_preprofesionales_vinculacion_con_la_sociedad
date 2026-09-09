import os
import pandas as pd

for rid in [24, 25]:
    path = f"data/parquet/{rid}/demografico.parquet"
    if os.path.exists(path):
        df = pd.read_parquet(path)
        print(f"\n=== Report {rid} ===")
        print(df.head(5).to_string())
