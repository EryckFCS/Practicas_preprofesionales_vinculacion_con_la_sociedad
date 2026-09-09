import os
import pandas as pd

for rid in range(35, 47):
    path = f"data/parquet/{rid}/demografico.parquet"
    if os.path.exists(path):
        df = pd.read_parquet(path)
        print(f"\n=== Report {rid} ===")
        print(df.to_string())
    else:
        print(f"Report {rid} parquet not found")
