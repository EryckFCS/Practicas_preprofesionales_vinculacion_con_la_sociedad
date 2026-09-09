import os
import pandas as pd

p_dir = "data/parquet/35"
for file in sorted(os.listdir(p_dir)):
    if file.endswith(".parquet"):
        df = pd.read_parquet(os.path.join(p_dir, file))
        print(f"{file:30} columns={len(df.columns)} rows={len(df)}")
