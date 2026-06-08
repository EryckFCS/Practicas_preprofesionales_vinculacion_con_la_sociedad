import os
import pandas as pd

parquet_dir = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/25"
for file in os.listdir(parquet_dir):
    if file.endswith(".parquet"):
        df = pd.read_parquet(os.path.join(parquet_dir, file))
        for col in df.columns:
            # Check if 89 is in any row
            try:
                matches = df[df[col].astype(str) == "89"]
                if not matches.empty:
                    print(f"Found 89 in {file}, column {col}")
            except Exception:
                pass
