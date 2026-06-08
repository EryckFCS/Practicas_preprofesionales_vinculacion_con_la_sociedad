import pandas as pd
import os

path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/26/satisfaccion.parquet"
df = pd.read_parquet(path)
for col in df.columns:
    print(f"--- Column {col} ---")
    print(df[col].tolist())
