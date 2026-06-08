import pandas as pd
import os

path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/26/satisfaccion.parquet"
if os.path.exists(path):
    df = pd.read_parquet(path)
    print("Columns:", df.columns.tolist())
    print("\nShape:", df.shape)
    print("\nData:")
    print(df.to_string())
else:
    print("File not found")
