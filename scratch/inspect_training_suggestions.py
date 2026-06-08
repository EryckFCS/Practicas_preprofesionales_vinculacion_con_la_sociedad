import pandas as pd
import os

path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/25/pregunta_abierta_temas_de_capacitacion_sugerida_por_los_empleadores.parquet"
if os.path.exists(path):
    df = pd.read_parquet(path)
    for col in df.columns[2:]:
        print(f"\n--- COLUMN: {col} ---")
        for idx, val in enumerate(df[col].dropna().tolist()):
            print(f"{idx}: {repr(val)}")
else:
    print("File not found")
