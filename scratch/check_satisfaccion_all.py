import os
import pandas as pd

for report_id in range(24, 35):
    parquet_path = f"/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/{report_id}/satisfaccion.parquet"
    if os.path.exists(parquet_path):
        df = pd.read_parquet(parquet_path)
        # Search for row describing "2) Percepción sobre el perfil de egreso" or "perfil de egreso"
        # and print values for Buena, Regular, Mala under "Pertinencia de los resultados de aprendizaje" or similar
        print(f"\n================= REPORT {report_id} =================")
        print("Shape:", df.shape)
        # Just search for "egres" or "perfil" or "enseñanza"
        for idx, row in df.iterrows():
            vals = [str(x) for x in row.values]
            if any("egres" in str(v).lower() or "enseñ" in str(v).lower() for v in vals):
                print(f"Row {idx}: {vals[:3]}")
    else:
        print(f"Report {report_id}: No satisfaccion.parquet")
