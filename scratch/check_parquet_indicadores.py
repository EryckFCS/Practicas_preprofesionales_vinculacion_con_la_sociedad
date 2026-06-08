import os
import pandas as pd
from src.config import CAREER_FILES

for rid in sorted(CAREER_FILES.keys()):
    if rid >= 35:
        path = f"data/parquet/{rid}/indicadores.parquet"
        if os.path.exists(path):
            df = pd.read_parquet(path)
            for idx, row in df.iterrows():
                # find row with 'total' in variables_encuesta_a_graduados
                col_grad = "variables_encuesta_a_graduados"
                if col_grad in df.columns:
                    val = str(row[col_grad]).strip().lower()
                    if "total" in val:
                        row_vals = {c: row[c] for c in df.columns if pd.notna(row[c]) and str(row[c]).strip() != ""}
                        print(f"Report {rid}: {row_vals}")
                        break
