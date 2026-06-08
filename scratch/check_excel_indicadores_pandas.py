import os
import pandas as pd
from src.config import CAREER_FILES, RESULTADOS_DIR

for rid in sorted(CAREER_FILES.keys()):
    if rid >= 35:
        path = os.path.join(RESULTADOS_DIR, CAREER_FILES[rid]["excel"])
        # read the sheet name that starts with 'Indicadores'
        xl = pd.ExcelFile(path)
        sheet_name = [s for s in xl.sheet_names if s.strip().startswith('Indicadores')][0]
        df = xl.parse(sheet_name)
        # find the 'Total :' row and print
        for idx, row in df.iterrows():
            row_vals = [str(x).strip() for x in row.values if pd.notna(x)]
            if any("total" in str(x).lower() for x in row.values):
                print(f"Report {rid}: sheet='{sheet_name}', row {idx}: {row_vals}")
