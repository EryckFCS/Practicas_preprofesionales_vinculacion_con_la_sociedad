import os
import openpyxl
from src.config import CAREER_FILES, RESULTADOS_DIR

excel_path = os.path.join(RESULTADOS_DIR, CAREER_FILES[35]["excel"])
wb = openpyxl.load_workbook(excel_path, data_only=True)
sheet = wb['Indicadores ']
for r_idx in range(1, 40):
    vals = [sheet.cell(r_idx, col).value for col in range(1, 15)]
    if any(v is not None for v in vals):
        print(f"Row {r_idx:2d}: {vals[:8]}")
