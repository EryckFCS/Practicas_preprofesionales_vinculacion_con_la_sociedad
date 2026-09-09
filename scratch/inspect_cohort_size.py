import os
import openpyxl
from src.config import CAREER_FILES, RESULTADOS_DIR

for rid in range(35, 47):
    career = CAREER_FILES[rid]
    excel_path = os.path.join(RESULTADOS_DIR, career["excel"])
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    dem_sheet = wb['Demográfico']
    print(f"\n=== Report {rid} ({career['excel']}) ===")
    for r in range(1, 10):
        row_vals = [dem_sheet.cell(r, col).value for col in range(1, 8)]
        if any(v is not None for v in row_vals):
            print(f"Row {r}: {row_vals}")
