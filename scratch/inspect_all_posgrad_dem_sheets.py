import os
import openpyxl
from src.config import CAREER_FILES, RESULTADOS_DIR

for rid in range(35, 47):
    career = CAREER_FILES[rid]
    excel_path = os.path.join(RESULTADOS_DIR, career["excel"])
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb['Demográfico']
    # Print the first row with cohorte info
    r2_vals = [sheet.cell(2, col).value for col in range(1, 5)]
    print(f"Report {rid} Demográfico Row 2: {r2_vals}")
