import os
import openpyxl
from src.config import CAREER_FILES, RESULTADOS_DIR

for rid in range(35, 47):
    career = CAREER_FILES[rid]
    excel_path = os.path.join(RESULTADOS_DIR, career["excel"])
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    print(f"\n=====================================")
    print(f"Report {rid}: {career['excel']}")
    print(f"=====================================")
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for r in range(1, 25): # scan first 25 rows
            for c in range(1, 15): # scan first 15 cols
                val = sheet.cell(r, c).value
                if val is not None:
                    s_val = str(val).lower()
                    if "poblaci" in s_val or "graduados" in s_val or "cohorte" in s_val or "total" in s_val or "muestra" in s_val:
                        # Print cell and surrounding values
                        row_vals = [sheet.cell(r, col).value for col in range(1, 10)]
                        print(f"Sheet '{sheet_name}', Row {r}: {row_vals}")
                        break
