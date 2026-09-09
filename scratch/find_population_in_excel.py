import os
import openpyxl
from src.config import CAREER_FILES, RESULTADOS_DIR

expected_poblaciones = {
    35: 14,
    36: 16,
    37: 15,
    38: 20,
    39: 24,
    40: 24,
    41: 13,
    42: 14,
    43: 22,
    44: 35,
    45: 19,
    46: 64
}

for rid, expected in expected_poblaciones.items():
    career = CAREER_FILES[rid]
    excel_path = os.path.join(RESULTADOS_DIR, career["excel"])
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    found = []
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            for c_idx, val in enumerate(row, 1):
                if val == expected:
                    # check if surrounding cells have 'poblacion' or 'graduados' or similar
                    surroundings = []
                    for dr in [-1, 0, 1]:
                        for dc in [-1, 0, 1]:
                            if dr == 0 and dc == 0:
                                continue
                            nr, nc = r_idx + dr, c_idx + dc
                            if 1 <= nr <= sheet.max_row and 1 <= nc <= sheet.max_column:
                                s_val = sheet.cell(nr, nc).value
                                if s_val is not None:
                                    surroundings.append(str(s_val))
                    found.append((sheet_name, r_idx, c_idx, surroundings))
    print(f"\n=== Report {rid} ({career['excel']}) - Expected Population {expected} ===")
    if not found:
        print("Not found in any sheet!")
    for f in found[:5]: # print first 5 matches
        print(f"Sheet: {f[0]}, Cell: ({f[1]}, {f[2]}), Context: {f[3]}")
