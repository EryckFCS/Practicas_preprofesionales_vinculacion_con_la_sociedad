import openpyxl

target_file = "/home/erick-fcs/Descargas/Practicas_preprofesionales/RESULTADOS-24-46 ERICK/25. IND-Derecho 2023.xlsx"
wb = openpyxl.load_workbook(target_file, data_only=True)
sheet = wb['Indicadores ']
print("--- Indicadores Sheet ---")
for r_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
    row_str = " | ".join(str(val) for val in row if val is not None)
    if row_str.strip():
        print(f"Row {r_idx}: {row_str}")
