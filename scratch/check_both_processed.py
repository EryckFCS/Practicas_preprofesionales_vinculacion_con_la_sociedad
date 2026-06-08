import docx
import os

paths = [
    "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/25.  INFORME SG DERECHO 2023.docx",
    "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/25.  INFORME SG DERECHO 2023.docx"
]

for p in paths:
    print(f"--- Path: {p} ---")
    if os.path.exists(p):
        doc = docx.Document(p)
        for r in doc.tables[0].rows:
            print(" | ".join(c.text.strip().replace("\n", " ") for c in r.cells))
    else:
        print("Not found")
