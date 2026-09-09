import docx
import os

docx_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/posgrado/38. INFORME SG PRODUC. ANIMAL 2025.docx"
if not os.path.exists(docx_path):
    print("File not found")
else:
    doc = docx.Document(docx_path)
    for idx, p in enumerate(doc.paragraphs):
        text = p.text
        if "muestra" in text.lower() or "población" in text.lower() or "poblacion" in text.lower():
            print(f"[{idx}]: {text}")
