import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/posgrado/35. INFORME SG REPROD. ANIMAL 2025.docx")
for idx, p in enumerate(doc.paragraphs):
    if "recolección" in p.text.lower() or "cohorte 2025 se tiene" in p.text.lower():
        print(f"Paragraph {idx}: {repr(p.text)}")
