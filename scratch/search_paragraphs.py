import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/35. INFORME SG REPROD. ANIMAL 2025.docx")
for idx, p in enumerate(doc.paragraphs):
    if "recolección de datos" in p.text.lower() or "cohorte" in p.text.lower():
        print(f"Paragraph {idx}: {p.text}")
