import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/posgrado/35. INFORME SG REPROD. ANIMAL 2025.docx")
for idx, p in enumerate(doc.paragraphs):
    if "17" in p.text:
        print(f"Paragraph {idx}: {repr(p.text)}")
