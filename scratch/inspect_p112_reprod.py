import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/posgrado/35. INFORME SG REPROD. ANIMAL 2025.docx")
for idx in range(110, 116):
    print(f"Paragraph {idx}: {repr(doc.paragraphs[idx].text)}")
