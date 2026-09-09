import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK/35. INFORME SG REPROD. ANIMAL 2025.docx")
for idx, p in enumerate(doc.paragraphs):
    if "seguimiento a graduados cohorte" in p.text.lower() and "se tiene un total de" in p.text.lower():
        print(f"Original Paragraph {idx}: {p.text}")
