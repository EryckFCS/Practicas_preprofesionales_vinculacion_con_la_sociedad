import docx

doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK/25.  INFORME SG DERECHO 2023.docx")
for idx, p in enumerate(doc.paragraphs):
    txt = p.text.strip()
    if "predominancia" in txt.lower() or "integrada por" in txt.lower():
        print(f"P{idx}: {txt}")
