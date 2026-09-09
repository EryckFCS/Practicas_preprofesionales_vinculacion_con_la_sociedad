import docx
doc = docx.Document("/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/35. INFORME SG REPROD. ANIMAL 2025.docx")
print("Paragraph 112:", doc.paragraphs[112].text)
