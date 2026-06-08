import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
print("=== Styles in original document ===")
for style in doc.styles:
    if style.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:
        print(f"Paragraph style: {repr(style.name)}")
