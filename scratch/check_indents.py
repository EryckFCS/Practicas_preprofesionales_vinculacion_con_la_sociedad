import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
for i in [126, 127, 128]:
    p = doc.paragraphs[i]
    print(f"[{i}]:")
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    if ind is not None:
        print(f"  ind: left={ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left')}, firstLine={ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine')}, hanging={ind.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hanging')}")
    else:
        print("  no ind element")
    print(f"  left_indent={p.paragraph_format.left_indent}, first_line_indent={p.paragraph_format.first_line_indent}")
