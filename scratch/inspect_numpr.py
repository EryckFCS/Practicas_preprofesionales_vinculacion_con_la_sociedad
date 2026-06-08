import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
for i in range(125, 129):
    p = doc.paragraphs[i]
    print(f"[{i}]: text={repr(p.text)}")
    pPr = p._p.get_or_add_pPr()
    numPr = pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr")
    if numPr is not None:
        print(f"  -> has numPr numbering details!")
        ilvl = numPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
        numId = numPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId")
        print(f"     ilvl={ilvl.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if ilvl is not None else None}")
        print(f"     numId={numId.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if numId is not None else None}")
