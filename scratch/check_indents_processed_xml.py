import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
for i in range(120, 137):
    p = doc.paragraphs[i]
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
    ind_str = ""
    if ind is not None:
        ind_str = f"ind XML attributes: {dict(ind.attrib)}"
    print(f"[{i}]: Style={repr(p.style.name)} Text={repr(p.text[:30])} left_indent={p.paragraph_format.left_indent} first_line_indent={p.paragraph_format.first_line_indent} {ind_str}")
