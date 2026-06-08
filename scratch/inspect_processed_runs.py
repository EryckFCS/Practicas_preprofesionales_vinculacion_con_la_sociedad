import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
for i in [128, 129, 130, 132]:
    p = doc.paragraphs[i]
    print(f"[{i}]: Style={repr(p.style.name)} Text={repr(p.text[:30])}")
    for r_idx, r in enumerate(p.runs):
        print(f"  Run [{r_idx}]: Text={repr(r.text[:30])} Font={r.font.name} Size={r.font.size}")
