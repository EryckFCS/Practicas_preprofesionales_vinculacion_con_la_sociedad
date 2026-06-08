import docx
import os

doc_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/26.  INFORME SG ECONOMIA 2023.docx"
doc = docx.Document(doc_path)
for idx, p in enumerate(doc.paragraphs):
    if "También se observan" in p.text:
        print(f"[{idx}]: Style={repr(p.style.name)} Text={repr(p.text)}")
        for r_idx, run in enumerate(p.runs):
            print(f"  Run [{r_idx}]: Text={repr(run.text)}, Bold={run.bold}, Size={run.font.size}")
