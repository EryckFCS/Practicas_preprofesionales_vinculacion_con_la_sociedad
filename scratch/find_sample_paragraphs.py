import docx
import os
from src.config import CAREER_FILES, INFORMES_DIR

for rid in range(35, 47):
    word_name = CAREER_FILES[rid]["word"]
    path = os.path.join(INFORMES_DIR, word_name)
    if os.path.exists(path):
        doc = docx.Document(path)
        found = False
        for idx, p in enumerate(doc.paragraphs):
            t = p.text
            if "la muestra para la" in t.lower() or "la muestra para el" in t.lower():
                print(f"Report {rid} [{idx}]: {t[:160]}...")
                found = True
        if not found:
            print(f"Report {rid}: Not found")
    else:
        print(f"Report {rid}: Path does not exist {path}")
