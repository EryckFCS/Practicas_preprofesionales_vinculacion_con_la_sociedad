import docx
import os
from src.config import CAREER_FILES, INFORMES_DIR

for rid in range(35, 47):
    word_name = CAREER_FILES[rid]["word"]
    path = os.path.join(INFORMES_DIR, word_name)
    if os.path.exists(path):
        doc = docx.Document(path)
        for idx, p in enumerate(doc.paragraphs):
            t = p.text
            if "seguimiento a graduados" in t.lower() or "un total de" in t.lower():
                print(f"Report {rid} [{idx}]: {t[:120]}...")
