import os
import docx
from src.config import CAREER_FILES, INFORMES_DIR

for rid in sorted(CAREER_FILES.keys()):
    if rid >= 35:
        path = os.path.join(INFORMES_DIR, CAREER_FILES[rid]["word"])
        if os.path.exists(path):
            doc = docx.Document(path)
            # Find correlation table
            for i, table in enumerate(doc.tables):
                all_text = " ".join(c.text.strip().lower() for r in table.rows for c in r.cells)
                if "correlación" in all_text or "correlacion" in all_text:
                    # check if row 0 has headers
                    headers = [c.text.strip() for c in table.rows[0].cells]
                    print(f"Report {rid}: Table {i} headers: {headers}")
                    break
