import os
import docx
from src.config import CAREER_FILES, INFORMES_DIR
from src.updater import identify_table_type

for rid in sorted(CAREER_FILES.keys()):
    if rid >= 35:
        path = os.path.join(INFORMES_DIR, CAREER_FILES[rid]["word"])
        if os.path.exists(path):
            doc = docx.Document(path)
            for i, table in enumerate(doc.tables):
                t_type = identify_table_type(table)
                if t_type == "correlation":
                    headers = [c.text.strip().replace('\n', ' ') for c in table.rows[0].cells]
                    print(f"Report {rid}: Table {i} matches correlation. Headers: {headers}")
