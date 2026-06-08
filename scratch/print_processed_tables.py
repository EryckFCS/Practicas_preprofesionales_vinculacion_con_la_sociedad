import docx
import os

for rid in range(35, 47):
    from src.config import CAREER_FILES
    word_name = CAREER_FILES[rid]["word"]
    path = os.path.join("processed", word_name)
    if os.path.exists(path):
        doc = docx.Document(path)
        # find Table 0 or the table with 'poblacion' or 'muestra'
        for i, table in enumerate(doc.tables):
            # check if it is metadata
            all_text = " ".join(c.text.strip().lower() for r in table.rows for c in r.cells)
            if "población" in all_text or "muestra" in all_text:
                print(f"=== Report {rid} Table {i} ===")
                for row in table.rows:
                    print([c.text.strip() for c in row.cells])
                break
