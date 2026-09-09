import docx
import os
from src.config import CAREER_FILES, INFORMES_DIR

for rid in range(35, 47):
    career = CAREER_FILES[rid]
    word_path = os.path.join(INFORMES_DIR, career["word"])
    if not os.path.exists(word_path):
        print(f"Report {rid}: {career['word']} not found")
        continue
    doc = docx.Document(word_path)
    poblacion = None
    muestra = None
    for table in doc.tables:
        # Check if it's the metadata table
        all_text = " ".join(c.text.strip().lower() for r in table.rows for c in r.cells)
        if "periodo del estudio" in all_text:
            for row in table.rows:
                cells = [c.text.strip().lower() for c in row.cells]
                if any("población" in cell or "poblacion" in cell for cell in cells):
                    poblacion = row.cells[-1].text.strip()
                if any("muestra" in cell or "encuestados" in cell for cell in cells) and not any("empleadores" in cell for cell in cells):
                    muestra = row.cells[-1].text.strip()
            break
    print(f"Report {rid}: {career['word']} -> Poblacion in Word: {poblacion}, Muestra in Word: {muestra}")
