import docx
import os
from src.config import CAREER_FILES

for rid, career in sorted(CAREER_FILES.items()):
    source_docx_path = os.path.join(
        "/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK",
        career["word"]
    )
    if os.path.exists(source_docx_path):
        try:
            doc = docx.Document(source_docx_path)
            # Find Table 0 / metadata table
            poblacion = "Not Found"
            muestra = "Not Found"
            for table in doc.tables:
                # check if it is metadata table
                all_text = " ".join(c.text.strip().lower() for r in table.rows for c in r.cells)
                if "población" in all_text or "poblacion" in all_text or "graduados" in all_text:
                    for row in table.rows:
                        cells = [c.text.strip() for c in row.cells]
                        cells_low = [c.lower() for c in cells]
                        if any("población" in c or "poblacion" in c for c in cells_low):
                            poblacion = cells[-1]
                        if any("encuestados" in c or "muestra" in c for c in cells_low) and not any("empleadores" in c for c in cells_low):
                            muestra = cells[-1]
                    break
            print(f"Report {rid}: {career['word']} -> Poblacion: {poblacion}, Muestra: {muestra}")
        except Exception as e:
            print(f"Report {rid}: Error {e}")
    else:
        print(f"Report {rid}: Document not found at {source_docx_path}")
