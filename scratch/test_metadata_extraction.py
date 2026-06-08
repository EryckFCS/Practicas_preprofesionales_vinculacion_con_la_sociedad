import docx
import os
from src.config import CAREER_FILES

def identify_table_type(table):
    if not table.rows:
        return "unknown"
    all_text = " ".join(c.text.strip().lower() for r in table.rows for c in r.cells)
    if "nombre de la carrera" in all_text and "periodo del estudio" in all_text:
        return "metadata"
    elif "nombre del programa de posgrado" in all_text and "periodo del estudio" in all_text:
        return "metadata"
    return "unknown"

for rid, career in sorted(CAREER_FILES.items()):
    source_docx_path = os.path.join(
        "/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK",
        career["word"]
    )
    if os.path.exists(source_docx_path):
        doc = docx.Document(source_docx_path)
        orig_poblacion = None
        for table in doc.tables:
            if identify_table_type(table) == "metadata":
                for row in table.rows:
                    cells = [c.text.strip().lower() for c in row.cells]
                    if any("población" in cell or "poblacion" in cell for cell in cells):
                        try:
                            orig_poblacion = int(row.cells[-1].text.strip())
                        except ValueError:
                            pass
                break
        print(f"Report {rid}: {career['word']} -> Extracted Population: {orig_poblacion}")
