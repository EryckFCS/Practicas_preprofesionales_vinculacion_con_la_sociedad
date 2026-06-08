import os
import docx
from src.config import INFORMES_DIR

path = os.path.join(INFORMES_DIR, "35. INFORME SG REPROD. ANIMAL 2025.docx")
doc = docx.Document(path)

body = doc.element.body
elements = []
for el in body:
    if el.tag.endswith('p'):
        elements.append(('p', docx.text.paragraph.Paragraph(el, doc)))
    elif el.tag.endswith('tbl'):
        elements.append(('tbl', docx.table.Table(el, doc)))

tbl_idx = None
for idx, (el_type, el_obj) in enumerate(elements):
    if el_type == 'tbl':
        all_text = " ".join(c.text.strip().lower() for r in el_obj.rows for c in r.cells)
        if "correlación" in all_text or "correlacion" in all_text:
            headers = [c.text.strip().lower() for c in el_obj.rows[0].cells]
            if "competencia" in headers:
                tbl_idx = idx
                tbl_obj = el_obj
                break

if tbl_idx is not None:
    print("=== Paragraphs before Table ===")
    for i in range(max(0, tbl_idx - 3), tbl_idx):
        if elements[i][0] == 'p':
            print(f"[{i}]: {elements[i][1].text}")

    print("\n=== Table 6 headers ===")
    print([c.text.strip().replace('\n', ' ') for c in tbl_obj.rows[0].cells])

    print("\n=== Paragraphs after Table ===")
    for i in range(tbl_idx + 1, min(len(elements), tbl_idx + 4)):
        if elements[i][0] == 'p':
            print(f"[{i}]: {elements[i][1].text}")
else:
    print("Correlation table not found.")
