import os
import docx

for rid in range(35, 47):
    dir_path = "processed/posgrado"
    found_file = None
    for file in os.listdir(dir_path):
        if file.startswith(str(rid)) and file.endswith(".docx"):
            found_file = os.path.join(dir_path, file)
            break
    if not found_file:
        print(f"Report {rid} processed docx not found")
        continue
    doc = docx.Document(found_file)
    print(f"\n=== Report {rid}: {os.path.basename(found_file)} ===")
    for p in doc.paragraphs:
        if "se tiene un total de" in p.text.lower():
            print(f"  Paragraph: {p.text}")
        if "la muestra para" in p.text.lower() or "la muestra de" in p.text.lower():
            print(f"  Sample paragraph: {p.text}")
