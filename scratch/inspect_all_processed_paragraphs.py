import os
import docx

for rid in range(35, 47):
    word_name = f"{rid}. INFORME SG "
    # let's find the actual filename in processed/posgrado/
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
        if "integrada por" in p.text.lower() or "predominancia" in p.text.lower():
            print(f"  Paragraph: {p.text}")
        if "perfil demográfico" in p.text.lower() and "conclusiones" in p.text.lower():
            print(f"  Conclusion: {p.text}")
        elif "perfil demográfico de la cohorte" in p.text.lower():
            print(f"  Conclusion: {p.text}")
