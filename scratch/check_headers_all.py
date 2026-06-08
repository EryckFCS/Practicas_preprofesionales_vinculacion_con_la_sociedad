import os
import docx

for report_id in range(24, 35):
    career = {
        24: "24.  INFORME SG CONTAB. Y AUDIT 2023.docx",
        25: "25.  INFORME SG DERECHO 2023.docx",
        26: "26.  INFORME SG ECONOMIA 2023.docx",
        27: "27.  INFORME SG ENFERMERIA 2023.docx",
        28: "28.  INFORME SG LAB CLINICO 2023.docx",
        29: "29.  INFORME SG MEDICINA HUMANA 2023.docx",
        30: "30.  INFORME SG ODONTOLOGIA 2023.docx",
        31: "31.  INFORME SG PSICOL. CLINICA 2023.docx",
        32: "32.  INFORME SG ADM. EMPRESAS 2023.docx",
        33: "33.  INFORME SG CONTABILIDAD Y AUDIT. 2023.docx",
        34: "34.  INFORME SG DERECHO 2023.docx"
    }[report_id]
    doc_path = os.path.join("/home/erick-fcs/Descargas/Practicas_preprofesionales/INFORMES 24-46 ERICK", career)
    if os.path.exists(doc_path):
        doc = docx.Document(doc_path)
        found_pe = []
        found_ea = []
        for idx, p in enumerate(doc.paragraphs):
            t = p.text.strip()
            if "percepción sobre el proceso de enseñanza-aprendizaje" in t.lower():
                found_ea.append((idx, t))
            elif "percepción sobre el perfil de egreso" in t.lower():
                found_pe.append((idx, t))
        print(f"Report {report_id}: EA={found_ea}, PE={found_pe}")
    else:
        print(f"Report {report_id}: file not found")
