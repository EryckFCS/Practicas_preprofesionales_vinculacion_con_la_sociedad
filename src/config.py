import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Folders configuration
INFORMES_DIR = os.path.join(BASE_DIR, "INFORMES 24-46 ERICK")
RESULTADOS_DIR = os.path.join(BASE_DIR, "RESULTADOS-24-46 ERICK")
PREGUNTAS_DIR = os.path.join(BASE_DIR, "PREGUNTAS ABIERTAS")
PARQUET_CACHE_DIR = os.path.join(BASE_DIR, "data", "parquet")
PROCESSED_DIR = os.path.join(BASE_DIR, "processed")

# Ensure directories exist
os.makedirs(PROCESSED_DIR, exist_ok=True)

# Explicit file mapping for safety and robustness
CAREER_FILES = {
    24: {
        "excel": "24. IND-Contabilidad y Audit 2023.xlsx",
        "word": "24.  INFORME SG CONTAB. Y AUDIT 2023.docx",
        "open_sheet": "BD CONTABILIDAD",
        "open_career_name": "Contabilidad y Auditoría - presencial"
    },
    25: {
        "excel": "25. IND-Derecho 2023.xlsx",
        "word": "25.  INFORME SG DERECHO 2023.docx",
        "open_sheet": "BD DERECHO",
        "open_career_name": "Derecho - presencial"
    },
    26: {
        "excel": "26. IND-Economia 2023.xlsx",
        "word": "26.  INFORME SG ECONOMIA 2023.docx",
        "open_sheet": "BD ECONOMIA",
        "open_career_name": "Economía"
    },
    27: {
        "excel": "27. IND-Enfermeria 2023.xlsx",
        "word": "27.  INFORME SG ENFERMERIA 2023.docx",
        "open_sheet": "BD ENFERMERIA",
        "open_career_name": "ENFERMERIA (régimen 2019)"
    },
    28: {
        "excel": "28. IND-Laboratorio Clinico 2023.xlsx",
        "word": "28.  INFORME SG LAB CLINICO 2023.docx",
        "open_sheet": "BD LAB CLINICO",
        "open_career_name": "LABORATORIO CLINICO (régimen 2019)"
    },
    29: {
        "excel": "29. IND-Medicina Humana 2023.xlsx",
        "word": "29.  INFORME SG MEDICINA HUMANA 2023.docx",
        "open_sheet": "BD MEDICINA",
        "open_career_name": "Medicina Humana"
    },
    30: {
        "excel": "30. IND-Odontologia 2023.xlsx",
        "word": "30.  INFORME SG ODONTOLOGIA 2023.docx",
        "open_sheet": "BD ODONTOLOGIA",
        "open_career_name": "ODONTOLOGIA (régimen 2019)"
    },
    31: {
        "excel": "31. IND-Psicologia Clinica 2023.xlsx",
        "word": "31.  INFORME SG PSICOL. CLINICA 2023.docx",
        "open_sheet": "BD PSICOLOGIA",
        "open_career_name": "PSICOLOGIA CLINICA (régimen 2019)"
    },
    32: {
        "excel": "32. IND-Administracion Empresas  2023.xlsx",
        "word": "32.  INFORME SG ADM. EMPRESAS 2023.docx",
        "open_sheet": "BD ADM DE EMPRESAS UED",
        "open_career_name": "Administración de Empresas - distancia"
    },
    33: {
        "excel": "33. IND-Contabilidad y Audit  2023.xlsx",
        "word": "33.  INFORME SG CONTABILIDAD Y AUDIT. 2023.docx",
        "open_sheet": "BD CONTABILIDAD UED",
        "open_career_name": "Contabilidad y Auditoría - distancia"
    },
    34: {
        "excel": "34. IND-Derecho UED  2023.xlsx",
        "word": "34.  INFORME SG DERECHO 2023.docx",
        "open_sheet": "BD DERECHO UED",
        "open_career_name": "Derecho - distancia"
    },
    35: {
        "excel": "35 IND-REPRODUCC ANIMAL 2025.xlsx",
        "word": "35. INFORME SG REPROD. ANIMAL 2025.docx",
        "open_sheet": "REPRODUCCION ANIMAL",
        "open_career_name": "MAESTRÍA ACADÉMICA CON TRAYECTORIA PROFESIONAL: REPRODUCCIÓN ANIMAL - presencial"
    },
    36: {
        "excel": "36 IND-SANIDAD ANIMAL 2025.xlsx",
        "word": "36. INFORME SG SANIDAD ANIMAL 2025.docx",
        "open_sheet": "SANIDAD ANIMAL",
        "open_career_name": "MAESTRÍA ACADÉMICA CON TRAYECTORIA DE INVESTIGACIÓN: SANIDAD ANIMAL - presencial"
    },
    37: {
        "excel": "37 IND-MEDICINA VETERINARIA 2025.xlsx",
        "word": "37. INFORME SG MEDICINA VETERINARIA 2025.docx",
        "open_sheet": "MEDICINA VETERINARIA",
        "open_career_name": "MAESTRÍA  ACADÉMICA  EN MEDICINA VETERINARIA - presencial"
    },
    38: {
        "excel": "38 IND-PRODUCCION ANIMAL 2025.xlsx",
        "word": "38. INFORME SG PRODUC. ANIMAL 2025.docx",
        "open_sheet": "PRODUCCION ANIMAL",
        "open_career_name": "MAESTRÍA EN PRODUCCIÓN ANIMAL - hibrida"
    },
    39: {
        "excel": "39 IND-PEDAG. INGLES COMO LE 2025.xlsx",
        "word": "39. INFORME SG PED. INGLES COMO LE 2025.docx",
        "open_sheet": "PEDAG. INGLES COMO LE",
        "open_career_name": "MAESTRÍA  ACADÉMICA  EN PEDAGOGÍA DEL INGLÉS COMO LENGUA EXTRANJERA - en linea"
    },
    40: {
        "excel": "40 IND-EDUCACION ESPECIAL 2025.xlsx",
        "word": "40. INFORME SG EDUCACION ESPECIAL 2025.docx",
        "open_sheet": "EDUCACION ESPECIAL",
        "open_career_name": "MAESTRÍA EN EDUCACIÓN ESPECIAL - en linea"
    },
    41: {
        "excel": "41 IND-EDUCACION BASICA 2025.xlsx",
        "word": "41. INFORME SG EDUCACION BASICA 2025.docx",
        "open_sheet": "EDUCACION BASICA",
        "open_career_name": "MAESTRÍA  ACADÉMICA CON TRAYECTORIA PROFESIONAL: EDUCACIÓN BÁSICA - presencial"
    },
    42: {
        "excel": "42 IND-EDUCACION INICIAL 2025.xlsx",
        "word": "42. INFORME SG EDUCACION INICIAL 2025.docx",
        "open_sheet": "EDUCACION INICIAL",
        "open_career_name": "MAESTRÍA ACADÉMICA CON TRAYECTORIA PROFESIONAL: EDUCACIÓN INICIAL - en linea"
    },
    43: {
        "excel": "43 IND-ENSEÑANZA LL 2025.xlsx",
        "word": "43. INFORME SG ENSEÑANZA DE LL 2025.docx",
        "open_sheet": "ENSEÑANZA LL",
        "open_career_name": "MAESTRÍA EN EDUCACIÓN (ENSEÑANZA DE LA  LENGUA Y LITERATURA) - hibrida"
    },
    44: {
        "excel": "44 IND-DOC INVES EN ES 2025.xlsx",
        "word": "44. INFORME SG DOCENCIA E INVE EN ES 2025.docx",
        "open_sheet": "DOC INVES EN ES",
        "open_career_name": "MAESTRÍA EN EDUCACIÓN (DOCENCIA E INVESTIGACIÓN EN EDUCACIÓN SUPERIOR) - semipresencial"
    },
    45: {
        "excel": "45 IND-GESTION EDUCATIVA 2025.xlsx",
        "word": "45. INFORME SG GESTION EDUCAT. 2025.docx",
        "open_sheet": "GESTION EDUCATIVA",
        "open_career_name": "MAESTRÍA ACADÉMICA EN GESTIÓN EDUCATIVA - en linea"
    },
    46: {
        "excel": "46 IND-EDUCAC DOC SUPE 2025.xlsx",
        "word": "46. INFORME SG DOCENCIA SUPERIOR 2025.docx",
        "open_sheet": "EDUCAC DOC SUP",
        "open_career_name": "MAESTRÍA ACADÉMICA EN EDUCACION (DOCENCIA SUPERIOR) - en linea"
    }
}

def is_postgraduate(report_id):
    """Returns True if the report corresponds to a Postgraduate/Masters program."""
    return report_id >= 35

def get_report_category(report_id):
    """Returns the execution category used to separate outputs and batch runs."""
    return "posgrado" if is_postgraduate(report_id) else "pregrado"
