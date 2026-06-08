import pandas as pd
import os
import re

SPANISH_CORRECTIONS = {
    r"\bpractica\b": "práctica",
    r"\bgestión publica\b": "gestión pública",
    r"\bgestion publica\b": "gestión pública",
    r"\bcontratacion publica\b": "contratación pública",
    r"\bcontratación publica\b": "contratación pública",
    r"\bconocimiento tecnologicos\b": "conocimientos tecnológicos",
    r"\bconocimiento tecnológicos\b": "conocimientos tecnológicos",
    r"\bconocimientos tecnologicos\b": "conocimientos tecnológicos",
    r"\bconocimiento tecnologico\b": "conocimientos tecnológicos",
    r"\btramites\b": "trámites",
    r"\bredaccion\b": "redacción",
    r"\bhablidades\b": "habilidades",
    r"\binteligencia artif\b": "inteligencia artificial",
    r"\bentorno juridico\b": "entorno jurídico",
    r"\btecnicas de litigacion\b": "técnicas de litigación",
    r"\blitigacion oral\b": "litigación oral",
    r"\blitigación oral\b": "litigación oral",
    r"\boratoria juridica\b": "oratoria jurídica",
    r"\boratoria jurídica\b": "oratoria jurídica",
    r"\bcodigo\b": "código",
    r"\bde lo contencioso administrativo\b": "de lo contencioso-administrativo",
    r"\bpenal\b": "penal",
    r"\bnotarial\b": "notarial",
    r"\btributario\b": "tributario",
    r"\bsocietario\b": "societario",
    r"\bconstitucional\b": "constitucional",
    r"\bmercantil\b": "mercantil",
    r"\bcootad\b": "COOTAD",
    r"\bunl\b": "UNL",
    r"\biess\b": "IESS",
    r"\bsn\b": "",
    r"\bde genero\b": "de género",
    r"\bviolencia política de género\b": "violencia política de género",
}

def correct_spanish_spelling(text):
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text).strip()
    
    # If the text is mostly uppercase (e.g. "EN El TEMA DE VIOLENCIA...")
    if len(text) > 0:
        upper_ratio = sum(1 for c in text if c.isupper()) / len(text)
        if text.isupper() or upper_ratio > 0.4:
            text = text.lower()
            
    for pattern, replacement in SPANISH_CORRECTIONS.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    if text:
        text = text[0].upper() + text[1:]
    text = re.sub(r"\bderecho\b", "Derecho", text, flags=re.IGNORECASE)
    for acr in ["COOTAD", "UNL", "IESS", "RUC", "RIMPE"]:
        text = re.sub(rf"\b{acr}\b", acr, text, flags=re.IGNORECASE)
    return text.strip()

def is_valid_open_answer(text):
    if not text:
        return False
    t_low = text.lower().strip()
    if len(t_low) <= 2 and t_low not in ["si", "no"]:
        return False
    placeholders = {
        "nan", "sn", "s/n", "ninguno", "ninguna", "sin respuesta", "sin comentarios", 
        "no aplica", "n/a", "no", "ninguno.", "ninguno/a", "ninguna de las anteriores",
        "no registra", "s", "n", "gf", "ningun", "ninguno/as", "ninguna.", "no aplica.",
        "sin observaciones", "ninguna observación", "ninguna observacion", "no, ninguna",
        "ninguno, todo bien", "ninguno, todo excelente", "ninguno/a.", "ningun comentario"
    }
    if t_low in placeholders:
        return False
    if "más preguntas" in t_low or "¿cuál es la importancia" in t_low or "¿por qué es necesario" in t_low or "¿cuál es el papel" in t_low:
        return False
    return True

path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/25/pregunta_abierta_temas_de_capacitacion_sugerida_por_los_empleadores.parquet"
df = pd.read_parquet(path)
answers = []
for col in df.columns[2:]:
    for a in df[col].dropna().tolist():
        a_str = str(a).strip()
        if is_valid_open_answer(a_str):
            corrected = correct_spanish_spelling(a_str)
            if corrected and corrected not in answers:
                answers.append(corrected)

for idx, ans in enumerate(answers):
    print(f"{idx}: {ans}")
