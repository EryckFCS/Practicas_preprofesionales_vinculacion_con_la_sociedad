import os
import logging
import pandas as pd
import unicodedata
from src.config import PARQUET_CACHE_DIR, CAREER_FILES, is_postgraduate

# Module logger
logger = logging.getLogger(__name__)
if not logger.handlers:
    # Basic configuration if not already configured by caller
    logging.basicConfig(level=logging.INFO)

def clean_label(text):
    """Normalizes text by removing accents, converting to lowercase, and stripping whitespace."""
    if not isinstance(text, str):
        return ""
    nfkd_form = unicodedata.normalize('NFKD', text)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower().strip()


def safe_float(val, default=0.0):
    try:
        if not val or str(val).strip() == "":
            return default
        # Replace comma with dot if present
        v_str = str(val).replace(",", ".").replace("%", "").strip()
        return float(v_str)
    except Exception:
        return default

def safe_int(val, default=0):
    try:
        if not val or str(val).strip() == "":
            return default
        v_str = str(val).split(".")[0].strip()
        return int(v_str)
    except Exception:
        return default

def extract_metadata(report_id, parquet_dir):
    """Extracts general metadata (population N, sample size n) from demografico and indicadores sheets."""
    # Preferred source: demografico.parquet (single source of truth)
    muestra = 0
    poblacion = 0
    dem_file = os.path.join(parquet_dir, "demografico.parquet")
    ind_file = os.path.join(parquet_dir, "indicadores.parquet")

    dem_muestra = 0
    if os.path.exists(dem_file):
        df_dem = pd.read_parquet(dem_file)
        if len(df_dem) > 0:
            # Row 0 contains 'n_de_graduados' (total sample size under the cohorte heading)
            dem_val = df_dem.iloc[0].get('n_de_graduados', 0)
            dem_muestra = safe_int(dem_val)

    ind_poblacion = 0
    if os.path.exists(ind_file):
        df_ind = pd.read_parquet(ind_file)
        for idx, row in df_ind.iterrows():
            v1 = str(row.get('variables_encuesta_a_graduados', '')).strip().lower()
            if "total" in v1:
                # Check potential columns containing the population value
                for col in ['unnamed_2', 'resultados', 'n_de_graduados']:
                    if col in df_ind.columns:
                        val = safe_int(row.get(col, 0))
                        if val > 0:
                            ind_poblacion = val
                            break
                if ind_poblacion > 0:
                    break

    # Apply policy: demografico.parquet is the single truth for final report sample size
    if dem_muestra > 0:
        muestra = dem_muestra
        if ind_poblacion > 0 and not is_postgraduate(report_id):
            poblacion = ind_poblacion
        else:
            poblacion = dem_muestra
        if ind_poblacion > 0 and not is_postgraduate(report_id) and ind_poblacion != dem_muestra:
            logger.info(
                "Report %s: demografico.muestra=%s differs from indicadores.poblacion=%s",
                report_id,
                dem_muestra,
                ind_poblacion,
            )
        else:
            logger.info("Report %s: using demografico.muestra=%s for report metadata", report_id, dem_muestra)
    else:
        # Fall back to previously used strategy when demografico is missing
        muestra = 0
        poblacion = 0
        # try demografico for muestra if possible (redundant but safe)
        if os.path.exists(dem_file):
            try:
                df_dem = pd.read_parquet(dem_file)
                if len(df_dem) > 0:
                    val = df_dem.iloc[0].get('n_de_graduados', 0)
                    muestra = safe_int(val)
            except Exception:
                muestra = 0

        # try indicadores for poblacion
        if ind_poblacion > 0:
            poblacion = ind_poblacion
            logger.info("Report %s: poblacion taken from indicadores.parquet = %s", report_id, poblacion)

        # Fallbacks if still 0
        if muestra == 0:
            muestra = 14  # Default sensible fallback
            logger.warning("Report %s: 'muestra' not found in parquet cache, applying fallback muestra=14", report_id)
        if poblacion == 0:
            poblacion = muestra * 2  # Default sensible fallback
            logger.warning("Report %s: 'poblacion' not found in parquet cache, applying fallback poblacion=%s", report_id, poblacion)
        
    return {
        "poblacion": poblacion,
        "muestra": muestra,
        "carrera": CAREER_FILES[report_id]["open_career_name"],
        "cohorte": "2025" if report_id >= 35 else "2023"
    }

def extract_demographics(parquet_dir):
    """Extracts gender, civil status, and age distributions."""
    df = pd.read_parquet(os.path.join(parquet_dir, "demografico.parquet"))
    
    # We map categories in unnamed_1, count in n_de_graduados, percentage in unnamed_3
    data = {
        # Gender
        "hombre_freq": 0, "hombre_pct": 0.0,
        "mujer_freq": 0, "mujer_pct": 0.0,
        # Civil Status
        "soltero_freq": 0, "soltero_pct": 0.0,
        "casado_freq": 0, "soltero_casado_pct": 0.0,  # Note: key names matching our table replacement targets
        "casado_pct": 0.0,
        "divorciado_freq": 0, "divorciado_pct": 0.0,
        "viudo_freq": 0, "viudo_pct": 0.0,
        "union_libre_freq": 0, "union_libre_pct": 0.0,
        # Age Groups
        "edad_21_25_freq": 0, "edad_21_25_pct": 0.0,
        "edad_26_30_freq": 0, "edad_26_30_pct": 0.0,
        "edad_31_35_freq": 0, "edad_31_35_pct": 0.0,
        "edad_36_40_freq": 0, "edad_36_40_pct": 0.0,
        "edad_41_45_freq": 0, "edad_41_45_pct": 0.0,
        "edad_mayor_45_freq": 0, "edad_mayor_45_pct": 0.0,
    }
    
    for idx, row in df.iterrows():
        cat = clean_label(row.get('unnamed_1', ''))
        freq = safe_int(row.get('n_de_graduados', 0))
        pct = safe_float(row.get('unnamed_3', 0.0))
        
        if "hombre" in cat:
            data["hombre_freq"] = freq
            data["hombre_pct"] = pct
        elif "mujer" in cat:
            data["mujer_freq"] = freq
            data["mujer_pct"] = pct
        elif "soltero" in cat:
            data["soltero_freq"] = freq
            data["soltero_pct"] = pct
        elif "casado" in cat:
            data["casado_freq"] = freq
            data["casado_pct"] = pct
        elif "divorciado" in cat:
            data["divorciado_freq"] = freq
            data["divorciado_pct"] = pct
        elif "viudo" in cat:
            data["viudo_freq"] = freq
            data["viudo_pct"] = pct
        elif "union libre" in cat:
            data["union_libre_freq"] = freq
            data["union_libre_pct"] = pct
        elif "21 y 25" in cat:
            data["edad_21_25_freq"] = freq
            data["edad_21_25_pct"] = pct
        elif "26 y 30" in cat:
            data["edad_26_30_freq"] = freq
            data["edad_26_30_pct"] = pct
        elif "31 y 35" in cat:
            data["edad_31_35_freq"] = freq
            data["edad_31_35_pct"] = pct
        elif "36 y 40" in cat:
            data["edad_36_40_freq"] = freq
            data["edad_36_40_pct"] = pct
        elif "41 y 45" in cat:
            data["edad_41_45_freq"] = freq
            data["edad_41_45_pct"] = pct
        elif "mayor" in cat or "45 anos" in cat:
            data["edad_mayor_45_freq"] = freq
            data["edad_mayor_45_pct"] = pct

    return data

def extract_training_needs(parquet_dir):
    """Extracts postgraduate training preferences, languages, and activity interests."""
    df = pd.read_parquet(os.path.join(parquet_dir, "necesidades_de_formacion.parquet"))
    
    data = {
        "interes_si_freq": 0, "interes_si_pct": 0.0,
        "interes_no_freq": 0, "interes_no_pct": 0.0,
        # Training types
        "capacitacion_seminarios_freq": 0, "capacitacion_seminarios_pct": 0.0,
        "capacitacion_certificacion_freq": 0, "capacitacion_certificacion_pct": 0.0,
        "capacitacion_mooc_freq": 0, "capacitacion_mooc_pct": 0.0,
        "capacitacion_elearning_freq": 0, "capacitacion_elearning_pct": 0.0,
        "capacitacion_diplomados_freq": 0, "capacitacion_diplomados_pct": 0.0,
        "capacitacion_especializaciones_freq": 0, "capacitacion_especializaciones_pct": 0.0,
        "capacitacion_maestria_freq": 0, "capacitacion_maestria_pct": 0.0,
        "capacitacion_doctorado_freq": 0, "capacitacion_doctorado_pct": 0.0,
    }
    
    # Needs matching on "variable", "categorias", "n_de_graduados", "porcentaje"
    # Note: Column names are cleaned, let's see which exist
    cols = df.columns.tolist()
    cat_col = 'categorias' if 'categorias' in cols else ('unnamed_1' if 'unnamed_1' in cols else cols[1])
    freq_col = 'n_de_graduados' if 'n_de_graduados' in cols else ('frecuencia' if 'frecuencia' in cols else cols[2])
    pct_col = 'porcentaje' if 'porcentaje' in cols else ('unnamed_3' if 'unnamed_3' in cols else cols[3])

    for idx, row in df.iterrows():
        cat = clean_label(row.get(cat_col, ''))
        freq = safe_int(row.get(freq_col, 0))
        pct = safe_float(row.get(pct_col, 0.0))
        
        # Interest
        if cat == "si" and idx < 4:
            data["interes_si_freq"] = freq
            data["interes_si_pct"] = pct
        elif cat == "no" and idx < 4:
            data["interes_no_freq"] = freq
            data["interes_no_pct"] = pct
        # Training types
        elif "simposios" in cat or "congresos" in cat:
            data["capacitacion_seminarios_freq"] = freq
            data["capacitacion_seminarios_pct"] = pct
        elif "certificacion" in cat:
            data["capacitacion_certificacion_freq"] = freq
            data["capacitacion_certificacion_pct"] = pct
        elif "mooc" in cat or "autoinstruccionales" in cat:
            data["capacitacion_mooc_freq"] = freq
            data["capacitacion_mooc_pct"] = pct
        elif "e-learning" in cat or "tutorias" in cat:
            data["capacitacion_elearning_freq"] = freq
            data["capacitacion_elearning_pct"] = pct
        elif "diplomado" in cat:
            data["capacitacion_diplomados_freq"] = freq
            data["capacitacion_diplomados_pct"] = pct
        elif "especializacion" in cat:
            data["capacitacion_especializaciones_freq"] = freq
            data["capacitacion_especializaciones_pct"] = pct
        elif "maestria" in cat:
            data["capacitacion_maestria_freq"] = freq
            data["capacitacion_maestria_pct"] = pct
        elif "doctorado" in cat:
            data["capacitacion_doctorado_freq"] = freq
            data["capacitacion_doctorado_pct"] = pct

    return data

def extract_employment_status(parquet_dir):
    """Extracts job situation values."""
    df = pd.read_parquet(os.path.join(parquet_dir, "empleabilidad.parquet"))
    
    data = {
        "empleo_publico_freq": 0, "empleo_publico_pct": 0.0,
        "empleo_privado_freq": 0, "empleo_privado_pct": 0.0,
        "empleo_libre_freq": 0, "empleo_libre_pct": 0.0,
        "empleo_emprendedor_freq": 0, "empleo_emprendedor_pct": 0.0,
        "empleo_desempleo_freq": 0, "empleo_desempleo_pct": 0.0,
    }
    
    cols = df.columns.tolist()
    cat_col = 'categorias' if 'categorias' in cols else ('unnamed_1' if 'unnamed_1' in cols else cols[1])
    freq_col = 'n_de_graduados' if 'n_de_graduados' in cols else ('frecuencia' if 'frecuencia' in cols else cols[2])
    pct_col = 'porcentaje' if 'porcentaje' in cols else ('unnamed_3' if 'unnamed_3' in cols else cols[3])

    for idx, row in df.iterrows():
        cat = clean_label(row.get(cat_col, ''))
        freq = safe_int(row.get(freq_col, 0))
        pct = safe_float(row.get(pct_col, 0.0))
        
        if "publico" in cat:
            data["empleo_publico_freq"] = freq
            data["empleo_publico_pct"] = pct
        elif "privado" in cat:
            data["empleo_privado_freq"] = freq
            data["empleo_privado_pct"] = pct
        elif "libre" in cat or "facturacion" in cat:
            data["empleo_libre_freq"] = freq
            data["empleo_libre_pct"] = pct
        elif "emprendedor" in cat or "ruc" in cat or "rimpe" in cat:
            data["empleo_emprendedor_freq"] = freq
            data["empleo_emprendedor_pct"] = pct
        elif "sin actividad" in cat or "desempleo" in cat or "desocupado" in cat:
            data["empleo_desempleo_freq"] = freq
            data["empleo_desempleo_pct"] = pct
            
    return data

def extract_study_continuity(parquet_dir):
    """Extracts subsequent studies status."""
    df = pd.read_parquet(os.path.join(parquet_dir, "continuidad_estudios.parquet"))
    
    data = {
        "cont_talleres_freq": 0, "cont_talleres_pct": 0.0,
        "cont_diplomado_freq": 0, "cont_diplomado_pct": 0.0,
        "cont_especialidad_freq": 0, "cont_especialidad_pct": 0.0,
        "cont_maestria_freq": 0, "cont_maestria_pct": 0.0,
        "cont_doctorado_freq": 0, "cont_doctorado_pct": 0.0,
        "cont_posdoctorado_freq": 0, "cont_posdoctorado_pct": 0.0,
        "cont_otra_carrera_freq": 0, "cont_otra_carrera_pct": 0.0,
        "cont_ninguno_freq": 0, "cont_ninguno_pct": 0.0,
        "apoyo_si_freq": 0, "apoyo_si_pct": 0.0,
        "apoyo_no_freq": 0, "apoyo_no_pct": 0.0,
        "apoyo_en_parte_freq": 0, "apoyo_en_parte_pct": 0.0,
    }
    
    cols = df.columns.tolist()
    cat_col = 'categorias' if 'categorias' in cols else ('unnamed_1' if 'unnamed_1' in cols else cols[1])
    freq_col = 'n_de_graduados' if 'n_de_graduados' in cols else ('frecuencia' if 'frecuencia' in cols else cols[2])
    pct_col = 'porcentaje' if 'porcentaje' in cols else ('unnamed_3' if 'unnamed_3' in cols else cols[3])

    for idx, row in df.iterrows():
        cat = clean_label(row.get(cat_col, ''))
        freq = safe_int(row.get(freq_col, 0))
        pct = safe_float(row.get(pct_col, 0.0))
        
        if "cursos" in cat or "talleres" in cat or "simposios" in cat:
            data["cont_talleres_freq"] = freq
            data["cont_talleres_pct"] = pct
        elif "diplomado" in cat:
            data["cont_diplomado_freq"] = freq
            data["cont_diplomado_pct"] = pct
        elif "especialidad" in cat:
            data["cont_especialidad_freq"] = freq
            data["cont_especialidad_pct"] = pct
        elif "maestria" in cat:
            data["cont_maestria_freq"] = freq
            data["cont_maestria_pct"] = pct
        elif "doctorado" in cat:
            data["cont_doctorado_freq"] = freq
            data["cont_doctorado_pct"] = pct
        elif "posdoc" in cat:
            data["cont_posdoctorado_freq"] = freq
            data["cont_posdoctorado_pct"] = pct
        elif "otra carrera" in cat:
            data["cont_otra_carrera_freq"] = freq
            data["cont_otra_carrera_pct"] = pct
        elif "no" in cat or "aun no" in cat:
            data["cont_ninguno_freq"] = freq
            data["cont_ninguno_pct"] = pct
            
    # Extract institutional support
    if "categorias_1" in cols:
        for idx, row in df.iterrows():
            cat_1 = str(row.get("categorias_1", "")).strip().lower()
            freq_1 = safe_int(row.get("n_de_graduados_1", 0))
            pct_1 = safe_float(row.get("porcentaje_1", 0.0))
            
            if cat_1 == "si":
                data["apoyo_si_freq"] = freq_1
                data["apoyo_si_pct"] = pct_1
            elif cat_1 == "no":
                data["apoyo_no_freq"] = freq_1
                data["apoyo_no_pct"] = pct_1
            elif "en parte" in cat_1:
                data["apoyo_en_parte_freq"] = freq_1
                data["apoyo_en_parte_pct"] = pct_1
                
    return data

def extract_demand_level(parquet_dir):
    """Extracts job demand level (Alta, Media, Baja, Muy Baja)."""
    df = pd.read_parquet(os.path.join(parquet_dir, "demanda.parquet"))
    
    data = {
        "demanda_alta_freq": 0, "demanda_alta_pct": 0.0,
        "demanda_media_freq": 0, "demanda_media_pct": 0.0,
        "demanda_baja_freq": 0, "demanda_baja_pct": 0.0,
        "demanda_muy_baja_freq": 0, "demanda_muy_baja_pct": 0.0,
    }
    
    cols = df.columns.tolist()
    cat_col = 'categorias' if 'categorias' in cols else ('unnamed_1' if 'unnamed_1' in cols else cols[1])
    freq_col = 'n_de_graduados' if 'n_de_graduados' in cols else ('frecuencia' if 'frecuencia' in cols else cols[2])
    pct_col = 'porcentaje' if 'porcentaje' in cols else ('unnamed_3' if 'unnamed_3' in cols else cols[3])

    for idx, row in df.iterrows():
        cat = clean_label(row.get(cat_col, ''))
        freq = safe_int(row.get(freq_col, 0))
        pct = safe_float(row.get(pct_col, 0.0))
        
        if cat == "alta":
            data["demanda_alta_freq"] = freq
            data["demanda_alta_pct"] = pct
        elif cat == "media":
            data["demanda_media_freq"] = freq
            data["demanda_media_pct"] = pct
        elif cat == "baja":
            data["demanda_baja_freq"] = freq
            data["demanda_baja_pct"] = pct
        elif cat == "muy baja":
            data["demanda_muy_baja_freq"] = freq
            data["demanda_muy_baja_pct"] = pct
            
    return data

def extract_insercion_laboral(parquet_dir):
    """Extracts search time, difficulties, and bolsa de empleo effectiveness."""
    path = os.path.join(parquet_dir, "insercion_laboral.parquet")
    if not os.path.exists(path):
        return {}
    df = pd.read_parquet(path)
    
    data = {
        "bolsa_si_freq": 0, "bolsa_si_pct": 0.0,
        "bolsa_no_freq": 0, "bolsa_no_pct": 0.0,
        "bolsa_no_mostrada_freq": 0, "bolsa_no_mostrada_pct": 0.0,
    }
    
    # We parse the section for "Efectividad de la bolsa de empleo UNL"
    start_idx = None
    for idx, raw_val in enumerate(df.iloc[:, 0].tolist()):
        val = str(raw_val).lower()
        if "efectividad de la bolsa de empleo" in val:
            start_idx = idx
            break
            
    if start_idx is not None:
        for i in range(start_idx + 1, len(df)):
            val_0 = str(df.iloc[i, 0]).strip()
            if val_0 != "nan" and val_0 != "":
                break
            cat = str(df.iloc[i, 1]).strip().lower()
            freq = safe_int(df.iloc[i, 2])
            pct = safe_float(df.iloc[i, 3])
            
            if cat == "si":
                data["bolsa_si_freq"] = freq
                data["bolsa_si_pct"] = pct
            elif cat == "no":
                data["bolsa_no_freq"] = freq
                data["bolsa_no_pct"] = pct
            elif "no mostrada" in cat:
                data["bolsa_no_mostrada_freq"] = freq
                data["bolsa_no_mostrada_pct"] = pct
                
    return data

def parse_section(df, section_name):
    start_idx = None
    for idx, val in enumerate(df.iloc[:, 0].astype(str).tolist()):
        if section_name in val:
            start_idx = idx
            break
    
    if start_idx is None:
        return {}

    op_idx = None
    for i in range(start_idx + 1, min(start_idx + 5, len(df))):
        if str(df.iloc[i, 0]).strip().lower() == 'opción':
            op_idx = i
            break
            
    if op_idx is None:
        return {}

    results = {}
    for i in range(op_idx + 1, len(df)):
        val_0 = str(df.iloc[i, 0]).strip()
        if val_0 == 'nan' or val_0 == '' or 'resumen de campo' in val_0.lower():
            break
        
        count = str(df.iloc[i, 1]).strip()
        pct = str(df.iloc[i, 2]).strip()
        results[val_0] = {
            'count': int(count.split('.')[0]) if count != 'nan' and count != '' else 0,
            'pct': float(pct) if pct != 'nan' and pct != '' else 0.0
        }

    # If all pct values are 0 but counts exist, compute pct from total (postgrad datae issue)
    total_count = sum(v['count'] for v in results.values())
    if total_count > 0 and all(v['pct'] == 0.0 for v in results.values()):
        for key in results:
            results[key]['pct'] = results[key]['count'] / total_count

    return results

def _extract_employer_total_from_d1(df):
    """Computes total employer respondents from D1(SQ01) section counts."""
    vals = df.iloc[:, 0].astype(str).tolist()
    for i, v in enumerate(vals):
        if 'Resumen de campo para D1(SQ01)' in v:
            op_idx = None
            for j in range(i + 1, min(i + 5, len(df))):
                if str(df.iloc[j, 0]).strip().lower() == 'opción':
                    op_idx = j
                    break
            if op_idx is not None:
                total = 0
                for k in range(op_idx + 1, min(op_idx + 10, len(df))):
                    v0 = str(df.iloc[k, 0]).strip()
                    if v0 == 'nan' or v0 == '' or 'resumen de campo' in v0.lower():
                        break
                    cnt = str(df.iloc[k, 1]).strip()
                    try:
                        total += int(cnt.split('.')[0])
                    except Exception:
                        pass
                if total > 0:
                    return total
    return 0


def _extract_text_section(df, section_name):
    """Extracts free-text answers from a datae section (e.g., C2, C3, E1)."""
    vals = df.iloc[:, 0].astype(str).tolist()
    answers = []
    for i, v in enumerate(vals):
        if section_name in v:
            # Skip the question row and 'Opción' header, collect text answers
            for j in range(i + 2, min(i + 80, len(df))):
                v0 = str(df.iloc[j, 0]).strip()
                if 'resumen de campo' in v0.lower():
                    break
                if v0 and v0.lower() not in ['nan', '', 'opción', 'sin respuesta', 'no mostrada']:
                    # Also try col_1 which may have the text content
                    v1 = str(df.iloc[j, 1]).strip() if len(df.columns) > 1 else ''
                    text = v0 if len(v0) > 5 else v1
                    if text and text.lower() not in ['nan', '', 'sin respuesta', 'no mostrada', '0', '1']:
                        cleaned = text.replace('_x000D_', '').strip()
                        if cleaned and cleaned not in answers:
                            answers.append(cleaned)
            break
    return answers


def extract_employer_data(parquet_dir):
    parquet_path = os.path.join(parquet_dir, "datae.parquet")
    if not os.path.exists(parquet_path):
        return {}
    df = pd.read_parquet(parquet_path)
    
    b1 = parse_section(df, "Resumen de campo para B1")
    b2 = parse_section(df, "Resumen de campo para B2")
    b3 = parse_section(df, "Resumen de campo para B3")
    
    d1 = {}
    for i in range(1, 10):
        sec_name = f"Resumen de campo para D1(SQ0{i})"
        d1[f"sq{i}"] = parse_section(df, sec_name)

    # Total employer respondents (from D1 SQ01 counts if B1 totals zero)
    total_employers = sum(v.get('count', 0) for v in b1.values())
    if total_employers == 0:
        total_employers = _extract_employer_total_from_d1(df)

    # Extract employer cargos/funciones from datae C sections (for postgrad where parquet has 0 rows)
    employer_cargos_datae = _extract_text_section(df, "Resumen de campo para C2")
    employer_funciones_datae = _extract_text_section(df, "Resumen de campo para C3")

    # Employer training suggestions from datae E1 section
    employer_training_datae = _extract_text_section(df, "Resumen de campo para E1")

    return {
        "b1": b1,
        "b2": b2,
        "b3": b3,
        "d1": d1,
        "total_employers": total_employers,
        "employer_cargos_datae": employer_cargos_datae,
        "employer_funciones_datae": employer_funciones_datae,
        "employer_training_datae": employer_training_datae,
    }

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

def extract_open_questions(parquet_dir):
    """Extracts all non-empty open question text answers with spelling corrections."""
    open_data = {}
    if not os.path.exists(parquet_dir):
        return open_data

    for file in os.listdir(parquet_dir):
        if file.startswith("pregunta_abierta_") and file.endswith(".parquet"):
            path = os.path.join(parquet_dir, file)
            df = pd.read_parquet(path)
            
            if "cargos_y_funciones_de_los_graduados_empleadores" in file:
                cargos = []
                funciones = []
                if len(df.columns) >= 4:
                    col_cargos = df.columns[2]
                    col_funciones = df.columns[3]
                    
                    for a in df[col_cargos].dropna().tolist():
                        a_str = str(a).replace("_x000D_", "").strip()
                        if is_valid_open_answer(a_str):
                            corrected = correct_spanish_spelling(a_str)
                            if corrected and corrected not in cargos:
                                cargos.append(corrected)
                    
                    for a in df[col_funciones].dropna().tolist():
                        a_str = str(a).replace("_x000D_", "").strip()
                        if is_valid_open_answer(a_str):
                            corrected = correct_spanish_spelling(a_str)
                            if corrected and corrected not in funciones:
                                funciones.append(corrected)
                
                open_data["employer_cargos"] = cargos
                open_data["employer_funciones"] = funciones
                continue
            
            text_cols = [c for c in df.columns if c not in ["w", "carrera_o_programa", "carrera_o_programa_de_posgrado", "facultad"]]
            answers = []
            for col in text_cols:
                raw_answers = df[col].dropna().tolist()
                for a in raw_answers:
                    a_str = str(a).replace("_x000D_", "").strip()
                    if is_valid_open_answer(a_str):
                        corrected = correct_spanish_spelling(a_str)
                        if corrected and corrected not in answers:
                            answers.append(corrected)
            
            key_name = file.replace("pregunta_abierta_", "").replace(".parquet", "")
            open_data[key_name] = answers
                
    return open_data

def extract_correlation_data(parquet_dir):
    parquet_path = os.path.join(parquet_dir, "correlacion.parquet")
    if not os.path.exists(parquet_path):
        return {}
    df = pd.read_parquet(parquet_path)
    
    split_idx = None
    for idx, row in df.iterrows():
        row_vals = [str(x).upper() for x in row.values]
        if any("CORRELACION EMPLEADORES" in val or "CORRELACIÓN EMPLEADORES" in val for val in row_vals):
            split_idx = idx
            break
            
    if split_idx is None:
        split_idx = 12
        
    grad_data = {}
    emp_data = {}
    
    for idx, row in df.iterrows():
        label = str(row.iloc[6]).strip()
        val = row.iloc[7]
        lvl = str(row.iloc[8]).strip()
        
        if label == 'nan' or label == '' or 'detalle' in label.lower() or 'correlacion' in label.lower():
            continue
            
        try:
            val_float = float(str(val).replace(',', '.'))
        except:
            val_float = 0.0
            
        if idx < split_idx:
            grad_data[label] = {
                'value': val_float,
                'level': lvl
            }
        else:
            emp_data[label] = {
                'value': val_float,
                'level': lvl
            }
            
    return {
        'graduados': grad_data,
        'empleadores': emp_data
    }

def extract_satisfaccion(parquet_dir):
    """Extracts satisfaction with teaching-learning process and profile of egress from satisfaccion.parquet."""
    path = os.path.join(parquet_dir, "satisfaccion.parquet")
    if not os.path.exists(path):
        return {}

    df = pd.read_parquet(path)
    data = {
        # Perfil de egreso frequencies/percents
        "pe_resultados_buena_freq": 0, "pe_resultados_buena_pct": 0.0,
        "pe_resultados_regular_freq": 0, "pe_resultados_regular_pct": 0.0,
        "pe_resultados_mala_freq": 0, "pe_resultados_mala_pct": 0.0,
        "pe_resultados_no_mostrada_freq": 0, "pe_resultados_no_mostrada_pct": 0.0,
        "pe_cumplimiento_buena_freq": 0, "pe_cumplimiento_buena_pct": 0.0,
        "pe_cumplimiento_regular_freq": 0, "pe_cumplimiento_regular_pct": 0.0,
        "pe_cumplimiento_mala_freq": 0, "pe_cumplimiento_mala_pct": 0.0,
        "pe_cumplimiento_no_mostrada_freq": 0, "pe_cumplimiento_no_mostrada_pct": 0.0,
    }

    # Locate Pertinencia de los resultados de aprendizaje
    res_idx = None
    cump_idx = None
    for idx, row in df.iterrows():
        val = str(row.iloc[0]).strip().lower()
        if "resultados de aprendizaje" in val or "resultados de aprnedizaje" in val:
            res_idx = idx
        elif "cumplimiento del perfil" in val:
            cump_idx = idx

    is_pg = False
    try:
        is_pg = is_postgraduate(int(os.path.basename(parquet_dir)))
    except:
        pass

    # For postgraduate reports, the satisfaction data resides in datag.parquet instead of satisfaccion.parquet
    if is_pg or (res_idx is None and cump_idx is None):
        datag_path = os.path.join(parquet_dir, "datag.parquet")
        if os.path.exists(datag_path):
            try:
                df_g = pd.read_parquet(datag_path)
                col0 = df_g.columns[0]
                for idx_g, val_g in enumerate(df_g[col0].astype(str).tolist()):
                    val_lower = val_g.lower()
                    if "resultados de aprendizaje" in val_lower or "resultados de aprnedizaje" in val_lower:
                        # parse option rows
                        for i in range(idx_g + 1, min(idx_g + 7, len(df_g))):
                            opt = str(df_g.iloc[i, 0]).lower()
                            if "opción" in opt:
                                continue
                            cnt = safe_int(df_g.iloc[i, 1])
                            pct = safe_float(df_g.iloc[i, 2])
                            if "buena" in opt:
                                data["pe_resultados_buena_freq"] = cnt
                                data["pe_resultados_buena_pct"] = pct
                            elif "regular" in opt:
                                data["pe_resultados_regular_freq"] = cnt
                                data["pe_resultados_regular_pct"] = pct
                            elif "mala" in opt:
                                data["pe_resultados_mala_freq"] = cnt
                                data["pe_resultados_mala_pct"] = pct
                            elif "no mostrada" in opt or "sin respuesta" in opt:
                                data["pe_resultados_no_mostrada_freq"] += cnt
                                data["pe_resultados_no_mostrada_pct"] += pct
                    elif "cumplimiento del perfil" in val_lower:
                        # parse option rows
                        for i in range(idx_g + 1, min(idx_g + 7, len(df_g))):
                            opt = str(df_g.iloc[i, 0]).lower()
                            if "opción" in opt:
                                continue
                            cnt = safe_int(df_g.iloc[i, 1])
                            pct = safe_float(df_g.iloc[i, 2])
                            if "buena" in opt:
                                data["pe_cumplimiento_buena_freq"] = cnt
                                data["pe_cumplimiento_buena_pct"] = pct
                            elif "regular" in opt:
                                data["pe_cumplimiento_regular_freq"] = cnt
                                data["pe_cumplimiento_regular_pct"] = pct
                            elif "mala" in opt:
                                data["pe_cumplimiento_mala_freq"] = cnt
                                data["pe_cumplimiento_mala_pct"] = pct
                            elif "no mostrada" in opt or "sin respuesta" in opt:
                                data["pe_cumplimiento_no_mostrada_freq"] += cnt
                                data["pe_cumplimiento_no_mostrada_pct"] += pct
                return data
            except Exception as e:
                logger.warning(f"Error parsing datag fallback for satisfaction: {e}")

    # If not found by name, fallback to expected indexes (20 for resultados, 24 for cumplimiento)
    if res_idx is None:
        res_idx = 20
    if cump_idx is None:
        cump_idx = 24

    # Result categories: row res_idx is Buena, res_idx+1 is Regular, res_idx+2 is Mala, res_idx+3 is No mostrada
    try:
        data["pe_resultados_buena_freq"] = safe_int(df.iloc[res_idx, 2])
        data["pe_resultados_buena_pct"] = safe_float(df.iloc[res_idx, 3])
        data["pe_resultados_regular_freq"] = safe_int(df.iloc[res_idx+1, 2])
        data["pe_resultados_regular_pct"] = safe_float(df.iloc[res_idx+1, 3])
        data["pe_resultados_mala_freq"] = safe_int(df.iloc[res_idx+2, 2])
        data["pe_resultados_mala_pct"] = safe_float(df.iloc[res_idx+2, 3])
        data["pe_resultados_no_mostrada_freq"] = safe_int(df.iloc[res_idx+3, 2])
        data["pe_resultados_no_mostrada_pct"] = safe_float(df.iloc[res_idx+3, 3])
    except Exception as e:
        logger.warning(f"Error parsing resultados de aprendizaje: {e}")

    try:
        data["pe_cumplimiento_buena_freq"] = safe_int(df.iloc[cump_idx, 2])
        data["pe_cumplimiento_buena_pct"] = safe_float(df.iloc[cump_idx, 3])
        data["pe_cumplimiento_regular_freq"] = safe_int(df.iloc[cump_idx+1, 2])
        data["pe_cumplimiento_regular_pct"] = safe_float(df.iloc[cump_idx+1, 3])
        data["pe_cumplimiento_mala_freq"] = safe_int(df.iloc[cump_idx+2, 2])
        data["pe_cumplimiento_mala_pct"] = safe_float(df.iloc[cump_idx+2, 3])
        data["pe_cumplimiento_no_mostrada_freq"] = safe_int(df.iloc[cump_idx+3, 2])
        data["pe_cumplimiento_no_mostrada_pct"] = safe_float(df.iloc[cump_idx+3, 3])
    except Exception as e:
        logger.warning(f"Error parsing cumplimiento del perfil: {e}")

    return data

def extract_all_data(report_id):
    """Main extraction driver for a given report_id."""
    parquet_dir = os.path.join(PARQUET_CACHE_DIR, str(report_id))
    if not os.path.exists(parquet_dir):
        raise FileNotFoundError(f"Parquet cache directory not found for report {report_id}. Please run ingest first.")

    dataset = {}
    dataset["metadata"] = extract_metadata(report_id, parquet_dir)
    dataset["demographics"] = extract_demographics(parquet_dir)
    dataset["training_needs"] = extract_training_needs(parquet_dir)
    dataset["employment"] = extract_employment_status(parquet_dir)
    dataset["continuity"] = extract_study_continuity(parquet_dir)
    dataset["demand"] = extract_demand_level(parquet_dir)
    dataset["open_questions"] = extract_open_questions(parquet_dir)
    dataset["employer"] = extract_employer_data(parquet_dir)
    dataset["correlation"] = extract_correlation_data(parquet_dir)
    dataset["insercion"] = extract_insercion_laboral(parquet_dir)
    dataset["satisfaccion"] = extract_satisfaccion(parquet_dir)
    
    return dataset


if __name__ == "__main__":
    # Test print demographics of Career 35
    data = extract_all_data(35)
    print("DEMOGRAPHICS EXTRACTED:")
    print(data["demographics"])
    print("\nMETADATA EXTRACTED:")
    print(data["metadata"])
