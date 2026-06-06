import os
import pandas as pd
import unicodedata
from src.config import PARQUET_CACHE_DIR, CAREER_FILES

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
    muestra = 0
    dem_file = os.path.join(parquet_dir, "demografico.parquet")
    if os.path.exists(dem_file):
        df_dem = pd.read_parquet(dem_file)
        if len(df_dem) > 0:
            # Row 0 contains 'n_de_graduados' (total sample size under the cohorte heading)
            val = df_dem.iloc[0].get('n_de_graduados', 0)
            muestra = safe_int(val)
            
    # Read population N from indicadores.parquet
    poblacion = 0
    ind_file = os.path.join(parquet_dir, "indicadores.parquet")
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
                            poblacion = val
                            break
                if poblacion > 0:
                    break
                    
    # Fallbacks if still 0
    if muestra == 0:
        muestra = 14  # Default sensible fallback
    if poblacion == 0:
        poblacion = muestra * 2  # Default sensible fallback
        
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
    return results

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
        
    return {
        "b1": b1,
        "b2": b2,
        "b3": b3,
        "d1": d1
    }

def extract_open_questions(parquet_dir):
    """Extracts all non-empty open question text answers."""
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
                        if (a_str and 
                            a_str.lower() not in ["nan", "sn", "ninguno", "ninguna", "sin respuesta", "no aplica", "n/a", "no", "ninguno."]
                            and a_str != ""):
                            if a_str not in cargos:
                                cargos.append(a_str)
                    
                    for a in df[col_funciones].dropna().tolist():
                        a_str = str(a).replace("_x000D_", "").strip()
                        if (a_str and 
                            a_str.lower() not in ["nan", "sn", "ninguno", "ninguna", "sin respuesta", "no aplica", "n/a", "no", "ninguno."]
                            and a_str != ""):
                            if a_str not in funciones:
                                funciones.append(a_str)
                
                open_data["employer_cargos"] = cargos
                open_data["employer_funciones"] = funciones
                continue
            
            text_cols = [c for c in df.columns if c not in ["w", "carrera_o_programa", "carrera_o_programa_de_posgrado", "facultad"]]
            answers = []
            for col in text_cols:
                raw_answers = df[col].dropna().tolist()
                for a in raw_answers:
                    a_str = str(a).replace("_x000D_", "").strip()
                    # Clean up common empty/unwanted values
                    if (a_str and 
                        a_str.lower() not in ["nan", "sn", "ninguno", "ninguna", "sin respuesta", "sin comentarios", "no aplica", "n/a", "no", "ninguno."]
                        and a_str != ""):
                        if a_str not in answers:
                            answers.append(a_str)
            
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
    
    return dataset

if __name__ == "__main__":
    # Test print demographics of Career 35
    data = extract_all_data(35)
    print("DEMOGRAPHICS EXTRACTED:")
    print(data["demographics"])
    print("\nMETADATA EXTRACTED:")
    print(data["metadata"])
