import os
import pandas as pd
import unicodedata
import re
from src.config import CAREER_FILES, RESULTADOS_DIR, PREGUNTAS_DIR, PARQUET_CACHE_DIR

def sanitize_column_name(col):
    """
    Cleans column names by removing leading/trailing spaces, converting to lowercase,
    removing accents, and replacing spaces and special characters with underscores.
    """
    if not isinstance(col, str):
        return f"col_{col}"
    # Convert to string, strip whitespace, and lowercase
    name = col.strip().lower()
    # Normalize unicode to remove accents (NFD decomposes characters)
    name = unicodedata.normalize('NFD', name)
    name = "".join([c for c in name if not unicodedata.combining(c)])
    # Replace spaces and punctuation with underscores
    name = re.sub(r'[^a-z0-9_]', '_', name)
    # Replace multiple underscores with a single one
    name = re.sub(r'_+', '_', name)
    # Strip leading/trailing underscores
    name = name.strip('_')
    return name

def sanitize_dataframe(df):
    """Sanitizes columns, removes completely empty rows, and coerces object columns to strings to prevent PyArrow conversion failures."""
    df = df.copy()
    # Clean column names
    df.columns = [sanitize_column_name(c) for c in df.columns]
    # Remove completely empty rows
    df = df.dropna(how="all")
    
    # Coerce any 'object' type columns to string to prevent PyArrow type mismatch errors
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).replace("nan", "")
            # Clean up leading/trailing whitespaces in text columns
            df[col] = df[col].str.strip()
            
    return df


def ingest_career_data(report_id, force=False):
    """
    Converts Excel sheets for the given report_id to Parquet files.
    Also extracts the open questions matching the career's sheet name.
    """
    if report_id not in CAREER_FILES:
        raise ValueError(f"Report ID {report_id} not registered in CAREER_FILES.")

    career = CAREER_FILES[report_id]
    career_parquet_dir = os.path.join(PARQUET_CACHE_DIR, str(report_id))
    os.makedirs(career_parquet_dir, exist_ok=True)

    # 1. Ingest Main Excel file
    excel_path = os.path.join(RESULTADOS_DIR, career["excel"])
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Excel results file not found: {excel_path}")

    print(f"[Ingest] Processing main Excel for Career {report_id}: {career['excel']}")
    xl = pd.ExcelFile(excel_path)
    for sheet in xl.sheet_names:
        parquet_path = os.path.join(career_parquet_dir, f"{sanitize_column_name(sheet)}.parquet")
        if not force and os.path.exists(parquet_path):
            continue
        try:
            # We don't skip rows during basic ingestion, extractor will handle row-skips if needed
            df = xl.parse(sheet)
            df = sanitize_dataframe(df)
            df.to_parquet(parquet_path, index=False)
            print(f"  Saved sheet '{sheet}' -> {os.path.basename(parquet_path)}")
        except Exception as e:
            print(f"  Error ingesting sheet '{sheet}': {e}")

    # 2. Ingest Open Questions
    print(f"[Ingest] Extracting open questions for Career {report_id} (Career name: {career['open_career_name']})")
    if os.path.exists(PREGUNTAS_DIR):
        for file in os.listdir(PREGUNTAS_DIR):
            if file.endswith(".xlsx") or file.endswith(".xls"):
                file_path = os.path.join(PREGUNTAS_DIR, file)
                base_name = sanitize_column_name(os.path.splitext(file)[0])
                parquet_path = os.path.join(career_parquet_dir, f"pregunta_abierta_{base_name}.parquet")
                
                if not force and os.path.exists(parquet_path):
                    continue

                try:
                    xl_q = pd.ExcelFile(file_path)
                    # We know they have 'Hoja1'
                    sheet_name = 'Hoja1' if 'Hoja1' in xl_q.sheet_names else xl_q.sheet_names[0]
                    df_q = xl_q.parse(sheet_name)
                    df_q = sanitize_dataframe(df_q)

                    # Identify the career column (usually 'carrera_o_programa' or 'w')
                    career_col = [c for c in df_q.columns if "carrera" in c or "programa" in c or c == "w"]
                    if career_col:
                        col = career_col[0]
                        # Filter rows matching the career name (case-insensitive strip comparison)
                        target_name = career["open_career_name"].strip().lower()
                        filtered_df = df_q[df_q[col].astype(str).str.strip().str.lower() == target_name].copy()
                        
                        # Save the filtered dataframe
                        filtered_df.to_parquet(parquet_path, index=False)
                        print(f"  Saved open question '{file}' (Filtered for '{career['open_career_name']}', Rows: {len(filtered_df)}) -> {os.path.basename(parquet_path)}")
                    else:
                        print(f"  Warning: No career column found in {file}. Columns: {df_q.columns.tolist()}")
                except Exception as e:
                    print(f"  Error reading open question '{file}': {e}")
    print(f"[Ingest] Ingestion completed for Career {report_id}.")

if __name__ == "__main__":
    # Test run for Career 35
    ingest_career_data(35, force=True)
