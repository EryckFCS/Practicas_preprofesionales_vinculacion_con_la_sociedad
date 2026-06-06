import pandas as pd
import os

def parse_section(df, section_name):
    # Find the row index of the section
    start_idx = None
    for idx, val in enumerate(df.iloc[:, 0].astype(str).tolist()):
        if section_name in val:
            start_idx = idx
            break
    
    if start_idx is None:
        return {}

    # The options start after 'Opción' row (usually 2 rows down: section_name, question, Opción)
    op_idx = None
    for i in range(start_idx + 1, min(start_idx + 5, len(df))):
        if str(df.iloc[i, 0]).strip().lower() == 'opción':
            op_idx = i
            break
            
    if op_idx is None:
        return {}

    results = {}
    # Extract options until next blank line or next "Resumen de campo"
    for i in range(op_idx + 1, len(df)):
        val_0 = str(df.iloc[i, 0]).strip()
        # Stop conditions
        if val_0 == 'nan' or val_0 == '' or 'resumen de campo' in val_0.lower():
            break
        
        count = str(df.iloc[i, 1]).strip()
        pct = str(df.iloc[i, 2]).strip()
        results[val_0] = {
            'count': int(count.split('.')[0]) if count != 'nan' and count != '' else 0,
            'pct': float(pct) if pct != 'nan' and pct != '' else 0.0
        }
    return results

def extract_employer_data_test(report_id):
    parquet_path = f"data/parquet/{report_id}/datae.parquet"
    if not os.path.exists(parquet_path):
        return None
    df = pd.read_parquet(parquet_path)
    
    b1 = parse_section(df, "Resumen de campo para B1")
    b2 = parse_section(df, "Resumen de campo para B2")
    b3 = parse_section(df, "Resumen de campo para B3")
    
    # D1(SQ01) to D1(SQ09)
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

if __name__ == '__main__':
    data = extract_employer_data_test(24)
    print("B1 Requisitos:")
    print(data["b1"])
    print("\nB2 Aspectos:")
    print(data["b2"])
    print("\nB3 Demanda:")
    print(data["b3"])
    print("\nD1 Performance General (SQ01):")
    print(data["d1"]["sq1"])
