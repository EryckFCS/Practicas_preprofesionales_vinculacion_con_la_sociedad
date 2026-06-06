import pandas as pd
import os

def extract_correlation(report_id):
    parquet_path = f"data/parquet/{report_id}/correlacion.parquet"
    if not os.path.exists(parquet_path):
        return {}
    df = pd.read_parquet(parquet_path)
    
    # Let's find rows for Graduates and Employers
    # We look for "CORRELACION EMPLEADORES" in any cell
    split_idx = None
    for idx, row in df.iterrows():
        row_vals = [str(x).upper() for x in row.values]
        if any("CORRELACION EMPLEADORES" in val or "CORRELACIÓN EMPLEADORES" in val for val in row_vals):
            split_idx = idx
            break
            
    if split_idx is None:
        split_idx = 12 # fallback
        
    grad_data = {}
    emp_data = {}
    
    # Column indices:
    # Column 6: correlacion_graduados (labels)
    # Column 7: unnamed_7 (values)
    # Column 8: unnamed_8 (levels)
    
    for idx, row in df.iterrows():
        # Label is in column 6
        label = str(row.iloc[6]).strip()
        val = row.iloc[7]
        lvl = str(row.iloc[8]).strip()
        
        # Skip header or empty
        if label == 'nan' or label == '' or 'detalle' in label.lower() or 'correlacion' in label.lower():
            continue
            
        # Parse value
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

if __name__ == '__main__':
    data = extract_correlation(24)
    print("GRADUADOS:")
    for k, v in data['graduados'].items():
        print(f"  {k}: {v}")
    print("\nEMPLEADORES:")
    for k, v in data['empleadores'].items():
        print(f"  {k}: {v}")
