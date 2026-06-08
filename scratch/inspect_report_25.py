import pandas as pd
import os

parquet_dir = "/home/erick-fcs/Descargas/Practicas_preprofesionales/data/parquet/25"
ind_file = os.path.join(parquet_dir, "indicadores.parquet")

if os.path.exists(ind_file):
    df_ind = pd.read_parquet(ind_file)
    print("--- ALL COLUMNS ---")
    print(df_ind.columns)
    
    print("\n--- FIND 89 IN DATAFRAME ---")
    for col in df_ind.columns:
        matches = df_ind[df_ind[col].astype(str).str.contains("89|89.0", na=False)]
        if not matches.empty:
            print(f"Column: {col}")
            print(matches[[col]])
            
    print("\n--- FULL DATAFRAME ---")
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    print(df_ind)
