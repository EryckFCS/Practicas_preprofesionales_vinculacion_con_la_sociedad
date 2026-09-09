import pandas as pd
df = pd.read_parquet("data/parquet/35/indicadores.parquet")
for idx, val in enumerate(df['variables_encuesta_a_graduados'].tolist()):
    print(f"Row {idx}: {val}")
