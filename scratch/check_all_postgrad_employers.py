import pandas as pd
import os
from src.config import CAREER_FILES

print("=== CHECKING EMPLOYER SURVEY RESPONDENT COUNTS PER POSTGRAD CAREER ===")
for r_id in range(35, 47):
    career = CAREER_FILES[r_id]
    open_career_name = career["open_career_name"]
    parquet_path = f"data/parquet/{r_id}/datae.parquet"
    if not os.path.exists(parquet_path):
        print(f"Report {r_id}: Parquet file {parquet_path} does not exist.")
        continue
    
    df = pd.read_parquet(parquet_path)
    # Search for row containing open_career_name
    found = False
    for idx, row in df.iterrows():
        row_str = " | ".join(str(val) for val in row.values)
        if open_career_name.lower() in row_str.lower():
            # Usually the second column is frequency/count
            print(f"Report {r_id} ({career['open_sheet']}): Row found: {row_str[:120]}")
            found = True
            break
    if not found:
        print(f"Report {r_id} ({career['open_sheet']}): Career '{open_career_name}' NOT found in datae.")
