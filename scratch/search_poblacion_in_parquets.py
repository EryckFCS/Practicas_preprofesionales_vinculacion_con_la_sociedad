import os
import pandas as pd

expected_poblaciones = {
    35: 14,
    36: 16,
    37: 15,
    38: 20,
    39: 24,
    40: 24,
    41: 13,
    42: 14,
    43: 22,
    44: 35,
    45: 19,
    46: 64
}

for rid, expected in expected_poblaciones.items():
    p_dir = f"data/parquet/{rid}"
    if not os.path.exists(p_dir):
        continue
    print(f"\n=== Report {rid} ===")
    found = False
    for file in os.listdir(p_dir):
        if file.endswith(".parquet"):
            df = pd.read_parquet(os.path.join(p_dir, file))
            # search for the expected value in all cells
            for col in df.columns:
                matches = df[df[col] == expected]
                if not matches.empty:
                    print(f"  Found in {file}, Column: {col}")
                    print(matches.head(2).to_string())
                    found = True
    if not found:
        print("  Not found in any parquet file!")
