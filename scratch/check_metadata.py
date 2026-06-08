import os
import pandas as pd
from src.extractor import extract_metadata
from src.config import CAREER_FILES

for rid in sorted(CAREER_FILES.keys()):
    if rid >= 35:
        p_dir = f"data/parquet/{rid}"
        meta = extract_metadata(rid, p_dir)
        print(f"Report {rid}: {meta}")
