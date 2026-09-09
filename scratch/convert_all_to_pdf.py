import subprocess
import os
from src.config import CAREER_FILES

print("=== CONVERTING ALL POSTGRAD REPORTS TO PDF ===")

for r_id in range(35, 47):
    career = CAREER_FILES[r_id]
    word_file = career["word"]
    word_path = f"processed/posgrado/{word_file}"
    
    if not os.path.exists(word_path):
        print(f"Report {r_id}: Word file not found at {word_path}")
        continue
        
    print(f"Converting Report {r_id}: {word_file}...")
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", "processed/posgrado/",
        word_path
    ]
    try:
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        print(f"  Successfully converted Report {r_id}")
    except Exception as e:
        print(f"  Error converting Report {r_id}: {e}")
