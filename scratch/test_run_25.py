from src.orchestrator import run_report_pipeline
import docx
import os

# Run the pipeline for report 25
run_report_pipeline(25, force_ingest=True)

processed_path = "/home/erick-fcs/Descargas/Practicas_preprofesionales/processed/pregrado/25.  INFORME SG DERECHO 2023.docx"
if not os.path.exists(processed_path):
    print("Processed file not found at:", processed_path)
    exit(1)

doc = docx.Document(processed_path)
print("\n--- PROCESSED PARAGRAPHS ---")
for idx, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if "3.5" in text or "E1" in text or "capacitación" in text.lower() or "Los empleadores consultados sugirieron" in text or "•" in text or "Conocimientos" in text:
        print(f"[{idx}]: {text}")
