from src.ingest import ingest_career_data
from src.extractor import extract_all_data
from src.updater import update_report

def run_report_pipeline(report_id, force_ingest=False, target_docx_path=None):
    """
    Coordinates ingestion, extraction, and update pipeline for a single report ID.
    """
    print("\n==================================================")
    print(f"Starting pipeline for Report ID {report_id}...")
    print("==================================================")
    
    # 1. Ingest Excel and Open Questions databases to Parquet
    ingest_career_data(report_id, force=force_ingest)
    
    # 2. Extract dataset
    print("[Orchestrator] Extracting sanitized data...")
    data = extract_all_data(report_id)
    
    # 3. Perform consistency checks
    muestra = data["metadata"]["muestra"]
    hombre = data["demographics"]["hombre_freq"]
    mujer = data["demographics"]["mujer_freq"]
    total_genero = hombre + mujer
    
    print("[Orchestrator] Performing consistency checks:")
    print(f"  Sample size (n): {muestra}")
    print(f"  Gender count sum (Male: {hombre} + Female: {mujer}): {total_genero}")
    if total_genero != muestra:
        print(f"  [Warning] Gender sum ({total_genero}) does not match sample size ({muestra})!")
    else:
        print("  [Success] Gender distribution sum is consistent.")

    # 4. Update Word Document
    print("[Orchestrator] Updating Word report...")
    update_report(report_id, data, target_docx_path=target_docx_path)
    
    print("==================================================")
    print(f"Pipeline completed successfully for Report ID {report_id}!")
    print("==================================================\n")

if __name__ == "__main__":
    # Test orchestrator for Career 35
    run_report_pipeline(35, force_ingest=True, target_docx_path="INFORMES 24-46 ERICK/35. INFORME SG REPROD. ANIMAL 2025.docx")
