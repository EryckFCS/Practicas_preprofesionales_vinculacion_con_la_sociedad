# Report Synchronization and Validation Tool

A modular, robust, and highly reliable Python automation utility to synchronize survey data from Excel files (`.xlsx` and LimeSurvey `"RG"` files) with Word report documents (`.docx`), specifically built for the **Unidad de Seguimiento a Graduados (USG)** of the Faculty of Juridical, Social, and Administrative Sciences.

---

## 🌟 Key Features

1.  **Excel Ingestion to Parquet Cache**: Converts mixed-type raw Excel sheets into type-safe, highly efficient Apache Parquet databases before processing.
2.  **Open-Ended Question Database Filtering**: Parses and extracts student text responses by dynamically filtering the combined 1964-row open questions database matching the career's exact name.
3.  **Adaptive Postgraduate Layout Support**: Dynamically detects whether a report ID corresponds to an undergraduate ($\le 34$) or postgraduate/master ($\ge 35$) program, adjusting statistical terminology (e.g. `"maestrantes"` and `"Maestría"`) and paragraph templates accordingly.
4.  **Raw Data Safety**: Originals are **never** modified in-place. Source documents are safely duplicated inside a dedicated output directory (`processed/`), and all operations are safely performed there.
5.  **Automated Unit Testing & Linting**: Built-in automated unit tests validating extraction, open questions, and safe document dry-run modification with a completely clean linter validation report.

---

## 📁 Directory Structure

```
practicas_preprofesionales/
├── src/
│   ├── config.py         # Mappings, paths configuration, and postgraduate validations
│   ├── ingest.py         # Excel-to-Parquet conversion & sanitization pipeline
│   ├── extractor.py      # Sanitized data extraction from Parquet caches
│   ├── updater.py        # Word document modifier (tables cells & paragraph interpretations)
│   └── orchestrator.py   # Driver pipeline coordinating ingestion, extraction, and updates
├── tests/
│   └── test_pipeline.py  # Automated unit tests
├── processed/            # Safe output directory containing the updated reports
├── data/                 # Sanitized Parquet database cache
├── run.py                # User-friendly interactive command-line menu runner
├── pyproject.toml        # Declarative project dependencies metadata
└── uv.lock               # Generated deterministic lockfile
```

---

## 🚀 Quick Start

Ensure you have your environment resolved:

```bash
# Verify dependency lock file is active
uv sync
```

### Run the Interactive Menu

Start the command-line menu to process single reports, all pending posgrado reports, or all files:

```bash
PYTHONPATH=. .venv/bin/python run.py
```

*   **Option 1**: Update a single report (e.g. enter `35`).
*   **Option 2**: **(Recommended)** Update the entire pending posgrado range (`35` to `46`).
*   **Option 3**: Update all 23 reports (`24` to `46`).

### Run the Pipeline from CLI directly

You can also run the orchestrator for a specific report ID directly from the command line:

```bash
PYTHONPATH=. .venv/bin/python run.py 35
```

---

## 🧪 Automated Testing & Linting

Run automated unit tests to verify the pipeline's extraction and modification precision:

```bash
PYTHONPATH=. .venv/bin/python -m unittest tests/test_pipeline.py
```

Run static analysis checks using the Ruff linter:

```bash
.venv/bin/ruff check .
```
