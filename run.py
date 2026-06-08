import sys
from src.orchestrator import run_report_pipeline, run_report_batch
from src.config import CAREER_FILES

def print_banner():
    print("=" * 60)
    print("   SISTEMA DE ACTUALIZACIÓN Y VALIDACIÓN DE INFORMES")
    print("        FACULTAD JURÍDICA, SOCIAL Y ADMINISTRATIVA")
    print("=" * 60)

def show_menu():
    print("\nSeleccione una opción de ejecución:")
    print("  [1] Ejecutar un solo informe (Ej: 35)")
    print("  [2] Ejecutar solo pregrado (24 al 34)")
    print("  [3] Ejecutar solo posgrado (35 al 46)")
    print("  [4] Ejecutar ambos por separado (primero pregrado, luego posgrado)")
    print("  [5] Salir")
    print("-" * 60)

def get_valid_choice():
    while True:
        try:
            choice = int(input("Ingrese su opción: "))
            if choice in [1, 2, 3, 4, 5]:
                return choice
            print("Opción inválida. Intente de nuevo.")
        except ValueError:
            print("Por favor, ingrese un número.")

def main():
    print_banner()
    while True:
        show_menu()
        choice = get_valid_choice()
        
        if choice == 1:
            try:
                report_id = int(input("\nIngrese el número del informe a ejecutar (24-46): "))
                if report_id in CAREER_FILES:
                    run_report_pipeline(report_id, force_ingest=True)
                else:
                    print(f"Error: El informe {report_id} no está registrado en config.py.")
            except ValueError:
                print("Error: Por favor ingrese un número de informe válido.")
                
        elif choice == 2:
            print("\nIniciando ejecución de pregrado (24 al 34)...")
            run_report_batch(range(24, 35), force_ingest=True)
            print("\n¡Ejecución de pregrado completada!")
            
        elif choice == 3:
            print("\nIniciando ejecución de posgrado (35 al 46)...")
            run_report_batch(range(35, 47), force_ingest=True)
            print("\n¡Ejecución de posgrado completada!")

        elif choice == 4:
            confirm = input("\n¿Está seguro de querer ejecutar PREGRADO y POSGRADO por separado? (s/n): ").strip().lower()
            if confirm == 's':
                print("\nIniciando ejecución de pregrado (24 al 34)...")
                run_report_batch(range(24, 35), force_ingest=True)
                print("\nIniciando ejecución de posgrado (35 al 46)...")
                run_report_batch(range(35, 47), force_ingest=True)
                print("\n¡Ejecución completa por categoría finalizada!")
            else:
                print("Operación cancelada.")
                
        elif choice == 5:
            print("\nSaliendo del programa. ¡Hasta luego!")
            sys.exit(0)

if __name__ == "__main__":
    # If a command line argument is provided, run that directly
    if len(sys.argv) > 1:
        try:
            rid = int(sys.argv[1])
            if rid in CAREER_FILES:
                run_report_pipeline(rid, force_ingest=True)
            else:
                print(f"Error: Informe {rid} no registrado.")
        except ValueError:
            print("Error: Ingrese un número de informe válido como argumento.")
    else:
        main()
