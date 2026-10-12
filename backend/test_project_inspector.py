
from pathlib import Path

from project_inspector import inspect_project_directory


PROJECT_DIR = Path(
    r"C:\Users\Ewerton\OneDrive\Documentos\MBA EXPERIUM\DATA CHALLENGES\1\CASE_OLIST_PBIP"
)


def main():
    result = inspect_project_directory(PROJECT_DIR)

    print("=== INSPEÇÃO DO PROJETO PBIP ===")
    print(f"Pastas de relatório: {result['report_count']}")
    print(f"Pastas de modelo semântico: {result['semantic_model_count']}")

    print("\n=== RELATÓRIOS ===")
    for name in result["report_directories"]:
        print(f"- {name}")

    print("\n=== MODELOS SEMÂNTICOS ===")
    for name in result["semantic_model_directories"]:
        print(f"- {name}")


if __name__ == "__main__":
    main()
