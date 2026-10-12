
from pathlib import Path

from backend.tmdl_extractor import extract_tmdl_model
from backend.report_generator import generate_docx


PROJECT_DIR = Path(
    r"C:\Users\Ewerton\OneDrive\Documentos\MBA EXPERIUM\DATA CHALLENGES\1\CASE_OLIST_PBIP"
)

SEMANTIC_MODEL_DIR = PROJECT_DIR / "CASE OLIST_PBIP.SemanticModel"
REPORT_DIR = PROJECT_DIR / "CASE OLIST_PBIP.Report"

# Arquivo de validação salvo na raiz do projeto BI Generator.
OUTPUT_PATH = Path(__file__).resolve().parent.parent / "documentacao_olist_validacao.docx"


def main():
    if not SEMANTIC_MODEL_DIR.is_dir():
        raise FileNotFoundError(
            f"Modelo semântico não encontrado: {SEMANTIC_MODEL_DIR}"
        )

    model = extract_tmdl_model(SEMANTIC_MODEL_DIR)

    generate_docx(
        model=model,
        output_path=OUTPUT_PATH,
        project_name="CASE OLIST",
        report_dir=REPORT_DIR,
    )

    print("=== TESTE DO GERADOR DOCX ===")
    print(f"Arquivo gerado: {OUTPUT_PATH}")
    print(f"Tamanho: {OUTPUT_PATH.stat().st_size} bytes")
    print(f"Existe: {OUTPUT_PATH.is_file()}")


if __name__ == "__main__":
    main()
