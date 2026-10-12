
import tempfile
from pathlib import Path

from tmdl_extractor import extract_tmdl_model
from report_generator import generate_docx


PROJECT_DIR = Path(
    r"C:\Users\Ewerton\OneDrive\Documentos\MBA EXPERIUM\DATA CHALLENGES\1\CASE_OLIST_PBIP"
)

SEMANTIC_MODEL_DIR = PROJECT_DIR / "CASE OLIST_PBIP.SemanticModel"
REPORT_DIR = PROJECT_DIR / "CASE OLIST_PBIP.Report"


def main():
    model = extract_tmdl_model(SEMANTIC_MODEL_DIR)

    with tempfile.TemporaryDirectory(prefix="ebi_doc_test_") as temp_dir:
        output_path = Path(temp_dir) / "documentacao_olist.docx"

        generate_docx(
            model=model,
            output_path=output_path,
            project_name="CASE OLIST",
            report_dir=REPORT_DIR,
        )

        print("=== TESTE DO GERADOR DOCX ===")
        print(f"Arquivo gerado: {output_path.name}")
        print(f"Tamanho: {output_path.stat().st_size} bytes")
        print(f"Existe: {output_path.is_file()}")


if __name__ == "__main__":
    main()
