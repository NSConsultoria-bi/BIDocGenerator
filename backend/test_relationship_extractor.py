
import json
from pathlib import Path

from relationship_extractor import extract_relationships


PROJECT_DIR = Path(
    r"C:\Users\Ewerton\OneDrive\Documentos\MBA EXPERIUM\DATA CHALLENGES\1\CASE_OLIST_PBIP"
)

SEMANTIC_MODEL_DIR = (
    PROJECT_DIR / "CASE OLIST_PBIP.SemanticModel"
)


def main():
    result = extract_relationships(SEMANTIC_MODEL_DIR)

    print("\n=== RESUMO DOS RELACIONAMENTOS ===")
    print(f"Relacionamentos: {result['relationship_count']}")
    print(
        "Relacionamentos incompletos: "
        f"{result['incomplete_relationship_count']}"
    )

  
    for relationship in result["relationships"]:
        print(
            f"- {relationship['from_column']} -> "
            f"{relationship['to_column']} | "
            f"Ativo: {relationship['is_active']} | "
            f"Cardinalidade origem: {relationship['from_cardinality']} | "
            f"Cardinalidade destino: {relationship['to_cardinality']} | "
            f"Filtro: {relationship['cross_filtering_behavior']} | "
            f"Comportamento de data: {relationship['join_on_date_behavior']}"
        )


    output_file = Path(__file__).parent / "relationships_test_result.json"
    output_file.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\nResultado salvo em: {output_file}")


if __name__ == "__main__":
    main()
