
import json
import re
from pathlib import Path

from tmdl_extractor import extract_tmdl_model


# Caminho do projeto Power BI no seu computador.
PROJECT_DIR = Path(
    r"C:\Users\Ewerton\OneDrive\Documentos\MBA EXPERIUM\DATA CHALLENGES\1\CASE_OLIST_PBIP"
)

SEMANTIC_MODEL_DIR = (
    PROJECT_DIR / "CASE OLIST_PBIP.SemanticModel"
)


def inspect_measure_source():
    """Exibe a declaração original de Faturamento_Total."""

    source = (
        SEMANTIC_MODEL_DIR
        / "definition"
        / "tables"
        / "Medidas.tmdl"
    )

    print("\n=== FONTE ORIGINAL: Faturamento_Total ===")
    print(f"Arquivo: {source}")

    if not source.is_file():
        print("ERRO: arquivo Medidas.tmdl não encontrado.")
        return

    try:
        lines = source.read_text(
            encoding="utf-8-sig"
        ).splitlines()
    except (OSError, UnicodeDecodeError) as exc:
        print(f"ERRO ao ler o arquivo: {exc}")
        return

    target = "measure Faturamento_Total ="

    for index, line in enumerate(lines):
        if line.strip().casefold() == target.casefold():
            for number, source_line in enumerate(
                lines[index:index + 15],
                start=index + 1,
            ):
                print(f"{number}: {source_line!r}")
            return

    print("Medida Faturamento_Total não encontrada na fonte.")


def inspect_extracted_measure(result):
    """Exibe a expressão extraída da medida Faturamento_Total."""

    print("\n=== RESULTADO EXTRAÍDO: Faturamento_Total ===")

    for measure in result.get("measures", []):
        if measure["name"].casefold() == "faturamento_total":
            print(json.dumps(
                measure,
                ensure_ascii=False,
                indent=2,
            ))
            return

    print("Medida não encontrada no resultado do extrator.")



def validate_extracted_measures(result):
    """Executa verificações básicas de integridade das medidas DAX."""

    measures = result.get("measures", [])

    empty = []
    suspicious = []
    possible_properties = []

    property_pattern = re.compile(
        r"(?im)^\s*(?:"
        r"formatString|formatStringDefinition|displayFolder|"
        r"lineageTag|annotation|extendedProperty|"
        r"sourceColumn|summarizeBy|dataType"
        r")\s*(?::|=)"
    )

    for measure in measures:
        name = measure.get("name", "(sem nome)")
        expression = measure.get("expression", "")
        stripped = expression.strip()

        if not stripped:
            empty.append(name)

        if "```" in expression:
            suspicious.append(name)

        if property_pattern.search(expression):
            possible_properties.append(name)

    print("\n=== VALIDAÇÃO DAS MEDIDAS ===")
    print(f"Total de medidas: {len(measures)}")
    print(f"Expressões vazias: {len(empty)}")
    print(f"Expressões com crases residuais: {len(suspicious)}")
    print(
        "Expressões com possíveis propriedades TMDL: "
        f"{len(possible_properties)}"
    )

    for label, names in [
        ("Vazias", empty),
        ("Com crases", suspicious),
        ("Com possíveis propriedades TMDL", possible_properties),
    ]:
        if names:
            print(f"\n{label} (máximo 20):")
            for name in names[:20]:
                print(f"- {name}")



def inspect_suspicious_measures(result):
    """Localiza declarações de medidas por nome, sem exigir igualdade exata."""

    source = (
        SEMANTIC_MODEL_DIR
        / "definition"
        / "tables"
        / "Medidas.tmdl"
    )

    lines = source.read_text(encoding="utf-8-sig").splitlines()

    print("\n=== DIAGNÓSTICO DAS MEDIDAS COM CRases ===")

    suspicious = [
        measure
        for measure in result.get("measures", [])
        if measure.get("expression", "").strip().startswith("```")
    ]

    for measure in suspicious[:5]:
        name = measure["name"].casefold()

        print(f"\nMedida: {measure['name']}")
        print(f"Expressão extraída: {measure['expression'][:160]!r}")

        matches = [
            index
            for index, line in enumerate(lines)
            if line.strip().casefold().startswith("measure ")
            and name in line.casefold()
        ]

        if not matches:
            print("Nenhuma declaração correspondente localizada.")
            continue

        for index in matches[:2]:
            print("Declaração original:")

            for source_line in lines[index:index + 10]:
                print(repr(source_line))


def inspect_backticks_in_source():
    """Verifica se existem delimitadores de crases no TMDL original."""

    source = (
        SEMANTIC_MODEL_DIR
        / "definition"
        / "tables"
        / "Medidas.tmdl"
    )

    print("\n=== VERIFICAÇÃO DE CRases NA FONTE ===")

    if not source.is_file():
        print("Arquivo Medidas.tmdl não encontrado.")
        return

    lines = source.read_text(encoding="utf-8-sig").splitlines()

    matches = [
        (number, line)
        for number, line in enumerate(lines, start=1)
        if "```" in line
    ]

    print(f"Linhas contendo crases triplas: {len(matches)}")

    for number, line in matches[:20]:
        print(f"Linha {number}: {line!r}")


def main():
    print("=== TESTE DO EXTRATOR TMDL ===")
    print(f"Projeto: {PROJECT_DIR}")
    print(f"Modelo semântico: {SEMANTIC_MODEL_DIR}")

    if not SEMANTIC_MODEL_DIR.is_dir():
        print(
            "\nERRO: pasta do modelo semântico não encontrada. "
            "Corrija PROJECT_DIR no início do arquivo."
        )
        return

    try:
        result = extract_tmdl_model(SEMANTIC_MODEL_DIR)
    except Exception as exc:
        print(
            f"\nERRO durante a extração: "
            f"{type(exc).__name__}: {exc}"
        )
        raise

    print("\n=== RESUMO DA EXTRAÇÃO ===")
    print(f"Arquivos examinados: {result['files_scanned']}")
    print(f"Tabelas identificadas: {result['table_count']}")
    print(f"Colunas identificadas: {result['column_count']}")
    print(f"Medidas identificadas: {result['measure_count']}")
    print(f"Relacionamentos identificados: {result['relationship_count']}")
    print(
    "Relacionamentos incompletos: "
    f"{result['incomplete_relationship_count']}"
        )
    print(
        "Arquivo de relacionamentos encontrado: "
        f"{result['relationships_file_exists']}"
    )
    print(f"Erros de leitura: {len(result['errors'])}")

    validate_extracted_measures(result)
    inspect_suspicious_measures(result)
    inspect_backticks_in_source()
    inspect_measure_source()
    inspect_extracted_measure(result)

    output_file = (
        Path(__file__).parent / "tmdl_test_result.json"
    )

    output_file.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\nResultado completo salvo em: {output_file}")
    
    print("\n=== DIAGNÓSTICO DE TABELAS E MEDIDAS ===")

    for table in result["tables"]:
        table_name = table["name"]
        measure_count = sum(
            1
            for measure in result["measures"]
            if measure["table"] == table_name
        )
        column_count = sum(
            1
            for column in result["columns"]
            if column["table"] == table_name
        )

        print(
            f"{table_name} | "
            f"Colunas: {column_count} | "
            f"Medidas: {measure_count}"
        )

    print("\n=== RELACIONAMENTOS COM TABELAS AUXILIARES ===")

    for relationship in result["relationships"]:
        source = relationship["from_column"]
        target = relationship["to_column"]

        if "LocalDateTable_" in source or "LocalDateTable_" in target:
            print(f"{source} -> {target}")


if __name__ == "__main__":
    main()
