import re
from pathlib import Path

from backend.relationship_extractor import extract_relationships


TABLE_RE = re.compile(r"^table\s+(.+?)\s*$", re.IGNORECASE)

MEASURE_RE = re.compile(
    r"^\s*measure\s+(.+?)\s*=\s*(.*)$",
    re.IGNORECASE,
)

# Captura a declaração da coluna e uma possível expressão na mesma linha.
COLUMN_RE = re.compile(
    r"^\s*column\s+(.+?)(?:\s*=\s*(.*))?$",
    re.IGNORECASE,
)

PROPERTY_RE = re.compile(
    r"^\s*(?:"
    r"formatString|formatStringDefinition|displayFolder|description|"
    r"isHidden|lineageTag|annotation|extendedProperty|"
    r"dataType|summarizeBy|sourceColumn|sortByColumn|dataCategory|"
    r"changedProperty|variation|encodingHint|summarizeBy"
    r")\s*(?::|=|\b)",
    re.IGNORECASE,
)


def clean_name(value: str) -> str:
    """Remove espaços e aspas externas de nomes TMDL."""

    value = value.strip()

    if len(value) >= 2 and value[0] == "'" and value[-1] == "'":
        return value[1:-1].replace("\\'", "'")

    return value


def extract_tmdl_model(semantic_model_dir: Path) -> dict:
    """
    Extrai tabelas, colunas, medidas e relacionamentos
    de um modelo semântico Power BI no formato TMDL.

    Uma coluna é considerada calculada quando sua declaração TMDL
    contém uma expressão DAX, na mesma linha ou em linhas seguintes.
    """

    semantic_model_dir = Path(semantic_model_dir)
    definition_dir = semantic_model_dir / "definition"
    tables_dir = definition_dir / "tables"

    if not definition_dir.is_dir():
        raise ValueError(
            "A pasta definition do modelo não foi encontrada."
        )

    if not tables_dir.is_dir():
        raise ValueError(
            "A pasta definition/tables não foi encontrada."
        )

    table_files = sorted(tables_dir.glob("*.tmdl"))

    tables = []
    columns = []
    measures = []
    errors = []

    for file_path in table_files:
        try:
            lines = file_path.read_text(
                encoding="utf-8-sig"
            ).splitlines()
        except (OSError, UnicodeDecodeError) as exc:
            errors.append({
                "file": file_path.name,
                "error": str(exc),
            })
            continue

        current_table = None
        current_measure = None
        current_column = None

        def finish_measure():
            """Finaliza e registra a medida atualmente em leitura."""
            nonlocal current_measure

            if current_measure is None:
                return

            expression_lines = current_measure.pop(
                "_expression_lines"
            )

            current_measure["expression"] = "\n".join(
                expression_lines
            ).strip()

            measures.append(current_measure)
            current_measure = None

        def finish_column():
            """Finaliza a coluna e registra se possui expressão DAX."""
            nonlocal current_column

            if current_column is None:
                return

            expression = "\n".join(
                current_column.pop("_expression_lines")
            ).strip()

            # Uma expressão não vazia, associada à declaração column,
            # identifica uma coluna calculada em DAX.
            current_column["expression"] = expression
            current_column["is_calculated"] = bool(expression)
            columns.append(current_column)
            current_column = None

        for line in lines:
            # Declarações de tabela são de primeiro nível.
            table_match = (
                TABLE_RE.match(line)
                if line and not line[0].isspace()
                else None
            )

            if table_match:
                finish_measure()
                finish_column()
                current_table = clean_name(table_match.group(1))

                tables.append({
                    "name": current_table,
                    "file": file_path.name,
                })
                continue

            if current_table is None:
                continue

            # Reconhece uma nova declaração de medida.
            measure_match = MEASURE_RE.match(line)

            if measure_match:
                finish_column()
                finish_measure()

                initial_expression = measure_match.group(2).strip()

                # Delimitadores Markdown não fazem parte do DAX.
                if initial_expression == "```":
                    initial_expression = ""

                current_measure = {
                    "name": clean_name(measure_match.group(1)),
                    "table": current_table,
                    "file": file_path.name,
                    "_expression_lines": (
                        [initial_expression]
                        if initial_expression
                        else []
                    ),
                }
                continue

            # Reconhece declarações de coluna.
            column_match = COLUMN_RE.match(line)

            if column_match:
                finish_measure()
                finish_column()

                column_name = column_match.group(1).strip()
                initial_expression = (column_match.group(2) or "").strip()

                # TMDL pode delimitar expressões multilinha com ```.
                if initial_expression == "```":
                    initial_expression = ""

                current_column = {
                    "name": clean_name(column_name),
                    "table": current_table,
                    "file": file_path.name,
                    "_expression_lines": (
                        [initial_expression]
                        if initial_expression
                        else []
                    ),
                }
                continue

            # Propriedades TMDL encerram a expressão atual.
            if PROPERTY_RE.match(line):
                finish_measure()
                finish_column()
                continue

            stripped_line = line.strip()

            # Ignora delimitadores TMDL de expressão multilinha.
            if stripped_line == "```":
                continue

            if current_measure is not None:
                current_measure["_expression_lines"].append(stripped_line)
            elif current_column is not None:
                current_column["_expression_lines"].append(stripped_line)

        finish_measure()
        finish_column()

    # Extrai relacionamentos usando o módulo específico.
    relationships = {
        "relationship_count": 0,
        "relationships": [],
        "incomplete_relationship_count": 0,
        "incomplete_relationships": [],
    }

    relationships_file = definition_dir / "relationships.tmdl"

    if relationships_file.is_file():
        try:
            relationships = extract_relationships(semantic_model_dir)
        except (OSError, UnicodeDecodeError, ValueError) as exc:
            errors.append({
                "file": "relationships.tmdl",
                "error": str(exc),
            })

    return {
        "table_count": len(tables),
        "column_count": len(columns),
        "calculated_column_count": sum(
            1 for column in columns if column["is_calculated"]
        ),
        "measure_count": len(measures),
        "relationship_count": relationships["relationship_count"],
        "incomplete_relationship_count": (
            relationships["incomplete_relationship_count"]
        ),
        "tables": tables,
        "columns": columns,
        "measures": measures,
        "relationships": relationships["relationships"],
        "incomplete_relationships": (
            relationships["incomplete_relationships"]
        ),
        "model_file_exists": (
            definition_dir / "model.tmdl"
        ).is_file(),
        "relationships_file_exists": relationships_file.is_file(),
        "files_scanned": len(table_files),
        "errors": errors,
    }
