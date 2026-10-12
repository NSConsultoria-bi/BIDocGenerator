
import re
from pathlib import Path


RELATIONSHIP_RE = re.compile(
    r"^\s*relationship\s+(.+?)\s*$",
    re.IGNORECASE,
)

PROPERTY_RE = re.compile(
    r"^\s*([A-Za-z][A-Za-z0-9]*)\s*:\s*(.*?)\s*$"
)


PROPERTY_MAPPING = {
    "fromColumn": "from_column",
    "toColumn": "to_column",
    "fromCardinality": "from_cardinality",
    "toCardinality": "to_cardinality",
    "crossFilteringBehavior": "cross_filtering_behavior",
    "securityFilteringBehavior": "security_filtering_behavior",
    "isActive": "is_active",
    "joinOnDateBehavior": "join_on_date_behavior",
    "relyOnReferentialIntegrity": "rely_on_referential_integrity",
}


BOOLEAN_FIELDS = {
    "is_active",
    "rely_on_referential_integrity",
}


def parse_value(value: str):
    value = value.strip()

    if value.lower() == "true":
        return True

    if value.lower() == "false":
        return False

    return value or None


def extract_relationships(semantic_model_dir: Path) -> dict:
    """Extrai relacionamentos e propriedades do arquivo relationships.tmdl."""

    relationships_file = (
        Path(semantic_model_dir)
        / "definition"
        / "relationships.tmdl"
    )

    if not relationships_file.is_file():
        raise ValueError(
            "O arquivo definition/relationships.tmdl não foi encontrado."
        )

    try:
        lines = relationships_file.read_text(
            encoding="utf-8-sig"
        ).splitlines()
    except (OSError, UnicodeDecodeError) as exc:
        raise ValueError(
            f"Não foi possível ler relationships.tmdl: {exc}"
        ) from exc

    relationships = []
    current = None

    def finish_relationship():
        nonlocal current

        if current is not None:
            relationships.append(current)
            current = None

    for line in lines:
        declaration = RELATIONSHIP_RE.match(line)

        if declaration:
            finish_relationship()

            current = {
                "name": declaration.group(1).strip(),
                "from_column": None,
                "to_column": None,
                "from_cardinality": None,
                "to_cardinality": None,
                "cross_filtering_behavior": None,
                "security_filtering_behavior": None,
                "is_active": True,
                "join_on_date_behavior": None,
                "rely_on_referential_integrity": None,
            }
            continue

        if current is None:
            continue

        match = PROPERTY_RE.match(line)

        if not match:
            continue

        source_key, raw_value = match.groups()

        # Ignora propriedades não relevantes para este extrator.
        if source_key not in PROPERTY_MAPPING:
            continue

        target_key = PROPERTY_MAPPING[source_key]
        value = parse_value(raw_value)

        if target_key in BOOLEAN_FIELDS and isinstance(value, str):
            if value.lower() not in {"true", "false"}:
                current[target_key] = value
                continue

        current[target_key] = value

    finish_relationship()

    incomplete = [
        relationship["name"]
        for relationship in relationships
        if not relationship["from_column"]
        or not relationship["to_column"]
    ]

    return {
        "relationship_count": len(relationships),
        "relationships": relationships,
        "incomplete_relationship_count": len(incomplete),
        "incomplete_relationships": incomplete,
    }
