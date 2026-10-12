
import json
from pathlib import Path, PurePosixPath, PureWindowsPath


def validate_pbip_file(file_path: Path) -> dict:
    """Valida a estrutura básica de um arquivo de projeto PBIP."""

    if not file_path.is_file():
        raise ValueError("O arquivo informado não existe.")

    if file_path.suffix.lower() != ".pbip":
        raise ValueError("O arquivo precisa ter extensão .pbip.")

    try:
        with file_path.open("r", encoding="utf-8-sig") as file:
            content = json.load(file)
    except (UnicodeDecodeError, json.JSONDecodeError):
        raise ValueError("O arquivo PBIP não contém JSON válido.") from None

    if not isinstance(content, dict):
        raise ValueError("A raiz do PBIP deve ser um objeto JSON.")

    version = content.get("version")

    if not isinstance(version, str) or not version.strip():
        raise ValueError("Versão do PBIP ausente ou inválida.")

    artifacts = content.get("artifacts")

    if not isinstance(artifacts, list) or not artifacts:
        raise ValueError("O PBIP não contém artefatos válidos.")

    validated_artifacts = []

    for index, artifact in enumerate(artifacts):
        if not isinstance(artifact, dict):
            raise ValueError(f"O artefato {index} não é um objeto JSON.")

        report = artifact.get("report")

        if not isinstance(report, dict):
            raise ValueError(
                f"O artefato {index} não contém um objeto 'report' válido."
            )

        report_path = report.get("path")

        if not isinstance(report_path, str) or not report_path.strip():
            raise ValueError(
                f"O artefato {index} não contém 'report.path' válido."
            )

        # Aceita somente caminhos relativos simples para pastas do projeto.
        posix_path = PurePosixPath(report_path)
        windows_path = PureWindowsPath(report_path)

        if (
            posix_path.is_absolute()
            or windows_path.is_absolute()
            or ".." in posix_path.parts
            or ".." in windows_path.parts
            or ":" in report_path
        ):
            raise ValueError(
                f"O artefato {index} contém um caminho inseguro."
            )

        if len(posix_path.parts) != 1 or report_path in {".", ".."}:
            raise ValueError(
                f"O caminho do artefato {index} deve indicar uma pasta relativa."
            )

        validated_artifacts.append({
            "type": "report",
            "path": report_path,
        })

    return {
        "valid": True,
        "version": version,
        "artifact_count": len(validated_artifacts),
        "artifacts": validated_artifacts,
        "settings_present": isinstance(content.get("settings", {}), dict),
    }
