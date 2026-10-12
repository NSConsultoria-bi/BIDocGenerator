
from pathlib import Path


def inspect_project_directory(project_dir: Path) -> dict:
    """
    Inspeciona a estrutura de diretórios de um projeto PBIP.

    Não abre nem executa os arquivos encontrados.
    """
    if not project_dir.exists() or not project_dir.is_dir():
        raise ValueError("O diretório do projeto não existe.")

    report_dirs = sorted(
        path.name
        for path in project_dir.iterdir()
        if path.is_dir() and path.name.endswith(".Report")
    )

    semantic_model_dirs = sorted(
        path.name
        for path in project_dir.iterdir()
        if path.is_dir() and path.name.endswith(".SemanticModel")
    )

    return {
        "report_directories": report_dirs,
        "semantic_model_directories": semantic_model_dirs,
        "report_count": len(report_dirs),
        "semantic_model_count": len(semantic_model_dirs),
    }
