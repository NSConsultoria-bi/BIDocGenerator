
import json
import shutil
import stat
import tempfile
import zipfile
from pathlib import Path, PurePosixPath, PureWindowsPath

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response

from backend.pbip_validator import validate_pbip_file
from backend.project_inspector import inspect_project_directory
from backend.tmdl_extractor import extract_tmdl_model
from backend.report_generator import generate_docx


app = FastAPI(
    title="E-BI DOC Generator",
    description="Backend para análise e documentação de projetos Power BI.",
    version="0.4.0",
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5500",
        "http://localhost:5500",
        "http://127.0.0.1:5501",
        "http://localhost:5501",
        "null",
    ],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


MAX_UPLOAD_SIZE = 200 * 1024 * 1024
MAX_EXTRACTED_SIZE = 500 * 1024 * 1024
MAX_ZIP_ENTRIES = 10_000
CHUNK_SIZE = 1024 * 1024

DOCX_MEDIA_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.document"
)


def safe_zip_path(name: str) -> Path:
    """Rejeita caminhos absolutos, traversal e caminhos Windows perigosos."""

    if not name or "\x00" in name or "\\" in name:
        raise ValueError("O ZIP contém um caminho inválido.")

    posix_path = PurePosixPath(name)
    windows_path = PureWindowsPath(name)

    if (
        posix_path.is_absolute()
        or windows_path.is_absolute()
        or windows_path.drive
        or ".." in posix_path.parts
        or ".." in windows_path.parts
    ):
        raise ValueError("O ZIP contém um caminho inseguro.")

    if any(part in {"", "."} for part in posix_path.parts):
        raise ValueError("O ZIP contém um caminho inválido.")

    return Path(*posix_path.parts)


def extract_zip_safely(zip_path: Path, destination: Path) -> int:
    """Extrai arquivos após validar caminhos, tamanho e tipos de entrada."""

    total_uncompressed = 0

    with zipfile.ZipFile(zip_path, "r") as archive:
        entries = archive.infolist()

        if not entries:
            raise ValueError("O arquivo ZIP está vazio.")

        if len(entries) > MAX_ZIP_ENTRIES:
            raise ValueError("O ZIP contém arquivos demais.")

        validated_entries = []

        for entry in entries:
            relative_path = safe_zip_path(entry.filename)

            unix_mode = entry.external_attr >> 16
            file_type = stat.S_IFMT(unix_mode)

            if file_type not in {0, stat.S_IFREG, stat.S_IFDIR}:
                raise ValueError(
                    "O ZIP contém um tipo de arquivo não permitido."
                )

            total_uncompressed += entry.file_size

            if total_uncompressed > MAX_EXTRACTED_SIZE:
                raise ValueError(
                    "O conteúdo extraído excederia o limite permitido."
                )

            if entry.file_size > 0 and entry.compress_size == 0:
                raise ValueError(
                    "O ZIP contém uma entrada com compressão inválida."
                )

            if (
                entry.compress_size > 0
                and entry.file_size / entry.compress_size > 200
            ):
                raise ValueError(
                    "O ZIP contém uma taxa de compressão excessiva."
                )

            target = destination / relative_path

            if entry.is_dir():
                target.mkdir(parents=True, exist_ok=True)
            else:
                target.parent.mkdir(parents=True, exist_ok=True)
                validated_entries.append((entry, target))

        for entry, target in validated_entries:
            with archive.open(entry, "r") as source:
                with target.open("xb") as output:
                    shutil.copyfileobj(
                        source,
                        output,
                        length=CHUNK_SIZE,
                    )

    return total_uncompressed


@app.get("/health")
def health_check():
    return {
        "status": "ok",
        "application": "E-BI DOC Generator",
        "version": "0.4.0",
    }


@app.post("/upload-project")
async def upload_project(file: UploadFile = File(...)):
    filename = Path(file.filename or "").name

    if not filename or Path(filename).suffix.lower() != ".zip":
        raise HTTPException(
            status_code=400,
            detail="Envie um arquivo ZIP contendo o projeto Power BI.",
        )

    try:
        with tempfile.TemporaryDirectory(prefix="ebi_doc_") as temp_dir:
            workspace = Path(temp_dir)
            zip_path = workspace / "upload.zip"
            extracted_dir = workspace / "project"
            extracted_dir.mkdir()

            total_size = 0

            with zip_path.open("wb") as output:
                while True:
                    chunk = await file.read(CHUNK_SIZE)

                    if not chunk:
                        break

                    total_size += len(chunk)

                    if total_size > MAX_UPLOAD_SIZE:
                        raise HTTPException(
                            status_code=413,
                            detail="O ZIP excede o limite de 200 MB.",
                        )

                    output.write(chunk)

            if total_size == 0:
                raise HTTPException(
                    status_code=400,
                    detail="O arquivo ZIP está vazio.",
                )

            try:
                extracted_size = extract_zip_safely(
                    zip_path,
                    extracted_dir,
                )
            except (
                ValueError,
                zipfile.BadZipFile,
                OSError,
                RuntimeError,
            ) as exc:
                raise HTTPException(
                    status_code=400,
                    detail=f"ZIP inválido ou inseguro: {exc}",
                ) from exc

            pbip_files = list(extracted_dir.rglob("*.pbip"))

            if len(pbip_files) != 1:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        "O ZIP deve conter exatamente um arquivo .pbip. "
                        f"Foram encontrados {len(pbip_files)}."
                    ),
                )

            pbip_file = pbip_files[0]
            project_root = pbip_file.parent

            try:
                validation = validate_pbip_file(pbip_file)
            except ValueError as exc:
                raise HTTPException(
                    status_code=400,
                    detail=f"Arquivo PBIP inválido: {exc}",
                ) from exc

            try:
                with pbip_file.open(
                    "r",
                    encoding="utf-8-sig",
                ) as source:
                    pbip_data = json.load(source)
            except (
                OSError,
                UnicodeDecodeError,
                json.JSONDecodeError,
            ) as exc:
                raise HTTPException(
                    status_code=400,
                    detail=f"Não foi possível ler o arquivo PBIP: {exc}",
                ) from exc

            artifact_paths = []

            for artifact in pbip_data["artifacts"]:
                report_path = artifact["report"]["path"]
                report_dir = (project_root / report_path).resolve()
                resolved_root = project_root.resolve()

                if not report_dir.is_relative_to(resolved_root):
                    raise HTTPException(
                        status_code=400,
                        detail="Referência de relatório fora do projeto.",
                    )

                if not report_dir.is_dir():
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            "A pasta de relatório referenciada não existe: "
                            f"{report_path}"
                        ),
                    )

                artifact_paths.append(report_path)

            try:
                project_inspection = inspect_project_directory(
                    project_root
                )
            except ValueError as exc:
                raise HTTPException(
                    status_code=400,
                    detail=f"Não foi possível inspecionar o projeto: {exc}",
                ) from exc

            semantic_model_dirs = [
                path
                for path in project_root.iterdir()
                if path.is_dir()
                and path.name.endswith(".SemanticModel")
            ]

            semantic_analysis = None
            semantic_analysis_errors = []

            if len(semantic_model_dirs) == 1:
                try:
                    semantic_analysis = extract_tmdl_model(
                        semantic_model_dirs[0]
                    )
                except (ValueError, OSError) as exc:
                    semantic_analysis_errors.append(str(exc))

            elif len(semantic_model_dirs) > 1:
                semantic_analysis_errors.append(
                    "Foi encontrada mais de uma pasta de modelo semântico; "
                    "a análise automática exige uma única pasta."
                )
            else:
                semantic_analysis_errors.append(
                    "Nenhuma pasta .SemanticModel foi encontrada."
                )

            if semantic_analysis is None:
                raise HTTPException(
                    status_code=422,
                    detail={
                        "message": (
                            "O projeto foi validado, mas não foi possível "
                            "extrair o modelo semântico para gerar o DOCX."
                        ),
                        "errors": semantic_analysis_errors,
                    },
                )

            docx_path = workspace / "documentacao_power_bi.docx"

            try:
                generate_docx(
                    model=semantic_analysis,
                    output_path=docx_path,
                    project_name=(
                        pbip_file.stem
                            if project_root.name.lower() in {"project", "projeto"}
                            else project_root.name
                        ),
                    report_dir=(
                        project_root / artifact_paths[0]
                        if artifact_paths
                        else None
                    ),
                )

                docx_content = docx_path.read_bytes()

            except (OSError, ValueError, RuntimeError) as exc:
                raise HTTPException(
                    status_code=500,
                    detail=f"Falha ao gerar o documento DOCX: {exc}",
                ) from exc

            return Response(
                content=docx_content,
                media_type=DOCX_MEDIA_TYPE,
                headers={
                    "Content-Disposition": (
                        'attachment; filename="documentacao_power_bi.docx"'
                    ),
                    "Content-Length": str(len(docx_content)),
                },
            )

    finally:
        await file.close()
