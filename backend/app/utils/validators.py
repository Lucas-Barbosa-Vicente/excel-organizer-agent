import os
from fastapi import HTTPException

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".ods"}


def validate_file(filename: str, size_bytes: int, max_mb: int) -> None:
    ext = os.path.splitext(filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=400,
            detail=f"Formato inválido '{ext}'. Formatos aceitos: {', '.join(ALLOWED_EXTENSIONS)}",
        )
    max_bytes = max_mb * 1024 * 1024
    if size_bytes > max_bytes:
        raise HTTPException(
            status_code=413,
            detail=f"Arquivo excede o limite de {max_mb}MB.",
        )


def validate_columns(requested: list[str], available: list[str]) -> None:
    missing = [c for c in requested if c not in available]
    if missing:
        raise HTTPException(
            status_code=400,
            detail=f"Colunas não encontradas: {missing}. Disponíveis: {available}",
        )
