import asyncio
import io
import uuid
from typing import List

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.schemas.rename import RenameResponse
from app.services.pdf_renamer import build_renamed_zip, load_employees

router = APIRouter()

_zip_map: dict[str, bytes] = {}


@router.post("/rename-pdfs", response_model=RenameResponse)
async def rename_pdfs(
    excel_file: UploadFile = File(...),
    pdf_files: List[UploadFile] = File(...),
    name_column: str = Form(...),
    id_column: str = Form(...),
    output_pattern: str = Form("{id} - {name}"),
):
    excel_bytes = await excel_file.read()
    try:
        employees = await asyncio.to_thread(
            load_employees, excel_bytes, name_column, id_column
        )
    except KeyError as e:
        raise HTTPException(status_code=422, detail=f"Coluna não encontrada na planilha: {e}")
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Erro ao ler planilha: {e}")

    pdf_data: list[tuple[str, bytes]] = []
    for pdf in pdf_files:
        content = await pdf.read()
        pdf_data.append((pdf.filename or "arquivo.pdf", content))

    zip_bytes, matched, unmatched, unmatched_files = await asyncio.to_thread(
        build_renamed_zip, pdf_data, employees, output_pattern
    )

    if matched == 0:
        raise HTTPException(
            status_code=400,
            detail=(
                "Nenhum PDF pôde ser associado a um funcionário. "
                "Verifique se os nomes dos arquivos contêm a matrícula ou o nome dos funcionários."
            ),
        )

    token = str(uuid.uuid4())
    _zip_map[token] = zip_bytes

    return RenameResponse(
        status="success",
        download_token=token,
        matched=matched,
        unmatched=unmatched,
        unmatched_files=unmatched_files,
        total=matched + unmatched,
    )


@router.get("/download-zip/{token}")
async def download_zip(token: str):
    zip_bytes = _zip_map.pop(token, None)
    if zip_bytes is None:
        raise HTTPException(
            status_code=404, detail="Token inválido ou ZIP já baixado."
        )
    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=pdfs_renomeados.zip"},
    )
