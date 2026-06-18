import json
import os
import asyncio
from fastapi import APIRouter, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse

from app.config import settings
from app.schemas.organize import OrganizeParameters, OrganizeResponse
from app.services.excel_processor import ExcelProcessor
from app.services.file_handler import save_upload, generate_output_path, cleanup_old_files
from app.services.agent import interpret_instruction
from app.utils.validators import validate_file

router = APIRouter()

_token_map: dict[str, str] = {}  # token -> file path


@router.post("/organize", response_model=OrganizeResponse)
async def organize(
    file: UploadFile = File(...),
    parameters: str = Form("{}"),
):
    content = await file.read()
    validate_file(file.filename or "upload.xlsx", len(content), settings.max_file_size_mb)

    upload_path = await save_upload(content, file.filename or "upload.xlsx")

    try:
        params = OrganizeParameters(**json.loads(parameters))
    except Exception:
        params = OrganizeParameters()

    force_override = params.force_override or False

    processor = ExcelProcessor(upload_path)
    await processor.load()

    available_columns: list[str] = []
    for df in processor.dataframes.values():
        available_columns = list(df.columns)
        break

    # Check for existing customizations before applying any changes
    if not force_override:
        customizations = await asyncio.to_thread(processor.detect_customizations)
        if customizations:
            await asyncio.to_thread(os.remove, upload_path)
            return OrganizeResponse(
                status="requires_confirmation",
                message="A planilha contém formatação pré-definida. Confirme para prosseguir.",
                transformations_applied=[],
                rows_before=0,
                rows_after=0,
                sheets_created=[],
                requires_confirmation=True,
                existing_customizations=customizations,
            )

    if params.natural_language_instruction:
        if not settings.anthropic_api_key:
            raise HTTPException(
                status_code=500,
                detail="ANTHROPIC_API_KEY não configurada. Adicione ao arquivo .env.",
            )
        params = await interpret_instruction(
            params.natural_language_instruction, available_columns
        )
        params.force_override = force_override

    stats = await asyncio.wait_for(processor.apply_parameters(params), timeout=300)

    ext = os.path.splitext(upload_path)[1]
    token, output_path = generate_output_path(ext)
    os.makedirs("storage/outputs", exist_ok=True)
    await processor.save(output_path)

    _token_map[token] = output_path

    await asyncio.to_thread(os.remove, upload_path)
    asyncio.create_task(cleanup_old_files())

    return OrganizeResponse(
        status="success",
        message="Planilha organizada com sucesso.",
        transformations_applied=processor.transformations,
        rows_before=stats["rows_before"],
        rows_after=stats["rows_after"],
        sheets_created=stats["sheets_created"],
        download_token=token,
    )


@router.get("/download/{token}")
async def download(token: str):
    file_path = _token_map.get(token)
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Token inválido ou arquivo expirado.")

    async def remove_after():
        await asyncio.sleep(1)
        try:
            os.remove(file_path)
            _token_map.pop(token, None)
        except FileNotFoundError:
            pass

    asyncio.create_task(remove_after())
    return FileResponse(
        path=file_path,
        filename="organizado.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
