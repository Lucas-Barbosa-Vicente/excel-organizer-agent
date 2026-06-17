import logging
import os

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.config import settings
from app.database import create_tables
from app.routes import health, organize, profiles

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="Excel Organizer Agent",
    version="1.0.0",
    description="Organize planilhas Excel com inteligência artificial",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.middleware("http")
async def log_requests(request: Request, call_next):
    logger.info("%s %s", request.method, request.url.path)
    response = await call_next(request)
    logger.info("Status: %s", response.status_code)
    return response


@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.error("Erro não tratado: %s", exc, exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"status": "error", "message": str(exc)},
    )


@app.on_event("startup")
async def startup():
    os.makedirs("storage/uploads", exist_ok=True)
    os.makedirs("storage/outputs", exist_ok=True)
    create_tables()
    if not settings.anthropic_api_key:
        logger.warning(
            "ANTHROPIC_API_KEY não configurada — instruções em linguagem natural estarão desativadas."
        )
    else:
        logger.info("ANTHROPIC_API_KEY detectada.")


app.include_router(health.router, tags=["Saúde"])
app.include_router(organize.router, prefix="/api", tags=["Organizar"])
app.include_router(profiles.router, prefix="/api", tags=["Perfis"])
