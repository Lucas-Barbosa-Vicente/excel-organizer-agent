import logging

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.config import settings
from app.routes import health, rename

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="PDF Renamer",
    version="1.0.0",
    description="Renomeie PDFs de funcionários usando uma planilha Excel como base",
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


app.include_router(health.router, tags=["Saúde"])
app.include_router(rename.router, prefix="/api", tags=["Renomear PDFs"])
