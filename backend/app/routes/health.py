from fastapi import APIRouter
from app.config import settings

router = APIRouter()


@router.get("/health")
async def health():
    return {
        "status": "ok",
        "version": "1.0.0",
        "environment": settings.environment,
    }
