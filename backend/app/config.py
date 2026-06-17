from pydantic_settings import BaseSettings
from pydantic import field_validator
from typing import List
import json


class Settings(BaseSettings):
    anthropic_api_key: str = ""
    database_url: str = "sqlite:///./storage/organizer.db"
    max_file_size_mb: int = 50
    cors_origins: List[str] = ["http://localhost:3000", "https://localhost:3000"]
    environment: str = "development"

    @field_validator("cors_origins", mode="before")
    @classmethod
    def parse_cors_origins(cls, v):
        if isinstance(v, str):
            return json.loads(v)
        return v

    @field_validator("anthropic_api_key")
    @classmethod
    def validate_api_key(cls, v):
        if not v or v == "sua_chave_aqui":
            return ""
        if not v.startswith("sk-"):
            raise ValueError("ANTHROPIC_API_KEY inválida — deve começar com 'sk-'")
        return v

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"


settings = Settings()
