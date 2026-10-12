# app/config.py
from pathlib import Path
from pydantic import BaseSettings, Field

class Settings(BaseSettings):
    """Application configuration."""

    DATABASE_URL: str = Field(default="sqlite+aiosqlite:///./app.db")
    STORAGE_DIR: Path = Field(default=Path("./storage"))
    ALLOWED_ORIGINS: list[str] = Field(default=["*"])

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

settings = Settings()