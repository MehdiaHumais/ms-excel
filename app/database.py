# app/database.py
from sqlmodel import SQLModel, create_engine
from sqlalchemy.ext.asyncio import AsyncEngine, create_async_engine
from .config import settings

# Async engine for SQLite (aiosqlite)
engine: AsyncEngine = create_async_engine(settings.DATABASE_URL, echo=False, future=True)

async def init_db() -> None:
    """Create database tables."""
    async with engine.begin() as conn:
        await conn.run_sync(SQLModel.metadata.create_all)