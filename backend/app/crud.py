import os
import shutil
from typing import List, Optional

from sqlalchemy import select, delete
from sqlalchemy.exc import IntegrityError
from sqlalchemy.ext.asyncio import AsyncSession

from . import models, schemas

UPLOAD_DIR = os.getenv("UPLOAD_DIR", "./uploaded_spreadsheets")
os.makedirs(UPLOAD_DIR, exist_ok=True)

async def get_spreadsheet_by_id(db: AsyncSession, spreadsheet_id: int) -> Optional[models.Spreadsheet]:
    result = await db.execute(select(models.Spreadsheet).where(models.Spreadsheet.id == spreadsheet_id))
    return result.scalars().first()

async def get_spreadsheet_by_name(db: AsyncSession, filename: str) -> Optional[models.Spreadsheet]:
    result = await db.execute(select(models.Spreadsheet).where(models.Spreadsheet.filename == filename))
    return result.scalars().first()

async def list_spreadsheets(db: AsyncSession) -> List[models.Spreadsheet]:
    result = await db.execute(select(models.Spreadsheet).order_by(models.Spreadsheet.uploaded_at.desc()))
    return result.scalars().all()

async def create_spreadsheet(db: AsyncSession, file_path: str, filename: str) -> models.Spreadsheet:
    dest_path = os.path.join(UPLOAD_DIR, filename)
    if os.path.exists(dest_path):
        raise FileExistsError(f"File {filename} already exists")
    shutil.copyfile(file_path, dest_path)

    new_spreadsheet = models.Spreadsheet(filename=filename, path=dest_path)
    db.add(new_spreadsheet)
    try:
        await db.commit()
        await db.refresh(new_spreadsheet)
    except IntegrityError:
        await db.rollback()
        os.remove(dest_path)
        raise
    return new_spreadsheet

async def delete_spreadsheet(db: AsyncSession, spreadsheet_id: int) -> None:
    spreadsheet = await get_spreadsheet_by_id(db, spreadsheet_id)
    if not spreadsheet:
        raise FileNotFoundError("Spreadsheet not found")
    if os.path.exists(spreadsheet.path):
        os.remove(spreadsheet.path)
    await db.execute(delete(models.Spreadsheet).where(models.Spreadsheet.id == spreadsheet_id))
    await db.commit()