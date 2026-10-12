from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.ext.asyncio import AsyncSession
from typing import List
from .. import crud, schemas, database

router = APIRouter(prefix="/presentations", tags=["presentations"])

@router.post("/", response_model=schemas.PresentationRead, status_code=status.HTTP_201_CREATED)
async def create_presentation(
    presentation: schemas.PresentationCreate,
    db: AsyncSession = Depends(database.get_db),
):
    try:
        return await crud.create_presentation(db, presentation)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.get("/", response_model=List[schemas.PresentationRead])
async def list_presentations(
    skip: int = 0,
    limit: int = 100,
    db: AsyncSession = Depends(database.get_db),
):
    return await crud.get_presentations(db, skip, limit)

@router.get("/{presentation_id}", response_model=schemas.PresentationRead)
async def read_presentation(
    presentation_id: int,
    db: AsyncSession = Depends(database.get_db),
):
    presentation = await crud.get_presentation(db, presentation_id)
    if not presentation:
        raise HTTPException(status_code=404, detail="Presentation not found")
    return presentation

@router.delete("/{presentation_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_presentation(
    presentation_id: int,
    db: AsyncSession = Depends(database.get_db),
):
    presentation = await crud.get_presentation(db, presentation_id)
    if not presentation:
        raise HTTPException(status_code=404, detail="Presentation not found")
    await crud.delete_presentation(db, presentation_id)
    return None