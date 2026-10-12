from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from . import models, schemas
from typing import List

async def get_presentation(db: AsyncSession, presentation_id: int) -> models.Presentation | None:
    result = await db.execute(
        select(models.Presentation).where(models.Presentation.id == presentation_id)
    )
    return result.scalar_one_or_none()

async def get_presentations(db: AsyncSession, skip: int = 0, limit: int = 100) -> List[models.Presentation]:
    result = await db.execute(
        select(models.Presentation).offset(skip).limit(limit)
    )
    return result.scalars().all()

async def create_presentation(db: AsyncSession, presentation: schemas.PresentationCreate) -> models.Presentation:
    db_presentation = models.Presentation(
        title=presentation.title,
        description=presentation.description,
    )
    db.add(db_presentation)
    await db.flush()  # assign id

    if presentation.slides:
        for slide in presentation.slides:
            db_slide = models.Slide(
                title=slide.title,
                content=slide.content,
                presentation_id=db_presentation.id,
            )
            db.add(db_slide)

    await db.commit()
    await db.refresh(db_presentation)
    return db_presentation

async def delete_presentation(db: AsyncSession, presentation_id: int) -> None:
    result = await db.execute(
        select(models.Presentation).where(models.Presentation.id == presentation_id)
    )
    presentation = result.scalar_one_or_none()
    if presentation:
        await db.delete(presentation)
        await db.commit()