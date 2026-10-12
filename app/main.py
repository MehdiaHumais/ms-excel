from fastapi import FastAPI, HTTPException, status, Depends
from fastapi.middleware.cors import CORSMiddleware
from uuid import UUID, uuid4
from typing import List

from .schemas import (
    PresentationCreate,
    PresentationResponse,
    SlideCreate,
    SlideResponse,
)
from .storage import PresentationStorage, get_storage

app = FastAPI(
    title="PowerPoint Clone API",
    description="A minimal backend for managing presentations and slides.",
    version="1.0.0",
)

# CORS configuration - allow all origins for simplicity (adjust in production)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Dependency injection for storage
def get_presentation_storage() -> PresentationStorage:
    return get_storage()

@app.post(
    "/presentations",
    response_model=PresentationResponse,
    status_code=status.HTTP_201_CREATED,
)
async def create_presentation(
    payload: PresentationCreate,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Create a new presentation.
    """
    try:
        if not payload.title.strip():
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Title must not be empty.",
            )
        presentation = await storage.create_presentation(payload.title)
        return PresentationResponse.from_orm(presentation)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )

@app.get(
    "/presentations/{presentation_id}",
    response_model=PresentationResponse,
)
async def get_presentation(
    presentation_id: UUID,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Retrieve a presentation with its slides.
    """
    try:
        presentation = await storage.get_presentation(presentation_id)
        if not presentation:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Presentation not found.",
            )
        return PresentationResponse.from_orm(presentation)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )

@app.delete(
    "/presentations/{presentation_id}",
    status_code=status.HTTP_204_NO_CONTENT,
)
async def delete_presentation(
    presentation_id: UUID,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Delete a presentation and all its slides.
    """
    try:
        success = await storage.delete_presentation(presentation_id)
        if not success:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Presentation not found.",
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )

@app.post(
    "/presentations/{presentation_id}/slides",
    response_model=SlideResponse,
    status_code=status.HTTP_201_CREATED,
)
async def add_slide(
    presentation_id: UUID,
    payload: SlideCreate,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Add a new slide to a presentation.
    """
    try:
        if not payload.content.strip():
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Slide content must not be empty.",
            )
        slide = await storage.add_slide(presentation_id, payload)
        if not slide:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Presentation not found.",
            )
        return SlideResponse.from_orm(slide)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )

@app.put(
    "/presentations/{presentation_id}/slides/{slide_id}",
    response_model=SlideResponse,
)
async def update_slide(
    presentation_id: UUID,
    slide_id: UUID,
    payload: SlideCreate,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Update an existing slide.
    """
    try:
        if not payload.content.strip():
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Slide content must not be empty.",
            )
        slide = await storage.update_slide(presentation_id, slide_id, payload)
        if not slide:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Presentation or slide not found.",
            )
        return SlideResponse.from_orm(slide)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )

@app.delete(
    "/presentations/{presentation_id}/slides/{slide_id}",
    status_code=status.HTTP_204_NO_CONTENT,
)
async def delete_slide(
    presentation_id: UUID,
    slide_id: UUID,
    storage: PresentationStorage = Depends(get_presentation_storage),
):
    """
    Delete a slide from a presentation.
    """
    try:
        success = await storage.delete_slide(presentation_id, slide_id)
        if not success:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Presentation or slide not found.",
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e),
        )