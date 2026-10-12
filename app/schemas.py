from pydantic import BaseModel, Field, validator
from uuid import UUID
from typing import List, Optional

class SlideBase(BaseModel):
    content: str = Field(..., description="Plain text content of the slide.")
    order: int = Field(..., ge=0, description="Position of the slide in the deck.")

class SlideCreate(SlideBase):
    @validator("content")
    def content_not_empty(cls, v):
        if not v.strip():
            raise ValueError("Slide content must not be empty.")
        return v

class SlideResponse(SlideBase):
    id: UUID = Field(..., description="Unique identifier for the slide.")

    class Config:
        orm_mode = True

class PresentationBase(BaseModel):
    title: str = Field(..., description="Title of the presentation.")

class PresentationCreate(PresentationBase):
    @validator("title")
    def title_not_empty(cls, v):
        if not v.strip():
            raise ValueError("Title must not be empty.")
        return v

class PresentationResponse(PresentationBase):
    id: UUID = Field(..., description="Unique identifier for the presentation.")
    slides: List[SlideResponse] = Field(default_factory=list)

    class Config:
        orm_mode = True