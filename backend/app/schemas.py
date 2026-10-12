from datetime import datetime
from pydantic import BaseModel, Field, validator
from typing import Optional

class SpreadsheetCreate(BaseModel):
    filename: str = Field(..., min_length=1, max_length=255)

    @validator("filename")
    def validate_filename(cls, v):
        if "/" in v or "\\" in v:
            raise ValueError("Filename must not contain path separators")
        return v

class SpreadsheetRead(BaseModel):
    id: int
    filename: str
    uploaded_at: datetime

    class Config:
        orm_mode = True