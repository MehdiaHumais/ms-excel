# app/models.py
from datetime import datetime
from typing import Optional
from sqlmodel import SQLModel, Field

class Presentation(SQLModel, table=True):
    """Database model for a stored PowerPoint presentation."""

    id: Optional[int] = Field(default=None, primary_key=True)
    filename: str = Field(index=True, description="Original filename uploaded by the user")
    stored_path: str = Field(description="Absolute path where the file is stored on disk")
    upload_time: datetime = Field(default_factory=datetime.utcnow, description="Timestamp of upload")