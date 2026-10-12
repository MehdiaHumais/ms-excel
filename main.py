# main.py
import uvicorn
from fastapi import FastAPI, HTTPException, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from pydantic import BaseModel, Field, validator
from typing import List, Optional
import uuid

app = FastAPI(
    title="PowerPoint Clone Backend",
    description="API for managing presentations, slides, and assets.",
    version="1.0.0",
)

# CORS configuration - allow all origins for demo; restrict in production
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # TODO: replace with specific origins in prod
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In‑memory storage (replace with async DB in production)
presentations: dict[str, dict] = {}

# -------------------------
# Pydantic Schemas
# -------------------------

class SlideContent(BaseModel):
    """Content of a slide. For simplicity we only support plain text."""
    text: str = Field(..., min_length=1, description="Plain text content of the slide")

    @validator("text")
    def strip_text(cls, v: str) -> str:
        return v.strip()

class SlideCreate(BaseModel):
    """Data required to create a new slide."""
    title: str = Field(..., min_length=1, description="Slide title")
    content: SlideContent

    @validator("title")
    def strip_title(cls, v: str) -> str:
        return v.strip()

class Slide(SlideCreate):
    """Slide model returned by the API, includes an ID."""
    id: str = Field(..., description="Unique identifier for the slide")

class PresentationCreate(BaseModel):
    """Data required to create a new presentation."""
    name: str = Field(..., min_length=1, description="Presentation name")
    description: Optional[str] = Field(None, description="Optional description")
    slides: Optional[List[SlideCreate]] = Field(default_factory=list)

    @validator("name")
    def strip_name(cls, v: str) -> str:
        return v.strip()

class Presentation(PresentationCreate):
    """Presentation model returned by the API, includes an ID and slide IDs."""
    id: str = Field(..., description="Unique identifier for the presentation")
    slides: List[Slide] = Field(default_factory=list)

# -------------------------
# Utility Functions
# -------------------------

def generate_id() -> str:
    """Generate a short UUID4 string."""
    return uuid.uuid4().hex[:12]

# -------------------------
# API Endpoints
# -------------------------

@app.exception_handler(HTTPException)
async def http_exception_handler(request: Request, exc: HTTPException):
    """Return JSON response for HTTPException."""
    return JSONResponse(
        status_code=exc.status_code,
        content={"detail": exc.detail},
    )

@app.post("/presentations", response_model=Presentation, status_code=status.HTTP_201_CREATED)
async def create_presentation(payload: PresentationCreate):
    """Create a new presentation."""
    try:
        pres_id = generate_id()
        slides: List[Slide] = []
        for slide_data in payload.slides:
            slide = Slide(id=generate_id(), **slide_data.dict())
            slides.append(slide)

        presentation = Presentation(
            id=pres_id,
            name=payload.name,
            description=payload.description,
            slides=slides,
        )
        presentations[pres_id] = presentation.dict()
        return presentation
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/presentations", response_model=List[Presentation])
async def list_presentations():
    """List all presentations."""
    try:
        return [Presentation(**data) for data in presentations.values()]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/presentations/{presentation_id}", response_model=Presentation)
async def get_presentation(presentation_id: str):
    """Retrieve a single presentation by ID."""
    try:
        data = presentations.get(presentation_id)
        if not data:
            raise HTTPException(status_code=404, detail="Presentation not found")
        return Presentation(**data)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/presentations/{presentation_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_presentation(presentation_id: str):
    """Delete a presentation."""
    try:
        if presentation_id not in presentations:
            raise HTTPException(status_code=404, detail="Presentation not found")
        del presentations[presentation_id]
        return JSONResponse(status_code=status.HTTP_204_NO_CONTENT, content=None)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post(
    "/presentations/{presentation_id}/slides",
    response_model=Slide,
    status_code=status.HTTP_201_CREATED,
)
async def add_slide(presentation_id: str, payload: SlideCreate):
    """Add a new slide to an existing presentation."""
    try:
        pres_data = presentations.get(presentation_id)
        if not pres_data:
            raise HTTPException(status_code=404, detail="Presentation not found")
        slide = Slide(id=generate_id(), **payload.dict())
        pres_data["slides"].append(slide.dict())
        presentations[presentation_id] = pres_data
        return slide
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.delete(
    "/presentations/{presentation_id}/slides/{slide_id}",
    status_code=status.HTTP_204_NO_CONTENT,
)
async def delete_slide(presentation_id: str, slide_id: str):
    """Delete a slide from a presentation."""
    try:
        pres_data = presentations.get(presentation_id)
        if not pres_data:
            raise HTTPException(status_code=404, detail="Presentation not found")
        slides = pres_data["slides"]
        filtered = [s for s in slides if s["id"] != slide_id]
        if len(filtered) == len(slides):
            raise HTTPException(status_code=404, detail="Slide not found")
        pres_data["slides"] = filtered
        presentations[presentation_id] = pres_data
        return JSONResponse(status_code=status.HTTP_204_NO_CONTENT, content=None)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# -------------------------
# Run the application
# -------------------------

if __name__ == "__main__":
    # Use reload only in development
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)