import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from backend import crud, schemas, models, database
from sqlalchemy.ext.asyncio import AsyncSession
from typing import List

app = FastAPI(title="Powerpoint Clone API")

origins = ["*"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
async def startup():
    async with database.engine.begin() as conn:
        await conn.run_sync(models.Base.metadata.create_all)

@app.post("/presentations/", response_model=schemas.PresentationRead, status_code=201)
async def create_presentation(
    presentation: schemas.PresentationCreate,
    db: AsyncSession = database.get_db(),
):
    try:
        return await crud.create_presentation(db, presentation)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/presentations/", response_model=List[schemas.PresentationRead])
async def list_presentations(
    skip: int = 0,
    limit: int = 100,
    db: AsyncSession = database.get_db(),
):
    return await crud.get_presentations(db, skip, limit)

@app.get("/presentations/{presentation_id}", response_model=schemas.PresentationRead)
async def read_presentation(
    presentation_id: int,
    db: AsyncSession = database.get_db(),
):
    presentation = await crud.get_presentation(db, presentation_id)
    if not presentation:
        raise HTTPException(status_code=404, detail="Presentation not found")
    return presentation

@app.delete("/presentations/{presentation_id}", status_code=204)
async def delete_presentation(
    presentation_id: int,
    db: AsyncSession = database.get_db(),
):
    presentation = await crud.get_presentation(db, presentation_id)
    if not presentation:
        raise HTTPException(status_code=404, detail="Presentation not found")
    await crud.delete_presentation(db, presentation_id)
    return None

if __name__ == "__main__":
    uvicorn.run("backend.main:app", host="0.0.0.0", port=8000, reload=True)