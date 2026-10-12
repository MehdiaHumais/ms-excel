import os
from typing import List

from fastapi import APIRouter, Depends, File, UploadFile, HTTPException, status, Query
from fastapi.responses import FileResponse

from ..database import async_session
from .. import crud, schemas

router = APIRouter()

async def get_db():
    async with async_session() as session:
        yield session

@router.get("/spreadsheets", response_model=List[schemas.SpreadsheetRead])
async def list_spreadsheets(skip: int = Query(0, ge=0), limit: int = Query(100, ge=1, le=1000), db=Depends(get_db)):
    try:
        spreadsheets = await crud.list_spreadsheets(db)
        return spreadsheets[skip : skip + limit]
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/spreadsheets", response_model=schemas.SpreadsheetRead, status_code=status.HTTP_201_CREATED)
async def upload_spreadsheet(file: UploadFile = File(...), db=Depends(get_db)):
    if not file.filename:
        raise HTTPException(status_code=400, detail="File must have a filename")
    if not file.filename.lower().endswith(('.xlsx', '.xlsm', '.csv')):
        raise HTTPException(status_code=400, detail="Unsupported file type")
    temp_path = os.path.join("/tmp", file.filename)
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        spreadsheet = await crud.create_spreadsheet(db, temp_path, file.filename)
        return spreadsheet
    except FileExistsError:
        raise HTTPException(status_code=409, detail="File already exists")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@router.get("/spreadsheets/{spreadsheet_id}", response_class=FileResponse)
async def download_spreadsheet(spreadsheet_id: int, db=Depends(get_db)):
    spreadsheet = await crud.get_spreadsheet_by_id(db, spreadsheet_id)
    if not spreadsheet:
        raise HTTPException(status_code=404, detail="Spreadsheet not found")
    if not os.path.exists(spreadsheet.path):
        raise HTTPException(status_code=404, detail="File missing on server")
    return FileResponse(spreadsheet.path, filename=spreadsheet.filename, media_type="application/octet-stream")

@router.delete("/spreadsheets/{spreadsheet_id}", status_code=status.HTTP_204_NO_CONTENT)
async def delete_spreadsheet(spreadsheet_id: int, db=Depends(get_db)):
    try:
        await crud.delete_spreadsheet(db, spreadsheet_id)
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail="Spreadsheet not found")
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))