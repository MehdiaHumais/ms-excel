import datetime
from sqlalchemy import Column, Integer, String, DateTime
from .database import Base

class Spreadsheet(Base):
    __tablename__ = "spreadsheets"

    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, unique=True, nullable=False, index=True)
    path = Column(String, nullable=False)
    uploaded_at = Column(DateTime, default=datetime.datetime.utcnow)