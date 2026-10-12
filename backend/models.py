from sqlalchemy import Column, Integer, String, ForeignKey, Text
from sqlalchemy.orm import relationship
from sqlalchemy.ext.declarative import declarative_base

Base = declarative_base()

class Presentation(Base):
    __tablename__ = "presentations"

    id = Column(Integer, primary_key=True, index=True)
    title = Column(String, nullable=False)
    description = Column(Text, nullable=True)

    slides = relationship(
        "Slide",
        back_populates="presentation",
        cascade="all, delete-orphan",
        lazy="selectin"
    )

class Slide(Base):
    __tablename__ = "slides"

    id = Column(Integer, primary_key=True, index=True)
    presentation_id = Column(Integer, ForeignKey("presentations.id"), nullable=False)
    title = Column(String, nullable=False)
    content = Column(Text, nullable=True)

    presentation = relationship("Presentation", back_populates="slides")