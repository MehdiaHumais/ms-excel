import asyncio
from uuid import UUID, uuid4
from typing import Dict, List, Optional

from .schemas import PresentationCreate, SlideCreate, PresentationResponse, SlideResponse

class Slide:
    def __init__(self, content: str, order: int):
        self.id: UUID = uuid4()
        self.content: str = content
        self.order: int = order

class Presentation:
    def __init__(self, title: str):
        self.id: UUID = uuid4()
        self.title: str = title
        self.slides: Dict[UUID, Slide] = {}
        self._slide_order_counter: int = 0

    def add_slide(self, slide_data: SlideCreate) -> Slide:
        slide = Slide(content=slide_data.content, order=slide_data.order)
        self.slides[slide.id] = slide
        return slide

    def get_slides_sorted(self) -> List[Slide]:
        return sorted(self.slides.values(), key=lambda s: s.order)

class PresentationStorage:
    """
    In‑memory async storage for presentations.
    This is a placeholder for a real DB (e.g., PostgreSQL) and is safe for concurrent
    access thanks to an asyncio lock.
    """

    def __init__(self):
        self._presentations: Dict[UUID, Presentation] = {}
        self._lock = asyncio.Lock()

    async def create_presentation(self, title: str) -> PresentationResponse:
        async with self._lock:
            pres = Presentation(title=title)
            self._presentations[pres.id] = pres
            return PresentationResponse(
                id=pres.id,
                title=pres.title,
                slides=[],
            )

    async def get_presentation(self, pres_id: UUID) -> Optional[PresentationResponse]:
        async with self._lock:
            pres = self._presentations.get(pres_id)
            if not pres:
                return None
            slides = [
                SlideResponse(
                    id=s.id,
                    content=s.content,
                    order=s.order,
                )
                for s in pres.get_slides_sorted()
            ]
            return PresentationResponse(
                id=pres.id,
                title=pres.title,
                slides=slides,
            )

    async def delete_presentation(self, pres_id: UUID) -> bool:
        async with self._lock:
            return self._presentations.pop(pres_id, None) is not None

    async def add_slide(self, pres_id: UUID, slide_data: SlideCreate) -> Optional[SlideResponse]:
        async with self._lock:
            pres = self._presentations.get(pres_id)
            if not pres:
                return None
            slide = pres.add_slide(slide_data)
            return SlideResponse(
                id=slide.id,
                content=slide.content,
                order=slide.order,
            )

    async def update_slide(
        self,
        pres_id: UUID,
        slide_id: UUID,
        slide_data: SlideCreate,
    ) -> Optional[SlideResponse]:
        async with self._lock:
            pres = self._presentations.get(pres_id)
            if not pres:
                return None
            slide = pres.slides.get(slide_id)
            if not slide:
                return None
            slide.content = slide_data.content
            slide.order = slide_data.order
            return SlideResponse(
                id=slide.id,
                content=slide.content,
                order=slide.order,
            )

    async def delete_slide(self, pres_id: UUID, slide_id: UUID) -> bool:
        async with self._lock:
            pres = self._presentations.get(pres_id)
            if not pres:
                return False
            return pres.slides.pop(slide_id, None) is not None

# Singleton storage instance for the app lifecycle
_storage_instance: Optional[PresentationStorage] = None

def get_storage() -> PresentationStorage:
    global _storage_instance
    if _storage_instance is None:
        _storage_instance = PresentationStorage()
    return _storage_instance