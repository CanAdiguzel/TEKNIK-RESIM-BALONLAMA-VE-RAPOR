from __future__ import annotations

from uuid import uuid4

from pydantic import BaseModel, Field, model_validator


class BoundingBox(BaseModel):
    """PDF sayfasına göre 0-1 aralığında normalize edilmiş hedef kutusu."""

    x: float = Field(ge=0, le=1)
    y: float = Field(ge=0, le=1)
    width: float = Field(gt=0, le=1)
    height: float = Field(gt=0, le=1)

    @model_validator(mode="after")
    def stays_inside_page(self) -> "BoundingBox":
        if self.x + self.width > 1.000001 or self.y + self.height > 1.000001:
            raise ValueError("Hedef kutusu sayfa sınırları içinde olmalıdır.")
        return self


class Point(BaseModel):
    """PDF sayfasına göre 0-1 aralığında normalize edilmiş nokta."""

    x: float = Field(ge=0, le=1)
    y: float = Field(ge=0, le=1)


class Annotation(BaseModel):
    id: str = Field(default_factory=lambda: str(uuid4()))
    balloon_no: int = Field(ge=1)
    page: int = Field(ge=1)
    parameter: str = ""
    dimension_text: str = ""
    nominal: float | None = None
    upper_tol: float | None = None
    lower_tol: float | None = None
    max_limit: float | None = None
    min_limit: float | None = None
    tolerance_type: str = "Boyutsal"
    gdnt_reference: str = ""
    note: str = ""
    target_bbox: BoundingBox
    label_position: Point

    @model_validator(mode="after")
    def calculate_limits(self) -> "Annotation":
        if self.nominal is not None:
            if self.upper_tol is not None:
                self.max_limit = self.nominal + self.upper_tol
            if self.lower_tol is not None:
                self.min_limit = self.nominal - abs(self.lower_tol)
        return self


class AnnotationSet(BaseModel):
    annotations: list[Annotation] = Field(default_factory=list)

    @model_validator(mode="after")
    def unique_balloon_numbers(self) -> "AnnotationSet":
        numbers = [item.balloon_no for item in self.annotations]
        if len(numbers) != len(set(numbers)):
            raise ValueError("Balon numaraları benzersiz olmalıdır.")
        return self


class DocumentResponse(BaseModel):
    document_id: str
    filename: str
    page_count: int
    pdf_url: str
    annotations: list[Annotation]


class AnalyzeRequest(BaseModel):
    model: str = "qwen2.5vl:3b"
    dpi: int = Field(default=110, ge=90, le=200)
    rotation: int = 270

    @model_validator(mode="after")
    def valid_rotation(self) -> "AnalyzeRequest":
        if self.rotation not in (0, 90, 180, 270):
            raise ValueError("Döndürme 0, 90, 180 veya 270 olmalıdır.")
        return self


class AnalyzeResponse(BaseModel):
    status: str = "completed"
    model: str
    page_count: int
    annotations: list[Annotation]
    warnings: list[str] = Field(default_factory=list)
