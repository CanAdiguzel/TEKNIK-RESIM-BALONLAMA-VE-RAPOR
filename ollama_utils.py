from __future__ import annotations

import base64
import json
import re
from difflib import SequenceMatcher
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

import fitz
from pydantic import BaseModel, Field
from PIL import Image
import io

from .models import Annotation, BoundingBox, Point


OLLAMA_CHAT_URL = "http://127.0.0.1:11434/api/chat"


class DetectionBox(BaseModel):
    x1: int = Field(ge=0, le=1000)
    y1: int = Field(ge=0, le=1000)
    x2: int = Field(ge=0, le=1000)
    y2: int = Field(ge=0, le=1000)


class DetectedTolerance(BaseModel):
    parameter: str = "Toleranslı ölçü"
    dimension_text: str
    nominal: float | None = None
    upper_tol: float | None = None
    lower_tol: float | None = None
    tolerance_type: str = "Boyutsal"
    gdnt_reference: str = ""
    note: str = ""
    target_bbox: DetectionBox


class PageDetections(BaseModel):
    tolerances: list[DetectedTolerance]


def analysis_prompt(page_number: int) -> str:
    schema = json.dumps(PageDetections.model_json_schema(), ensure_ascii=False)
    return f"""
Bu görsel bir makine teknik resminin {page_number}. sayfasıdır.

Görev:
- Açıkça görülen toleranslı boyutları, limit ölçülerini, GD&T çerçevelerini ve kritik
  üretim parametrelerini tespit et.
- Görsel 90 derece dönük olabilir; yazıları doğru yönde okuyarak analiz et.
- Toleransı olmayan sıradan ölçüleri ekleme.
- Okunamayan değeri tahmin etme.
- dimension_text görseldeki metni mümkün olduğunca aynen korusun.
- ±0.02 için upper_tol=0.02 ve lower_tol=0.02 kullan.
- +0/-0.1 için upper_tol=0 ve lower_tol=0.1 kullan.
- target_bbox yalnızca ölçü/tolerans yazısını sıkıca çevrelesin.
- target_bbox koordinatlarını görselin sol üstü (0,0), sağ altı (1000,1000) olacak şekilde ver.
- x1 < x2 ve y1 < y2 olmalı.
- Hiç güvenilir tolerans bulunamazsa tolerances boş dizi olsun.

Yalnızca verilen JSON şemasına uygun veri döndür:
{schema}
""".strip()


def render_page_image(page: fitz.Page, dpi: int, rotation: int) -> bytes:
    scale = dpi / 72
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    image = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
    if rotation:
        image = image.rotate(rotation, expand=True, fillcolor="white")
    max_side = 1500
    if max(image.size) > max_side:
        ratio = max_side / max(image.size)
        image = image.resize(
            (round(image.width * ratio), round(image.height * ratio)),
            Image.Resampling.LANCZOS,
        )
    output = io.BytesIO()
    image.save(output, format="JPEG", quality=88, optimize=True)
    return output.getvalue()


def call_ollama(model: str, image_bytes: bytes, page_number: int) -> PageDetections:
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": analysis_prompt(page_number),
                "images": [base64.b64encode(image_bytes).decode("ascii")],
            }
        ],
        "stream": False,
        "format": PageDetections.model_json_schema(),
        "options": {"temperature": 0, "num_ctx": 4096, "num_predict": 2200},
        "keep_alive": "10m",
    }
    request = Request(
        OLLAMA_CHAT_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urlopen(request, timeout=360) as response:
            body = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Ollama API hatası ({exc.code}): {detail}") from exc
    except URLError as exc:
        raise RuntimeError(
            "Ollama servisine bağlanılamadı. Ollama uygulamasının açık olduğundan emin olun."
        ) from exc

    content = body.get("message", {}).get("content", "")
    if not content:
        raise RuntimeError("Ollama boş analiz yanıtı döndürdü.")
    return PageDetections.model_validate_json(content)


def safe_bbox(box: DetectionBox, rotation: int = 0) -> BoundingBox | None:
    x1, x2 = sorted((box.x1 / 1000, box.x2 / 1000))
    y1, y2 = sorted((box.y1 / 1000, box.y2 / 1000))
    if rotation == 90:
        x1, y1, x2, y2 = 1 - y2, x1, 1 - y1, x2
    elif rotation == 180:
        x1, y1, x2, y2 = 1 - x2, 1 - y2, 1 - x1, 1 - y1
    elif rotation == 270:
        x1, y1, x2, y2 = y1, 1 - x2, y2, 1 - x1
    width = x2 - x1
    height = y2 - y1
    if width < 0.003 or height < 0.003:
        return None
    return BoundingBox(
        x=max(0, min(x1, 0.997)),
        y=max(0, min(y1, 0.997)),
        width=min(width, 1 - x1),
        height=min(height, 1 - y1),
    )


def label_for_bbox(box: BoundingBox, index_on_page: int) -> Point:
    center_x = box.x + box.width / 2
    x = box.x + box.width + 0.12 if center_x < 0.68 else box.x - 0.12
    stagger = ((index_on_page % 5) - 2) * 0.035
    y = box.y - 0.055 + stagger
    return Point(x=max(0.035, min(x, 0.965)), y=max(0.035, min(y, 0.965)))


def normalized_text(value: str) -> str:
    value = value.casefold().replace("±", "+/-").replace("−", "-").replace(",", ".")
    return re.sub(r"\s+", "", value)


def find_text_layer_bbox(page: fitz.Page, target_text: str) -> BoundingBox | None:
    """Model metnini PDF metin katmanıyla eşleştirip kesin koordinatı bulur."""
    target = normalized_text(target_text)
    if not target:
        return None
    best: tuple[float, tuple[float, float, float, float]] | None = None
    page_dict = page.get_text("dict")

    for block in page_dict.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            text = "".join(str(span.get("text", "")) for span in spans).strip()
            candidate = normalized_text(text)
            if not candidate:
                continue
            score = SequenceMatcher(None, target, candidate).ratio()
            if target in candidate or candidate in target:
                score = max(score, min(len(target), len(candidate)) / max(len(target), len(candidate)))
            if best is None or score > best[0]:
                bbox = (
                    min(span["bbox"][0] for span in spans),
                    min(span["bbox"][1] for span in spans),
                    max(span["bbox"][2] for span in spans),
                    max(span["bbox"][3] for span in spans),
                )
                best = (score, bbox)

    if best is None or best[0] < 0.58:
        return None
    x1, y1, x2, y2 = best[1]
    padding_x = max(2, (x2 - x1) * 0.04)
    padding_y = max(2, (y2 - y1) * 0.12)
    x1 = max(0, x1 - padding_x)
    y1 = max(0, y1 - padding_y)
    x2 = min(page.rect.width, x2 + padding_x)
    y2 = min(page.rect.height, y2 + padding_y)
    return BoundingBox(
        x=x1 / page.rect.width,
        y=y1 / page.rect.height,
        width=(x2 - x1) / page.rect.width,
        height=(y2 - y1) / page.rect.height,
    )


def analyze_pdf(
    pdf_path: Path,
    model: str,
    dpi: int,
    rotation: int = 0,
) -> tuple[list[Annotation], list[str]]:
    document = fitz.open(pdf_path)
    annotations: list[Annotation] = []
    warnings: list[str] = []

    try:
        for page_index, page in enumerate(document):
            try:
                detections = call_ollama(
                    model,
                    render_page_image(page, dpi, rotation),
                    page_index + 1,
                )
            except Exception as exc:
                warnings.append(f"Sayfa {page_index + 1}: {exc}")
                continue

            for item_index, item in enumerate(detections.tolerances):
                bbox = find_text_layer_bbox(page, item.dimension_text) or safe_bbox(
                    item.target_bbox,
                    rotation,
                )
                if bbox is None:
                    warnings.append(
                        f"Sayfa {page_index + 1}: '{item.dimension_text}' için geçersiz kutu atlandı."
                    )
                    continue
                annotations.append(
                    Annotation(
                        balloon_no=len(annotations) + 1,
                        page=page_index + 1,
                        parameter=item.parameter,
                        dimension_text=item.dimension_text,
                        nominal=item.nominal,
                        upper_tol=item.upper_tol,
                        lower_tol=abs(item.lower_tol) if item.lower_tol is not None else None,
                        tolerance_type=item.tolerance_type,
                        gdnt_reference=item.gdnt_reference,
                        note=item.note,
                        target_bbox=bbox,
                        label_position=label_for_bbox(bbox, item_index),
                    )
                )
    finally:
        document.close()

    return annotations, warnings
