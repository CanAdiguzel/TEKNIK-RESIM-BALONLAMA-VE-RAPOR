from __future__ import annotations

import io
import re
from pathlib import Path

import fitz
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

from .models import Annotation
from .ollama_utils import DetectionBox, label_for_bbox, safe_bbox


_OCR: RapidOCR | None = None


def get_ocr() -> RapidOCR:
    global _OCR
    if _OCR is None:
        _OCR = RapidOCR()
    return _OCR


def parse_values(text: str) -> tuple[float | None, float | None, float | None]:
    normalized = text.upper().replace(",", ".").replace("−", "-").replace("O", "0")
    radius_match = re.search(r"R\s*(\d+(?:\.\d+)?)", normalized)
    nominal_match = radius_match or re.search(r"(?:Ø|⌀)?\s*(\d+(?:\.\d+)?)", normalized)
    nominal = float(nominal_match.group(1)) if nominal_match else None
    plus_minus = re.search(r"±\s*(\d+(?:\.\d+)?)", normalized)
    if plus_minus:
        value = float(plus_minus.group(1))
        return nominal, value, value
    upper_match = re.search(r"\+\s*(\d+(?:\.\d+)?)", normalized)
    lower_match = re.search(r"-\s*(\d+(?:\.\d+)?)", normalized)
    upper = float(upper_match.group(1)) if upper_match else None
    lower = float(lower_match.group(1)) if lower_match else None
    return nominal, upper, lower


def candidate_kind(text: str) -> str | None:
    compact = text.upper().replace(" ", "")
    if "±" in compact:
        return "Boyutsal"
    if re.search(r"(?:R|\d)\d*(?:\.\d+)?\+\d", compact):
        return "Boyutsal"
    if re.search(r"\d(?:\.\d+)?ABC$", compact):
        return "Geometrik"
    if re.search(r"M\d+(?:X|V)\d", compact) and ("THRU" in compact or "THR" in compact):
        return "Diş / Kritik"
    if compact in {"0.05", "0,05"}:
        return "Geometrik"
    return None


def box_bounds(box: list[list[float]]) -> tuple[float, float, float, float]:
    xs = [point[0] for point in box]
    ys = [point[1] for point in box]
    return min(xs), min(ys), max(xs), max(ys)


def union_box(
    first: tuple[float, float, float, float],
    second: tuple[float, float, float, float],
) -> tuple[float, float, float, float]:
    return (
        min(first[0], second[0]),
        min(first[1], second[1]),
        max(first[2], second[2]),
        max(first[3], second[3]),
    )


def normalized_detection_box(
    bounds: tuple[float, float, float, float],
    width: int,
    height: int,
) -> DetectionBox:
    x1, y1, x2, y2 = bounds
    padding_x = max(4, (x2 - x1) * 0.05)
    padding_y = max(4, (y2 - y1) * 0.18)
    return DetectionBox(
        x1=round(max(0, x1 - padding_x) / width * 1000),
        y1=round(max(0, y1 - padding_y) / height * 1000),
        x2=round(min(width, x2 + padding_x) / width * 1000),
        y2=round(min(height, y2 + padding_y) / height * 1000),
    )


def analyze_page_ocr(
    page: fitz.Page,
    page_number: int,
    rotation: int,
    start_balloon: int,
) -> list[Annotation]:
    pixmap = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), alpha=False)
    image = Image.open(io.BytesIO(pixmap.tobytes("png"))).convert("RGB")
    if rotation:
        image = image.rotate(rotation, expand=True, fillcolor="white")

    results, _ = get_ocr()(image)
    rows = []
    for raw in results or []:
        box, text, score = raw
        if score < 0.72:
            continue
        rows.append(
            {
                "text": str(text).strip(),
                "score": float(score),
                "bounds": box_bounds(box),
            }
        )

    annotations: list[Annotation] = []
    used: set[int] = set()
    for index, row in enumerate(rows):
        if index in used:
            continue
        kind = candidate_kind(row["text"])
        if kind is None:
            continue
        bounds = row["bounds"]
        combined_text = row["text"]

        # Birbirine yakın thread ve GD&T satırlarını tek annotation içinde birleştir.
        if kind in {"Diş / Kritik", "Geometrik"}:
            x1, y1, x2, y2 = bounds
            for other_index, other in enumerate(rows):
                if other_index == index or other_index in used:
                    continue
                ox1, oy1, ox2, oy2 = other["bounds"]
                horizontal_overlap = min(x2, ox2) - max(x1, ox1)
                near_vertical = 0 <= oy1 - y2 <= 55
                if horizontal_overlap > -40 and near_vertical:
                    other_kind = candidate_kind(other["text"])
                    if other_kind == "Geometrik" or re.search(r"\d", other["text"]):
                        combined_text = f"{combined_text} | {other['text']}"
                        bounds = union_box(bounds, other["bounds"])
                        used.add(other_index)
                        break

        detection_box = normalized_detection_box(bounds, image.width, image.height)
        bbox = safe_bbox(detection_box, rotation)
        if bbox is None:
            continue
        nominal, upper, lower = parse_values(combined_text)
        annotation = Annotation(
            balloon_no=start_balloon + len(annotations),
            page=page_number,
            parameter="Yerel OCR tespiti",
            dimension_text=combined_text,
            nominal=nominal,
            upper_tol=upper,
            lower_tol=lower,
            tolerance_type=kind,
            gdnt_reference="A | B | C" if "ABC" in combined_text.upper() else "",
            note="Yerel OCR ile otomatik tespit edildi; teknik olarak doğrulayın.",
            target_bbox=bbox,
            label_position=label_for_bbox(bbox, len(annotations)),
        )
        annotations.append(annotation)
        used.add(index)

    return annotations


def analyze_pdf_ocr(
    pdf_path: Path,
    rotation: int,
) -> list[Annotation]:
    document = fitz.open(pdf_path)
    annotations: list[Annotation] = []
    try:
        for page_index, page in enumerate(document):
            annotations.extend(
                analyze_page_ocr(
                    page,
                    page_index + 1,
                    rotation,
                    len(annotations) + 1,
                )
            )
    finally:
        document.close()
    return annotations
