from __future__ import annotations

import io
import math
import re
from pathlib import Path
from typing import Any

import fitz
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
from streamlit_image_coordinates import streamlit_image_coordinates


APP_TITLE = "AS9102 Teknik Resim Balonlama"
RED = (220, 20, 45)
RENDER_DPI = 120

FORM3_COLUMNS = [
    "Characteristic No / Balon No",
    "Sheet",
    "Zone",
    "View",
    "Characteristic Type",
    "Characteristic Designator",
    "Drawing Requirement",
    "Nominal",
    "Upper Tolerance",
    "Lower Tolerance",
    "Upper Limit",
    "Lower Limit",
    "Units",
    "GD&T / Datum Reference",
    "Quantity",
    "Inspection Method",
    "Measuring Equipment",
    "Designed / Qualified Tooling",
    "Result 1",
    "Result 2",
    "Result 3",
    "Result Summary",
    "Acceptance Status",
    "Nonconformance No",
    "Inspector",
    "Inspection Date",
    "Remarks",
    "_target_x",
    "_target_y",
    "_balloon_x",
    "_balloon_y",
]

VISIBLE_COLUMNS = FORM3_COLUMNS[:-4]


def init_state():
    defaults = {
        "pdf_bytes": None,
        "pdf_name": None,
        "page_images": [],
        "records": pd.DataFrame(columns=FORM3_COLUMNS),
        "last_click": None,
        "viewer_key": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


@st.cache_data(show_spinner=False)
def render_pdf(pdf_bytes: bytes, dpi: int = RENDER_DPI):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    images = []

    try:
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            images.append(pix.tobytes("png"))
    finally:
        doc.close()

    return images


def as_number(value: Any):
    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        pass

    text = str(value).strip().replace(",", ".")
    if text == "":
        return None

    try:
        return float(text)
    except Exception:
        return None


def parse_requirement(text: str):
    text = str(text).replace(",", ".").replace("−", "-").upper()

    nominal_match = re.search(r"(?:Ø|⌀|R)?\s*(\d+(?:\.\d+)?)", text)
    nominal = float(nominal_match.group(1)) if nominal_match else None

    pm = re.search(r"±\s*(\d+(?:\.\d+)?)", text)
    if pm:
        tol = float(pm.group(1))
        return nominal, tol, tol

    up = re.search(r"\+\s*(\d+(?:\.\d+)?)", text)
    low = re.search(r"(?:/|-)\s*(\d+(?:\.\d+)?)", text)

    upper = float(up.group(1)) if up else None
    lower = float(low.group(1)) if low else None

    return nominal, upper, lower


def suggest_method(requirement: str, designator: str, ctype: str):
    text = f"{requirement} {designator} {ctype}".casefold()

    if any(t in text for t in ["m6", "m8", "m10", "m12", "diş", "thread"]):
        return "Diş uygunluk kontrolü", "GO-NO GO diş tampon mastarı"

    if any(t in text for t in ["position", "pozisyon", "profile", "profil", "gd&t"]):
        return "GD&T ölçümü", "CMM"

    if "datum" in text:
        return "Datum doğrulama ve hizalama", "CMM / kontrol fikstürü"

    if any(t in text for t in ["flatness", "düzlemsellik"]):
        return "Düzlemsellik kontrolü", "CMM / granit pleyt + komparatör"

    if any(t in text for t in ["roughness", "pürüz", " ra"]):
        return "Yüzey pürüzlülüğü ölçümü", "Profilometre"

    if any(t in text for t in ["h7", "delik", "bore", "iç çap"]):
        return "Delik çapı kontrolü", "İç çap komparatörü / GO-NO GO tampon mastar / CMM"

    if any(t in text for t in ["ø", "⌀", "çap", "diameter"]):
        return "Çap ölçümü", "Mikrometre / CMM"

    if any(t in text for t in ["pah", "chamfer"]):
        return "Pah kontrolü", "Pah mastarı / profil projektör"

    if any(t in text for t in ["radyüs", "radius", " r"]):
        return "Radyüs kontrolü", "Radyüs mastarı / profil projektör"

    if any(t in text for t in ["material", "malzeme", "process", "proses", "kaplama", "pasivasyon"]):
        return "Doküman / sertifika kontrolü", "Teknik resim / sertifika / proses kaydı"

    if any(t in text for t in ["break", "edge", "çapak"]):
        return "Görsel kontrol", "Görsel kontrol / pah mastarı"

    return "Boyutsal ölçüm", "Dijital kumpas / mikrometre / yükseklik mihengiri"


def make_record(number: int, page: int, target_x: float, target_y: float):
    balloon_x = min(0.96, max(0.04, target_x + 0.07))
    balloon_y = min(0.96, max(0.04, target_y - 0.05))

    return {
        "Characteristic No / Balon No": number,
        "Sheet": page,
        "Zone": "Belirsiz",
        "View": "Belirsiz",
        "Characteristic Type": "Dimension",
        "Characteristic Designator": "",
        "Drawing Requirement": "",
        "Nominal": "Belirsiz",
        "Upper Tolerance": "Belirsiz",
        "Lower Tolerance": "Belirsiz",
        "Upper Limit": "Uygulanmaz",
        "Lower Limit": "Uygulanmaz",
        "Units": "mm",
        "GD&T / Datum Reference": "Uygulanmaz",
        "Quantity": 1,
        "Inspection Method": "Boyutsal ölçüm",
        "Measuring Equipment": "Dijital kumpas / mikrometre / yükseklik mihengiri",
        "Designed / Qualified Tooling": "Yok",
        "Result 1": "Ölçüm bekleniyor",
        "Result 2": "Ölçüm bekleniyor",
        "Result 3": "Ölçüm bekleniyor",
        "Result Summary": "Ölçüm bekleniyor",
        "Acceptance Status": "Bekliyor",
        "Nonconformance No": "Doldurulacak",
        "Inspector": "Doldurulacak",
        "Inspection Date": "Doldurulacak",
        "Remarks": "",
        "_target_x": target_x,
        "_target_y": target_y,
        "_balloon_x": balloon_x,
        "_balloon_y": balloon_y,
    }


def normalize_records(df: pd.DataFrame):
    df = df.copy()

    for col in FORM3_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if df.empty:
        return pd.DataFrame(columns=FORM3_COLUMNS)

    for idx, row in df.iterrows():
        req = str(row.get("Drawing Requirement", "") or "")

        nominal = as_number(row.get("Nominal"))
        upper = as_number(row.get("Upper Tolerance"))
        lower = as_number(row.get("Lower Tolerance"))

        if req and nominal is None:
            p_nom, p_up, p_low = parse_requirement(req)

            if p_nom is not None:
                nominal = p_nom
                df.at[idx, "Nominal"] = nominal

            if upper is None and p_up is not None:
                upper = p_up
                df.at[idx, "Upper Tolerance"] = upper

            if lower is None and p_low is not None:
                lower = p_low
                df.at[idx, "Lower Tolerance"] = lower

        if nominal is not None and upper is not None:
            df.at[idx, "Upper Limit"] = round(nominal + upper, 6)

        if nominal is not None and lower is not None:
            df.at[idx, "Lower Limit"] = round(nominal - abs(lower), 6)

        method, equipment = suggest_method(
            req,
            str(row.get("Characteristic Designator", "") or ""),
            str(row.get("Characteristic Type", "") or ""),
        )

        current_method = str(row.get("Inspection Method", "") or "")
        current_equipment = str(row.get("Measuring Equipment", "") or "")

        if current_method.strip() == "" or current_method == "Boyutsal ölçüm":
            df.at[idx, "Inspection Method"] = method

        if current_equipment.strip() == "" or "Dijital kumpas" in current_equipment:
            df.at[idx, "Measuring Equipment"] = equipment

        for result_col in ["Result 1", "Result 2", "Result 3", "Result Summary"]:
            if not str(row.get(result_col, "") or "").strip():
                df.at[idx, result_col] = "Ölçüm bekleniyor"

        if not str(row.get("Acceptance Status", "") or "").strip():
            df.at[idx, "Acceptance Status"] = "Bekliyor"

    return df[FORM3_COLUMNS]


def load_font(size: int):
    candidates = [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/calibrib.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]

    for path in candidates:
        if Path(path).exists():
            return ImageFont.truetype(path, size)

    return ImageFont.load_default()


def draw_preview(image_bytes: bytes, rows: pd.DataFrame, page_no: int):
    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    draw = ImageDraw.Draw(image)

    font = load_font(max(18, round(image.width / 65)))
    radius = max(18, round(image.width / 70))

    for _, row in rows.iterrows():
        if int(as_number(row.get("Sheet")) or 0) != page_no:
            continue

        number = int(as_number(row.get("Characteristic No / Balon No")) or 0)

        if number <= 0:
            continue

        target = (
            float(row["_target_x"]) * image.width,
            float(row["_target_y"]) * image.height,
        )

        balloon = (
            float(row["_balloon_x"]) * image.width,
            float(row["_balloon_y"]) * image.height,
        )

        if math.dist(target, balloon) > radius * 1.5:
            draw.line([balloon, target], fill=RED, width=max(2, radius // 8))
            draw.ellipse(
                [target[0] - 3, target[1] - 3, target[0] + 3, target[1] + 3],
                fill=RED,
            )

        draw.ellipse(
            [
                balloon[0] - radius,
                balloon[1] - radius,
                balloon[0] + radius,
                balloon[1] + radius,
            ],
            outline=RED,
            fill="white",
            width=max(3, radius // 6),
        )

        text = str(number)
        bbox = draw.textbbox((0, 0), text, font=font)

        draw.text(
            (
                balloon[0] - (bbox[2] - bbox[0]) / 2,
                balloon[1] - (bbox[3] - bbox[1]) / 2 - 2,
            ),
            text,
            fill=RED,
            font=font,
        )

    return image


def create_balloon_pdf(pdf_bytes: bytes, records: pd.DataFrame):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    red = (0.86, 0.04, 0.12)

    try:
        for page_index, page in enumerate(doc, start=1):
            rows = records[
                records["Sheet"].apply(lambda v: int(as_number(v) or 0) == page_index)
            ]

            for _, row in rows.iterrows():
                number = int(as_number(row.get("Characteristic No / Balon No")) or 0)

                if number <= 0:
                    continue

                target = fitz.Point(
                    float(row["_target_x"]) * page.rect.width,
                    float(row["_target_y"]) * page.rect.height,
                )

                balloon = fitz.Point(
                    float(row["_balloon_x"]) * page.rect.width,
                    float(row["_balloon_y"]) * page.rect.height,
                )

                radius = 11

                if math.dist((target.x, target.y), (balloon.x, balloon.y)) > radius * 1.7:
                    angle = math.atan2(target.y - balloon.y, target.x - balloon.x)
                    start = fitz.Point(
                        balloon.x + radius * math.cos(angle),
                        balloon.y + radius * math.sin(angle),
                    )
                    page.draw_line(start, target, color=red, width=1.2, overlay=True)
                    page.draw_circle(target, radius=2, color=red, fill=red, overlay=True)

                page.draw_circle(
                    balloon,
                    radius=radius,
                    color=red,
                    fill=(1, 1, 1),
                    width=1.7,
                    overlay=True,
                )

                text = str(number)
                width = fitz.get_text_length(text, fontname="helv", fontsize=9)

                page.insert_text(
                    fitz.Point(balloon.x - width / 2, balloon.y + 3.2),
                    text,
                    fontname="helv",
                    fontsize=9,
                    color=red,
                    overlay=True,
                )

        out = io.BytesIO()
        doc.save(out, garbage=4, deflate=True)
        return out.getvalue()

    finally:
        doc.close()


def border():
    side = Side(style="thin", color="4F4F4F")
    return Border(left=side, right=side, top=side, bottom=side)


def style_title(ws, title: str, end_col: int):
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=end_col)
    cell = ws.cell(1, 1, title)
    cell.font = Font(name="Arial", size=15, bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.sheet_view.showGridLines = False


def create_excel(records: pd.DataFrame, part_info: dict[str, str], process_info: dict[str, str]):
    wb = Workbook()

    ws = wb.active
    ws.title = "AS9102_Form3_Olcum_Raporu"

    style_title(ws, "AS9102 FAI FORM 3 - ÖLÇÜM RAPORU", len(VISIBLE_COLUMNS))

    b = border()
    header_fill = PatternFill("solid", fgColor="D9E2F3")

    for col_idx, title in enumerate(VISIBLE_COLUMNS, start=1):
        cell = ws.cell(4, col_idx, title)
        cell.font = Font(name="Arial", size=9, bold=True)
        cell.fill = header_fill
        cell.border = b
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    sorted_records = records.sort_values("Characteristic No / Balon No", kind="stable")

    for row_idx, (_, rec) in enumerate(sorted_records.iterrows(), start=5):
        for col_idx, key in enumerate(VISIBLE_COLUMNS, start=1):
            value = rec.get(key, "")

            if value is None or (not isinstance(value, str) and pd.isna(value)):
                value = ""

            cell = ws.cell(row_idx, col_idx, value)
            cell.border = b
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.font = Font(name="Arial", size=9)

            if key == "Characteristic No / Balon No":
                cell.font = Font(name="Arial", size=10, bold=True, color="C00000")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(VISIBLE_COLUMNS))}{max(5, ws.max_row)}"

    for col_idx in range(1, len(VISIBLE_COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 18

    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["P"].width = 26
    ws.column_dimensions["Q"].width = 30
    ws.column_dimensions["AA"].width = 32

    f1 = wb.create_sheet("Form1_Parca_Bilgileri")
    style_title(f1, "AS9102 FORM 1 - PARÇA BİLGİLERİ", 4)

    for r, (k, v) in enumerate(part_info.items(), start=4):
        f1.cell(r, 1, k).font = Font(bold=True)
        f1.cell(r, 1).fill = header_fill
        f1.cell(r, 2, v or "Doldurulacak")
        f1.cell(r, 1).border = b
        f1.cell(r, 2).border = b
        f1.column_dimensions["A"].width = 26
        f1.column_dimensions["B"].width = 35

    f2 = wb.create_sheet("Form2_Malzeme_Proses")
    style_title(f2, "AS9102 FORM 2 - MALZEME VE PROSES", 4)

    for r, (k, v) in enumerate(process_info.items(), start=4):
        f2.cell(r, 1, k).font = Font(bold=True)
        f2.cell(r, 1).fill = header_fill
        f2.cell(r, 2, v or "Doldurulacak")
        f2.cell(r, 1).border = b
        f2.cell(r, 2).border = b
        f2.column_dimensions["A"].width = 26
        f2.column_dimensions["B"].width = 35

    methods = wb.create_sheet("Olcum_Metodu_Listesi")
    style_title(methods, "ÖLÇÜM METODU VE EKİPMAN ÖNERİ LİSTESİ", 4)

    rows = [
        ("Karakteristik", "Önerilen Metot", "Ölçüm Ekipmanı", "Not"),
        ("Dış çap", "Dış çap ölçümü", "Mikrometre", ""),
        ("İç çap / H7", "Delik çapı kontrolü", "İç çap komparatörü / tampon mastar / CMM", ""),
        ("Diş", "Diş uygunluk kontrolü", "GO-NO GO diş mastarı", ""),
        ("Doğrusal", "Boyutsal ölçüm", "Kumpas / mikrometre / yükseklik mihengiri", ""),
        ("GD&T", "Koordinat ölçümü", "CMM", ""),
        ("Yüzey", "Pürüzlülük ölçümü", "Profilometre", ""),
        ("Not/proses", "Doküman kontrolü", "Sertifika / proses kaydı", ""),
    ]

    for r, row in enumerate(rows, start=4):
        for c, v in enumerate(row, start=1):
            cell = methods.cell(r, c, v)
            cell.border = b
            cell.alignment = Alignment(wrap_text=True, vertical="top")

            if r == 4:
                cell.font = Font(bold=True)
                cell.fill = header_fill

    for c, w in enumerate([24, 28, 42, 30], start=1):
        methods.column_dimensions[get_column_letter(c)].width = w

    out = io.BytesIO()
    wb.save(out)

    return out.getvalue()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    st.title("AS9102 Teknik Resim Balonlama")
    st.caption(
        "PDF üzerinde sade kırmızı numaralı balonlar oluşturur; "
        "tüm detaylar AS9102 Excel raporuna aktarılır."
    )

    with st.sidebar:
        uploaded = st.file_uploader("PDF teknik resmi yükleyin", type=["pdf"])
        st.info("SLDDRW desteklenmez. Teknik resmin PDF çıktısını yükleyin.")

    if uploaded is None:
        st.info("Başlamak için PDF dosyası yükleyin.")
        return

    pdf_bytes = uploaded.getvalue()

    if st.session_state.pdf_name != uploaded.name or st.session_state.pdf_bytes != pdf_bytes:
        try:
            images = render_pdf(pdf_bytes)
        except Exception as exc:
            st.error(f"PDF okunamadı: {exc}")
            return

        st.session_state.pdf_name = uploaded.name
        st.session_state.pdf_bytes = pdf_bytes
        st.session_state.page_images = images
        st.session_state.records = pd.DataFrame(columns=FORM3_COLUMNS)
        st.session_state.last_click = None
        st.session_state.viewer_key += 1

    st.success(f"{uploaded.name} yüklendi · {len(st.session_state.page_images)} sayfa")

    page_no = st.selectbox(
        "Sayfa seç",
        range(1, len(st.session_state.page_images) + 1),
        format_func=lambda n: f"Sayfa {n}",
    )

    base_preview = draw_preview(
        st.session_state.page_images[page_no - 1],
        st.session_state.records,
        page_no,
    )

    ow, oh = base_preview.size
    display_width = min(1200, ow)
    display_height = int(oh * display_width / ow)
    display_image = base_preview.resize(
        (display_width, display_height),
        Image.Resampling.LANCZOS,
    )

    st.subheader("PDF ön izleme ve tıklama alanı")
    st.caption(
        "Ölçü/toleransın yakınına tıklayın. "
        "Sonra “Tıklanan noktayı balona dönüştür” düğmesine basın."
    )

    click = streamlit_image_coordinates(
        display_image,
        key=f"pdf_click_{page_no}_{st.session_state.viewer_key}",
    )

    if click is not None:
        x = click["x"] / display_width
        y = click["y"] / display_height
        st.session_state.last_click = (page_no, x, y)
        st.info(f"Seçilen nokta: Sayfa {page_no}, X={x:.3f}, Y={y:.3f}")

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("Tıklanan noktayı balona dönüştür", type="primary"):
            if st.session_state.last_click is None:
                st.warning("Önce PDF üzerinde bir noktaya tıklayın.")
            else:
                p, x, y = st.session_state.last_click

                current = st.session_state.records.copy()
                existing = [
                    as_number(v) or 0
                    for v in current.get("Characteristic No / Balon No", [])
                ]

                next_no = int(max(existing, default=0)) + 1
                record = make_record(next_no, int(p), float(x), float(y))

                st.session_state.records = pd.concat(
                    [current, pd.DataFrame([record])],
                    ignore_index=True,
                )

                st.session_state.last_click = None
                st.rerun()

    with col2:
        if st.button("Önerileri güncelle"):
            st.session_state.records = normalize_records(st.session_state.records)
            st.rerun()

    with col3:
        if st.button("Tüm kayıtları temizle"):
            st.session_state.records = pd.DataFrame(columns=FORM3_COLUMNS)
            st.session_state.last_click = None
            st.rerun()

    st.subheader("AS9102 Form 3 ölçüm listesi")

    edited = st.data_editor(
        st.session_state.records,
        column_order=FORM3_COLUMNS,
        hide_index=True,
        num_rows="dynamic",
        disabled=["_target_x", "_target_y", "_balloon_x", "_balloon_y"],
        height=380,
    )

    st.session_state.records = normalize_records(edited)

    with st.expander("Balon koordinatları"):
        if st.session_state.records.empty:
            st.write("Henüz balon yok.")
        else:
            coord_cols = [
                "Characteristic No / Balon No",
                "Sheet",
                "_target_x",
                "_target_y",
                "_balloon_x",
                "_balloon_y",
            ]

            coords = st.data_editor(
                st.session_state.records[coord_cols],
                hide_index=True,
                height=220,
            )

            for col in coord_cols[2:]:
                st.session_state.records[col] = coords[col].values

    st.subheader("Form 1 ve Form 2 bilgileri")

    c1, c2 = st.columns(2)

    with c1:
        part_info = {
            "Part Number": st.text_input("Part Number / Parça No"),
            "Part Name": st.text_input("Part Name / Parça Adı"),
            "Drawing Number": st.text_input("Drawing Number", value=Path(uploaded.name).stem),
            "Revision": st.text_input("Revision"),
            "Supplier": st.text_input("Supplier"),
            "Customer": st.text_input("Customer"),
            "FAI Type": st.selectbox("FAI Type", ["Full FAI", "Partial FAI"]),
        }

    with c2:
        process_info = {
            "Material": st.text_input("Material"),
            "Material Specification": st.text_input("Material Specification"),
            "Special Process": st.text_input("Special Process"),
            "Finish": st.text_input("Finish / Kaplama"),
            "Certificate Required": st.text_input("Certificate Required"),
            "Remarks": st.text_area("Form 2 Remarks"),
        }

    if st.session_state.records.empty:
        st.warning("Çıktı almak için en az bir balon ekleyin.")
        return

    try:
        final_records = normalize_records(st.session_state.records)
        pdf_out = create_balloon_pdf(pdf_bytes, final_records)
        excel_out = create_excel(final_records, part_info, process_info)
    except Exception as exc:
        st.error(f"Çıktılar hazırlanamadı: {exc}")
        return

    st.subheader("Çıktılar")

    d1, d2 = st.columns(2)

    with d1:
        st.download_button(
            "Sade balonlu PDF indir",
            data=pdf_out,
            file_name="balonlu_teknik_resim_sade.pdf",
            mime="application/pdf",
        )

    with d2:
        st.download_button(
            "AS9102 FAI Excel indir",
            data=excel_out,
            file_name="AS9102_FAI_olcum_raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
