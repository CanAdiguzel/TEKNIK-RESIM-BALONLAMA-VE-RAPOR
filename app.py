from __future__ import annotations

import io
import math
import re
from pathlib import Path

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
DPI = 120

COLS = [
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
    "_x",
    "_y",
    "_bx",
    "_by",
]

VISIBLE = COLS[:-4]


def init_state():
    defaults = {
        "pdf_bytes": None,
        "pdf_name": None,
        "images": [],
        "records": pd.DataFrame(columns=COLS),
        "last_click": None,
        "viewer_key": 0,
    }

    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


@st.cache_data(show_spinner=False)
def render_pdf(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    zoom = DPI / 72
    matrix = fitz.Matrix(zoom, zoom)
    images = []

    try:
        for page in doc:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            images.append(pix.tobytes("png"))
    finally:
        doc.close()

    return images


def num(v):
    if v is None:
        return None

    try:
        if pd.isna(v):
            return None
    except Exception:
        pass

    s = str(v).replace(",", ".").strip()

    if not s:
        return None

    try:
        return float(s)
    except Exception:
        return None


def classify(text: str):
    t = text.casefold()

    if "position" in t or "pozisyon" in t:
        return "GD&T - Position"
    if "profile" in t or "profil" in t:
        return "GD&T - Profile"
    if "flatness" in t or "düzlemsellik" in t:
        return "GD&T - Flatness"
    if "datum" in t:
        return "Datum"
    if "roughness" in t or "ra " in t or "pürüz" in t:
        return "Surface Finish"
    if re.search(r"\bm\d+", t):
        return "Thread"
    if "ø" in t or "⌀" in t or "h7" in t or "çap" in t:
        return "Diameter"
    if " r" in t or t.startswith("r"):
        return "Radius"
    if "°" in t or "x45" in t:
        return "Angle"
    if "material" in t or "malzeme" in t or "process" in t or "proses" in t:
        return "Material / Process"
    if "break" in t or "edge" in t or "chamfer" in t:
        return "Visual Requirement"

    return "Dimension"


def suggest_method(req: str, ctype: str):
    t = f"{req} {ctype}".casefold()

    if any(x in t for x in ["m6", "m8", "m10", "m12", "diş", "thread"]):
        return "Diş uygunluk kontrolü", "GO-NO GO diş tampon mastarı"

    if any(x in t for x in ["position", "pozisyon", "profile", "profil", "gd&t"]):
        return "GD&T ölçümü", "CMM"

    if "datum" in t:
        return "Datum doğrulama ve hizalama", "CMM / kontrol fikstürü"

    if "flatness" in t or "düzlemsellik" in t:
        return "Düzlemsellik kontrolü", "CMM / granit pleyt + komparatör"

    if "roughness" in t or "pürüz" in t or "ra " in t:
        return "Yüzey pürüzlülüğü ölçümü", "Profilometre"

    if "h7" in t or "delik" in t or "iç çap" in t:
        return "Delik çapı kontrolü", "İç çap komparatörü / GO-NO GO tampon mastar / CMM"

    if "ø" in t or "⌀" in t or "çap" in t:
        return "Çap ölçümü", "Mikrometre / CMM"

    if "pah" in t or "chamfer" in t:
        return "Pah kontrolü", "Pah mastarı / profil projektör"

    if "radyüs" in t or "radius" in t:
        return "Radyüs kontrolü", "Radyüs mastarı / profil projektör"

    if "material" in t or "malzeme" in t or "proses" in t or "process" in t:
        return "Doküman / sertifika kontrolü", "Teknik resim / sertifika / proses kaydı"

    return "Boyutsal ölçüm", "Dijital kumpas / mikrometre / yükseklik mihengiri"


def parse_req(text: str):
    s = text.replace(",", ".").replace("−", "-").upper()

    m = re.search(r"(?:Ø|⌀|R)?\s*(\d+(?:\.\d+)?)", s)
    nominal = float(m.group(1)) if m else None

    pm = re.search(r"±\s*(\d+(?:\.\d+)?)", s)

    if pm:
        tol = float(pm.group(1))
        return nominal, tol, tol

    up = re.search(r"\+\s*(\d+(?:\.\d+)?)", s)
    low = re.search(r"(?:/|-)\s*(\d+(?:\.\d+)?)", s)

    return (
        nominal,
        float(up.group(1)) if up else None,
        float(low.group(1)) if low else None,
    )


def new_record(no: int, page: int, x: float, y: float, req: str = ""):
    ctype = classify(req)
    method, equipment = suggest_method(req, ctype)

    return {
        "Characteristic No / Balon No": no,
        "Sheet": page,
        "Zone": "Belirsiz",
        "View": "Belirsiz",
        "Characteristic Type": ctype,
        "Characteristic Designator": req[:60],
        "Drawing Requirement": req,
        "Nominal": "Belirsiz",
        "Upper Tolerance": "Belirsiz",
        "Lower Tolerance": "Belirsiz",
        "Upper Limit": "Uygulanmaz",
        "Lower Limit": "Uygulanmaz",
        "Units": "mm",
        "GD&T / Datum Reference": "Uygulanmaz",
        "Quantity": 1,
        "Inspection Method": method,
        "Measuring Equipment": equipment,
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
        "_x": x,
        "_y": y,
        "_bx": min(0.96, max(0.04, x + 0.07)),
        "_by": min(0.96, max(0.04, y - 0.05)),
    }


def normalize(df: pd.DataFrame):
    df = df.copy()

    for c in COLS:
        if c not in df.columns:
            df[c] = ""

    if df.empty:
        return pd.DataFrame(columns=COLS)

    for i, r in df.iterrows():
        req = str(r.get("Drawing Requirement", "") or "")

        nominal = num(r.get("Nominal"))
        upper = num(r.get("Upper Tolerance"))
        lower = num(r.get("Lower Tolerance"))

        if req and nominal is None:
            n, u, l = parse_req(req)

            if n is not None:
                nominal = n
                df.at[i, "Nominal"] = n

            if upper is None and u is not None:
                upper = u
                df.at[i, "Upper Tolerance"] = u

            if lower is None and l is not None:
                lower = l
                df.at[i, "Lower Tolerance"] = l

        if nominal is not None and upper is not None:
            df.at[i, "Upper Limit"] = round(nominal + upper, 6)

        if nominal is not None and lower is not None:
            df.at[i, "Lower Limit"] = round(nominal - abs(lower), 6)

        ctype = str(r.get("Characteristic Type", "") or "")
        method, equipment = suggest_method(req, ctype)

        if not str(r.get("Inspection Method", "") or "").strip():
            df.at[i, "Inspection Method"] = method

        if not str(r.get("Measuring Equipment", "") or "").strip():
            df.at[i, "Measuring Equipment"] = equipment

        for rc in ["Result 1", "Result 2", "Result 3", "Result Summary"]:
            if not str(r.get(rc, "") or "").strip():
                df.at[i, rc] = "Ölçüm bekleniyor"

        if not str(r.get("Acceptance Status", "") or "").strip():
            df.at[i, "Acceptance Status"] = "Bekliyor"

    return df[COLS]


def is_candidate(text: str):
    t = text.strip()
    u = t.upper()

    if len(t) < 2:
        return False

    keys = [
        "±",
        "+",
        "Ø",
        "⌀",
        "H7",
        "M6",
        "M8",
        "M10",
        "M12",
        "POSITION",
        "PROFILE",
        "FLATNESS",
        "DATUM",
        "BREAK",
        "CHAMFER",
        "ROUGHNESS",
        "MATERIAL",
        "PROCESS",
        "RA ",
    ]

    if any(k in u for k in keys):
        return True

    if re.search(r"\d+[,.]?\d*\s*[xX]\s*\d+", t):
        return True

    if re.search(r"\bR\s*\d", u):
        return True

    if re.search(r"\d+[,.]?\d*\s*°", t):
        return True

    return False


def extract_candidates(pdf_bytes: bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rows = []
    no = 1

    try:
        for page_no, page in enumerate(doc, start=1):
            words = page.get_text("words")
            lines = {}

            for item in words:
                x0, y0, x1, y1, word, block, line, word_no = item
                key = (block, line)
                lines.setdefault(key, []).append((x0, y0, x1, y1, word))

            for _, ws in lines.items():
                ws = sorted(ws, key=lambda a: a[0])
                text = " ".join(w[4] for w in ws).strip()

                if not is_candidate(text):
                    continue

                x0 = min(w[0] for w in ws)
                y0 = min(w[1] for w in ws)
                x1 = max(w[2] for w in ws)
                y1 = max(w[3] for w in ws)

                x = ((x0 + x1) / 2) / page.rect.width
                y = ((y0 + y1) / 2) / page.rect.height

                row = new_record(no, page_no, x, y, text)
                row["Remarks"] = "PDF metninden otomatik aday olarak çıkarıldı; teknik olarak doğrulayın."

                rows.append(row)
                no += 1

    finally:
        doc.close()

    return pd.DataFrame(rows, columns=COLS)


def font(size: int):
    paths = [
        "C:/Windows/Fonts/arialbd.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    ]

    for p in paths:
        if Path(p).exists():
            return ImageFont.truetype(p, size)

    return ImageFont.load_default()


def draw_preview(image_bytes: bytes, df: pd.DataFrame, page_no: int):
    img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    d = ImageDraw.Draw(img)

    f = font(max(18, round(img.width / 65)))
    radius = max(18, round(img.width / 70))

    for _, r in df.iterrows():
        if int(num(r.get("Sheet")) or 0) != page_no:
            continue

        no = int(num(r.get("Characteristic No / Balon No")) or 0)

        if no <= 0:
            continue

        x = float(r["_x"]) * img.width
        y = float(r["_y"]) * img.height
        bx = float(r["_bx"]) * img.width
        by = float(r["_by"]) * img.height

        if math.dist((x, y), (bx, by)) > radius * 1.5:
            d.line([(bx, by), (x, y)], fill=RED, width=max(2, radius // 8))
            d.ellipse([x - 3, y - 3, x + 3, y + 3], fill=RED)

        d.ellipse(
            [bx - radius, by - radius, bx + radius, by + radius],
            outline=RED,
            fill="white",
            width=max(3, radius // 6),
        )

        text = str(no)
        bb = d.textbbox((0, 0), text, font=f)

        d.text(
            (bx - (bb[2] - bb[0]) / 2, by - (bb[3] - bb[1]) / 2 - 2),
            text,
            fill=RED,
            font=f,
        )

    return img


def export_pdf(pdf_bytes: bytes, df: pd.DataFrame):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    red = (0.86, 0.04, 0.12)

    try:
        for page_no, page in enumerate(doc, start=1):
            rows = df[df["Sheet"].apply(lambda v: int(num(v) or 0) == page_no)]

            for _, r in rows.iterrows():
                no = int(num(r.get("Characteristic No / Balon No")) or 0)

                if no <= 0:
                    continue

                target = fitz.Point(float(r["_x"]) * page.rect.width, float(r["_y"]) * page.rect.height)
                balloon = fitz.Point(float(r["_bx"]) * page.rect.width, float(r["_by"]) * page.rect.height)

                radius = 11

                if math.dist((target.x, target.y), (balloon.x, balloon.y)) > radius * 1.7:
                    ang = math.atan2(target.y - balloon.y, target.x - balloon.x)
                    start = fitz.Point(
                        balloon.x + radius * math.cos(ang),
                        balloon.y + radius * math.sin(ang),
                    )
                    page.draw_line(start, target, color=red, width=1.2, overlay=True)
                    page.draw_circle(target, radius=2, color=red, fill=red, overlay=True)

                page.draw_circle(balloon, radius=radius, color=red, fill=(1, 1, 1), width=1.7, overlay=True)

                tw = fitz.get_text_length(str(no), fontname="helv", fontsize=9)

                page.insert_text(
                    fitz.Point(balloon.x - tw / 2, balloon.y + 3.2),
                    str(no),
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


def excel_border():
    s = Side(style="thin", color="4F4F4F")
    return Border(left=s, right=s, top=s, bottom=s)


def style_title(ws, title: str, end_col: int):
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=end_col)
    c = ws.cell(1, 1, title)
    c.font = Font(name="Arial", size=15, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F4E78")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.sheet_view.showGridLines = False


def export_excel(df: pd.DataFrame, part_info: dict, process_info: dict):
    wb = Workbook()
    b = excel_border()
    fill = PatternFill("solid", fgColor="D9E2F3")

    ws = wb.active
    ws.title = "AS9102_Form3_Olcum_Raporu"

    style_title(ws, "AS9102 FAI FORM 3 - ÖLÇÜM RAPORU", len(VISIBLE))

    for col, title in enumerate(VISIBLE, start=1):
        c = ws.cell(4, col, title)
        c.font = Font(size=9, bold=True)
        c.fill = fill
        c.border = b
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    df = df.sort_values("Characteristic No / Balon No", kind="stable")

    for row_i, (_, r) in enumerate(df.iterrows(), start=5):
        for col_i, key in enumerate(VISIBLE, start=1):
            val = r.get(key, "")

            if val is None or (not isinstance(val, str) and pd.isna(val)):
                val = ""

            c = ws.cell(row_i, col_i, val)
            c.border = b
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.font = Font(size=9)

            if key == "Characteristic No / Balon No":
                c.font = Font(size=10, bold=True, color="C00000")
                c.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(VISIBLE))}{max(5, ws.max_row)}"

    for i in range(1, len(VISIBLE) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 18

    ws.column_dimensions["G"].width = 30
    ws.column_dimensions["P"].width = 26
    ws.column_dimensions["Q"].width = 30
    ws.column_dimensions["AA"].width = 35

    f1 = wb.create_sheet("Form1_Parca_Bilgileri")
    style_title(f1, "AS9102 FORM 1 - PARÇA BİLGİLERİ", 4)

    for i, (k, v) in enumerate(part_info.items(), start=4):
        f1.cell(i, 1, k).font = Font(bold=True)
        f1.cell(i, 1).fill = fill
        f1.cell(i, 2, v or "Doldurulacak")
        f1.cell(i, 1).border = b
        f1.cell(i, 2).border = b
        f1.column_dimensions["A"].width = 26
        f1.column_dimensions["B"].width = 38

    f2 = wb.create_sheet("Form2_Malzeme_Proses")
    style_title(f2, "AS9102 FORM 2 - MALZEME VE PROSES", 4)

    for i, (k, v) in enumerate(process_info.items(), start=4):
        f2.cell(i, 1, k).font = Font(bold=True)
        f2.cell(i, 1).fill = fill
        f2.cell(i, 2, v or "Doldurulacak")
        f2.cell(i, 1).border = b
        f2.cell(i, 2).border = b
        f2.column_dimensions["A"].width = 26
        f2.column_dimensions["B"].width = 38

    m = wb.create_sheet("Olcum_Metodu_Listesi")
    style_title(m, "ÖLÇÜM METODU VE EKİPMAN ÖNERİ LİSTESİ", 4)

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
        for c, val in enumerate(row, start=1):
            cell = m.cell(r, c, val)
            cell.border = b
            cell.alignment = Alignment(wrap_text=True, vertical="top")

            if r == 4:
                cell.font = Font(bold=True)
                cell.fill = fill

    for c, w in enumerate([24, 28, 42, 30], start=1):
        m.column_dimensions[get_column_letter(c)].width = w

    out = io.BytesIO()
    wb.save(out)

    return out.getvalue()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    init_state()

    st.title("AS9102 Teknik Resim Balonlama")
    st.caption("PDF üzerinde sade kırmızı numaralı balonlar oluşturur. Detaylar AS9102 Excel raporuna aktarılır.")

    with st.sidebar:
        uploaded = st.file_uploader("PDF teknik resmi yükleyin", type=["pdf"])
        st.info("SLDDRW desteklenmez. PDF çıktısı yükleyin.")

    if uploaded is None:
        st.info("Başlamak için PDF dosyası yükleyin.")
        return

    pdf_bytes = uploaded.getvalue()

    if st.session_state.pdf_name != uploaded.name or st.session_state.pdf_bytes != pdf_bytes:
        try:
            st.session_state.images = render_pdf(pdf_bytes)
        except Exception as e:
            st.error(f"PDF okunamadı: {e}")
            return

        st.session_state.pdf_bytes = pdf_bytes
        st.session_state.pdf_name = uploaded.name
        st.session_state.records = pd.DataFrame(columns=COLS)
        st.session_state.last_click = None
        st.session_state.viewer_key += 1

    st.success(f"{uploaded.name} yüklendi · {len(st.session_state.images)} sayfa")

    scan_col, note_col = st.columns([1, 2])

    with scan_col:
        if st.button("PDF yazılarından aday çıkar"):
            with st.spinner("PDF içindeki seçilebilir metinler taranıyor..."):
                candidates = extract_candidates(pdf_bytes)

            if candidates.empty:
                st.warning("Aday bulunamadı. PDF taranmış/görsel olabilir. Manuel tıklama ile balon ekleyin.")
            else:
                st.session_state.records = normalize(candidates)
                st.session_state.viewer_key += 1
                st.success(f"{len(candidates)} adet aday karakteristik çıkarıldı.")
                st.rerun()

    with note_col:
        st.caption("Bu gerçek OCR değildir; PDF içindeki seçilebilir yazıları okur. Taranmış PDF’lerde çalışmayabilir.")

    page_no = st.selectbox(
        "Sayfa seç",
        range(1, len(st.session_state.images) + 1),
        format_func=lambda n: f"Sayfa {n}",
    )

    base = draw_preview(st.session_state.images[page_no - 1], st.session_state.records, page_no)

    ow, oh = base.size
    dw = min(1200, ow)
    dh = int(oh * dw / ow)

    display = base.resize((dw, dh), Image.Resampling.LANCZOS)

    st.subheader("PDF ön izleme ve tıklama alanı")
    st.caption("Ölçü/toleransın yakınına tıklayın. Sonra butona basıp balona dönüştürün.")

    click = streamlit_image_coordinates(display, key=f"pdf_{page_no}_{st.session_state.viewer_key}")

    if click is not None:
        x = click["x"] / dw
        y = click["y"] / dh
        st.session_state.last_click = (page_no, x, y)
        st.info(f"Seçilen nokta: Sayfa {page_no}, X={x:.3f}, Y={y:.3f}")

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("Tıklanan noktayı balona dönüştür", type="primary"):
            if st.session_state.last_click is None:
                st.warning("Önce PDF üzerinde bir noktaya tıklayın.")
            else:
                p, x, y = st.session_state.last_click

                current = st.session_state.records.copy()
                existing = [num(v) or 0 for v in current.get("Characteristic No / Balon No", [])]
                next_no = int(max(existing, default=0)) + 1

                row = new_record(next_no, int(p), float(x), float(y))
                st.session_state.records = pd.concat([current, pd.DataFrame([row])], ignore_index=True)
                st.session_state.last_click = None
                st.rerun()

    with c2:
        if st.button("Önerileri güncelle"):
            st.session_state.records = normalize(st.session_state.records)
            st.rerun()

    with c3:
        if st.button("Tüm kayıtları temizle"):
            st.session_state.records = pd.DataFrame(columns=COLS)
            st.session_state.last_click = None
            st.rerun()

    st.subheader("AS9102 Form 3 ölçüm listesi")

    edited = st.data_editor(
        st.session_state.records,
        column_order=COLS,
        hide_index=True,
        num_rows="dynamic",
        disabled=["_x", "_y", "_bx", "_by"],
        height=380,
    )

    st.session_state.records = normalize(edited)

    with st.expander("Balon koordinatları"):
        if st.session_state.records.empty:
            st.write("Henüz balon yok.")
        else:
            coord_cols = ["Characteristic No / Balon No", "Sheet", "_x", "_y", "_bx", "_by"]
            coords = st.data_editor(st.session_state.records[coord_cols], hide_index=True, height=220)

            for col in coord_cols[2:]:
                st.session_state.records[col] = coords[col].values

    st.subheader("Form 1 ve Form 2 bilgileri")

    f1, f2 = st.columns(2)

    with f1:
        part_info = {
            "Part Number": st.text_input("Part Number / Parça No"),
            "Part Name": st.text_input("Part Name / Parça Adı"),
            "Drawing Number": st.text_input("Drawing Number", value=Path(uploaded.name).stem),
            "Revision": st.text_input("Revision"),
            "Supplier": st.text_input("Supplier"),
            "Customer": st.text_input("Customer"),
            "FAI Type": st.selectbox("FAI Type", ["Full FAI", "Partial FAI"]),
        }

    with f2:
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
        final_df = normalize(st.session_state.records)
        pdf_out = export_pdf(pdf_bytes, final_df)
        xlsx_out = export_excel(final_df, part_info, process_info)
    except Exception as e:
        st.error(f"Çıktılar hazırlanamadı: {e}")
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
            data=xlsx_out,
            file_name="AS9102_FAI_olcum_raporu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
