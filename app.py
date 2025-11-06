import io, os, re
from datetime import datetime, date
from typing import List, Tuple, Any

import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2._page import PageObject
from PyPDF2.generic import ArrayObject, NameObject
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Kersia PDF Stamper", page_icon="🧰", layout="centered")
st.title("Kersia — PDF Stamper (Adobe-safe build v3)")

# ---------------- Helpers ----------------

def _coerce_int(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, (datetime, date)):
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:  # NaN
            return 0
        return int(round(value))
    s = str(value).strip()
    if not s:
        return 0
    m = re.search(r'-?\d+', s.replace(',', '.'))
    return int(m.group(0)) if m else 0

def _to_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v)

def _parse_excel(excel_bytes: bytes) -> List[Tuple[str, int, str]]:
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    rows: List[Tuple[str, int, str]] = []
    header = {str((ws.cell(1, c).value or "")).strip().lower(): c for c in range(1, ws.max_column+1)}
    col_z = header.get("zlecenie", 1)
    col_i = header.get("ilość palet", header.get("ilosc palet", 2))
    col_p = header.get("przewoźnik", header.get("przewoznik", 3))

    for r in range(2, ws.max_row+1):
        z = ws.cell(r, col_z).value
        i = ws.cell(r, col_i).value
        p = ws.cell(r, col_p).value
        if z is None and i is None and p is None:
            continue
        rows.append((_to_str(z).strip(), _coerce_int(i), _to_str(p).strip()))
    return rows

def _register_fonts():
    try_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf",
        os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf"),
    ]
    for p in try_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("DejaVuSans", p))
            return "DejaVuSans"
    return "Helvetica"

def _make_stamp_page(zlecenie: str, ilosc: int, przewoznik: str, width: float, height: float) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), bottomup=True)
    font_name = _register_fonts()
    c.setAuthor("Kersia PDF Stamper")
    c.setTitle(f"Zlecenie {zlecenie}")
    c.setCreator("Kersia PDF Stamper (Adobe-safe)")

    c.setFont(font_name, 14)
    margin = 15 * mm
    x = margin
    y = height - margin
    c.drawString(x, y, f"ZLECENIE: {zlecenie}")
    y -= 8 * mm
    c.drawString(x, y, f"ILOŚĆ PALET: {ilosc}")
    y -= 8 * mm
    c.drawString(x, y, f"PRZEWOŹNIK: {przewoznik}")
    c.showPage()
    c.save()
    return buf.getvalue()

def _normalize_contents(page: PageObject):
    contents = page.get(NameObject("/Contents"))
    if contents is None:
        return
    if not isinstance(contents, ArrayObject):
        page[NameObject("/Contents")] = ArrayObject([contents])

def annotate_pdf(pdf_bytes: bytes, excel_bytes: bytes, max_per_sheet: int = 3) -> bytes:
    reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
    rows = _parse_excel(excel_bytes)
    if not rows:
        raise ValueError("Nie znaleziono danych w Excelu. Upewnij się, że masz kolumny: ZLECENIE, ILOŚĆ PALET, PRZEWOŹNIK.")

    writer = PdfWriter()
    writer.add_metadata({
        "/Producer": "Kersia PDF Stamper (PyPDF2 + ReportLab)",
        "/Creator": "Kersia PDF Stamper (Adobe-safe)",
        "/Title": "Zlecenia",
    })

    data_idx = 0
    for page in reader.pages:
        _normalize_contents(page)
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            zlecenie, ilosc, przewoznik = rows[data_idx]
            data_idx += 1
            overlay_bytes = _make_stamp_page(zlecenie, ilosc, przewoznik, w, h)
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)

    while data_idx < len(rows):
        w = float(reader.pages[-1].mediabox.width)
        h = float(reader.pages[-1].mediabox.height)
        blank = PageObject.create_blank_page(width=w, height=h)
        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            zlecenie, ilosc, przewoznik = rows[data_idx]
            data_idx += 1
            overlay_bytes = _make_stamp_page(zlecenie, ilosc, przewoznik, w, h)
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            blank.merge_page(overlay_reader.pages[0])
        writer.add_page(blank)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()

# ---------------- UI ----------------
excel_file = st.file_uploader("Plik Excel (ZLECENIE, ilość palet, przewoźnik):", type=["xlsx", "xlsm", "xls"])
pdf_file   = st.file_uploader("Plik PDF (szablon/strony do ostemplowania):", type=["pdf"])
max_per_sheet = st.slider("Maks. wpisów na stronę", min_value=1, max_value=6, value=3, step=1)

if st.button("GENERUJ PDF", type="primary", disabled=not (excel_file and pdf_file)):
    try:
        result = annotate_pdf(pdf_file.read(), excel_file.read(), max_per_sheet)
        fname = "zlecenia_{}.pdf".format(datetime.now().strftime('%Y%m%d'))
        st.success("Gotowe! Poniżej przycisk pobierania.")
        st.download_button("Pobierz wynik", data=result, file_name=fname, mime="application/pdf")
    except Exception as e:
        st.error(f"Błąd: {e}")