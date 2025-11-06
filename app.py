import io, re, os
from datetime import datetime
from typing import List, Tuple

import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2._page import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---- Streamlit UI ----
st.set_page_config(page_title="Kersia PDF Stamper", page_icon="🧰", layout="centered")
st.title("Kersia — PDF Stamper (Adobe-safe build)")

# ---- Helpers ----

def _parse_excel(excel_bytes: bytes) -> List[Tuple[str, int, str]]:
    """Return rows of (zlecenie, ilosc_palet, przewoznik)."""
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    rows = []
    # Try to find header by names; fallback to fixed columns A/B/C.
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
        try:
            i = int(i)
        except Exception:
            i = 0 if i is None else int(float(i))
        rows.append((str(z or "").strip(), i, str(p or "").strip()))
    return rows

def _register_fonts():
    # Embed a Unicode font for Polish characters — avoids Adobe substitution issues.
    # DejaVuSans.ttf is widely available; try to load from system, else bundled copy.
    try_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf",
        os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf"),
    ]
    for p in try_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("DejaVuSans", p))
            return "DejaVuSans"
    # Fallback to Helvetica (may miss diacritics, but we tried).
    return "Helvetica"

def _make_stamp_page(zlecenie: str, ilosc: int, przewoznik: str, page_size=A4) -> bytes:
    """Create a single-page PDF overlay with the annotation text."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=page_size, bottomup=True)
    font_name = _register_fonts()
    c.setAuthor("Kersia PDF Stamper")
    c.setTitle(f"Zlecenie {zlecenie}")
    c.setCreator("Kersia PDF Stamper (Adobe-safe)")

    c.setFont(font_name, 14)
    margin = 15 * mm
    x = margin
    y = page_size[1] - margin

    c.drawString(x, y, f"ZLECENIE: {zlecenie}")
    y -= 8 * mm
    c.drawString(x, y, f"ILOŚĆ PALET: {ilosc}")
    y -= 8 * mm
    c.drawString(x, y, f"PRZEWOŹNIK: {przewoznik}")
    c.showPage()
    c.save()
    return buf.getvalue()

def _normalize_contents(page: PageObject):
    """Ensure /Contents is an array; this helps old Adobe readers."""
    from PyPDF2.generic import ArrayObject
    contents = page.get("/Contents")
    if contents is None:
        return
    if not isinstance(contents, ArrayObject):
        page[page.indirect_ref["/Contents"] if hasattr(page, "indirect_ref") else "/Contents"] = ArrayObject([contents])

def annotate_pdf(pdf_bytes: bytes, excel_bytes: bytes, max_per_sheet: int = 3) -> bytes:
    # Read source PDF and data
    reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)  # be lenient on incoming files
    rows = _parse_excel(excel_bytes)
    if not rows:
        raise ValueError("Nie znaleziono danych w Excelu. Upewnij się, że masz kolumny: ZLECENIE, ILOŚĆ PALET, PRZEWOŹNIK.")

    # Prepare writer
    writer = PdfWriter()
    writer.add_metadata({
        "/Producer": "Kersia PDF Stamper (PyPDF2 + ReportLab)",
        "/Creator": "Kersia PDF Stamper (Adobe-safe)",
        "/Title": "Zlecenia",
    })

    # Iterate pages and stamp in order
    data_idx = 0
    for page_num, page in enumerate(reader.pages):
        # Normalize contents for Acrobat strictness
        _normalize_contents(page)

        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            zlecenie, ilosc, przewoznik = rows[data_idx]
            data_idx += 1

            # Make a one-page overlay PDF and merge onto the current page
            overlay_bytes = _make_stamp_page(zlecenie, ilosc, przewoznik, page.mediabox[2:])
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            overlay_page = overlay_reader.pages[0]

            # Merge using merge_page (safe for Acrobat when contents normalized)
            page.merge_page(overlay_page)

        # After stamping, add to writer as a fresh page object
        writer.add_page(page)

    # If we still have more rows than pages, continue stamping on copies of the last page size
    while data_idx < len(rows):
        base_size = reader.pages[-1].mediabox[2:]
        # Create a blank page to carry more stamps
        blank = PageObject.create_blank_page(width=float(base_size[0]), height=float(base_size[1]))
        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            zlecenie, ilosc, przewoznik = rows[data_idx]
            data_idx += 1
            overlay_bytes = _make_stamp_page(zlecenie, ilosc, przewoznik, base_size)
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            overlay_page = overlay_reader.pages[0]
            blank.merge_page(overlay_page)
        writer.add_page(blank)

    out = io.BytesIO()
    writer.write(out)  # write fresh file (no incremental update) for Acrobat compatibility
    return out.getvalue()

# ---- UI ----
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