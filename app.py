import io, re, base64
from datetime import datetime
from typing import List

import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader, PdfWriter, Transformation
from PyPDF2._page import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from pdfminer.high_level import extract_text

# ------------- Branding -------------
st.set_page_config(page_title="Kersia PDF Stamper", page_icon="🧰", layout="centered")
st.image(base64.b64decode("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sLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIyJsLAyDMMwjImwsDIMwzCMibCwMgzDMIxpEP0/iSXy4mlODeAAAAAASUVORK5CYII="), width=240)
st.title("Kersia — PDF Stamper (web)")
st.caption("Wrzuc Excel + PDF, ustaw limit stron na kartkę i pobierz gotowy PDF.")

# ------------- Core (equal width + adaptive crop) -------------
SIDE_MARGIN_MM = 2
TOP_MARGIN_MM = 4
STAMP_BOTTOM_MM = 12
INTER_GAP_MM = 1

BASE_CROP_L = 6
BASE_CROP_R = 6
BASE_CROP_T = 8
BASE_CROP_B = 8

LOW_TEXT_LINES = 4
SHORT_TEXT_CHARS = 80
EXTRA_CROP_LR = 14
EXTRA_CROP_T  = 18
EXTRA_CROP_B  = 28

def strip_diacritics(s: str) -> str:
    import unicodedata
    if s is None: return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if ord(c) < 128)

def read_excel_lookup(file_like):
    wb = load_workbook(file_like, data_only=True)
    ws = wb.active
    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v is None: continue
        headers[str(v).strip().lower()] = col
    z_col = headers.get("zlecenie")
    ilo_col = headers.get("ilośc palet") or headers.get("ilosc palet") or headers.get("ilość palet")
    pr_col = headers.get("przewoźnik") or headers.get("przewoznik")
    if not z_col or not ilo_col or not pr_col:
        raise ValueError("Excel musi mieć kolumny: ZLECENIE, ilość palet, przewoźnik (nagłówki w 1. wierszu).")
    lookup = {}
    all_nums = set()
    for row in range(2, ws.max_row + 1):
        z = ws.cell(row=row, column=z_col).value
        il = ws.cell(row=row, column=ilo_col).value
        pr = ws.cell(row=row, column=pr_col).value
        z = "" if z is None else str(z).strip()
        il = "" if il is None else str(il).strip()
        pr = "" if pr is None else str(pr).strip()
        parts = [p.strip() for p in re.split(r"[+;,/\s]+", z) if p.strip()]
        for p in parts:
            p2 = "".join(ch for ch in p if ch.isdigit())
            if p2.isdigit():
                all_nums.add(p2)
                lookup[p2] = (z, il, pr)
    return lookup, all_nums

NBSP = "\u00A0"; NNBSP = "\u202F"; THINSP = "\u2009"
def normalize_digits(s: str) -> str:
    return re.sub(r"[\s\-{}{}{}]".format(NBSP, NNBSP, THINSP), "", s)

def extract_candidates(text: str) -> List[str]:
    normal = re.findall(r"\b\d{4,8}\b", text)
    fancy = re.findall(r"(?<!\d)(?:\d[\s\u00A0\u202F\u2009\-]?){4,9}(?!\d)", text)
    fancy = [normalize_digits(s) for s in fancy]
    so = [normalize_digits(m.group(1)) for m in re.finditer(r"Sales\s*[\r\n ]*Order[\s:]*([0-9\s\u00A0\u202F\u2009\-]{4,12})", text, flags=re.I)]
    cands = normal + fancy + so
    cands = [c for c in cands if c.isdigit() and 4 <= len(c) <= 8]
    out, seen = [], set()
    for c in cands:
        if c not in seen:
            out.append(c); seen.add(c)
    return out

def make_blank_page_bytes(width: float, height: float) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    c.showPage(); c.save()
    return buf.getvalue()

def make_stamp_overlay_bytes(width: float, height: float, header: str, footer: str, font_size: int = 12, margin_mm: int = 8) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    try: c.setFont("Helvetica-Bold", font_size)
    except Exception: c.setFont("Helvetica", font_size)
    margin = margin_mm * mm
    c.drawRightString(width - margin, margin + font_size + 1, header)
    if footer:
        c.drawRightString(width - margin, margin, footer)
    c.save()
    return buf.getvalue()

def chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def adaptive_crop_extra(text: str):
    lines = [ln for ln in (text or "").splitlines() if ln.strip()]
    sparse = (len(lines) <= 4) or (len((text or "")) < 80)
    if sparse:
        from reportlab.lib.units import mm as _mm
        return (14*_mm, 14*_mm, 18*_mm, 28*_mm)
    return (0,0,0,0)

def annotate_pdf(pdf_bytes: bytes, xlsx_bytes: bytes, max_per_sheet: int) -> bytes:
    lookup, excel_numbers = read_excel_lookup(io.BytesIO(xlsx_bytes))
    reader = PdfReader(io.BytesIO(pdf_bytes))

    groups, page_meta, page_text_cache = {}, {}, {}
    for i, _ in enumerate(reader.pages):
        page_text = extract_text(io.BytesIO(pdf_bytes), page_numbers=[i]) or ""
        page_text_cache[i] = page_text
        cands = extract_candidates(page_text)
        picked = next((n for n in cands if n in excel_numbers), None)
        mapped = lookup.get(picked) if picked else None

        if mapped:
            z_full, il, pr = mapped
            key = z_full
            header = f"ZLECENIA (laczone): {{strip_diacritics(z_full)}}" if "+" in z_full else f"ZLECENIE: {{strip_diacritics(z_full)}}"
            footer = f"ilosc palet: {{strip_diacritics(il)}} | przewoznik: {{strip_diacritics(pr)}}"
        elif picked:
            key = picked
            header = f"ZLECENIE: {{picked}}"
            footer = "(brak danych w Excelu)"
        else:
            key = f"_NO_ORDER_{{i+1}}"
            header = "(nie znaleziono numeru zlecenia na tej stronie)"
            footer = ""

        groups.setdefault(key, []).append(i)
        page_meta[i] = (header, footer)

    def key_sort(k: str):
        nums = [int(x) for x in re.findall(r"\d+", k)]
        return (min(nums) if nums else 10**9, k)
    ordered_keys = sorted(groups.keys(), key=key_sort)

    writer = PdfWriter()
    W, H = A4

    margin_x = SIDE_MARGIN_MM * mm
    top_margin = TOP_MARGIN_MM * mm
    bottom_for_stamp = STAMP_BOTTOM_MM * mm
    gap = INTER_GAP_MM * mm

    avail_w = W - 2 * margin_x
    avail_h = H - top_margin - bottom_for_stamp

    base_crop_l = BASE_CROP_L * mm
    base_crop_r = BASE_CROP_R * mm
    base_crop_t = BASE_CROP_T * mm
    base_crop_b = BASE_CROP_B * mm

    def make_blank():
        return PdfReader(io.BytesIO(make_blank_page_bytes(W, H))).pages[0]

    for gkey in ordered_keys:
        idxs = groups[gkey]
        for start in range(0, len(idxs), max_per_sheet):
            batch = idxs[start:start+max_per_sheet]
            items, total_h = [], 0.0
            for idx in batch:
                src = reader.pages[idx]
                sw = float(src.mediabox.right - src.mediabox.left)
                sh = float(src.mediabox.top - src.mediabox.bottom)
                ex_l, ex_r, ex_t, ex_b = adaptive_crop_extra(page_text_cache[idx])
                cl = base_crop_l + ex_l
                cr = base_crop_r + ex_r
                ct = base_crop_t + ex_t
                cb = base_crop_b + ex_b
                cw = max(10.0, sw - cl - cr)
                ch = max(10.0, sh - ct - cb)
                s = avail_w / cw
                dh = s * ch
                items.append((idx, cl, cr, ct, cb, s, dh))
                total_h += dh
            total_h += gap * max(0, len(batch)-1)
            down = min(1.0, avail_h / total_h) if total_h > 0 else 1.0
            page = make_blank()
            y = H - top_margin
            from PyPDF2._page import PageObject
            for (idx, cl, cr, ct, cb, s, dh) in items:
                s *= down; dh *= down
                x = margin_x - s * cl
                y2 = y - dh
                tmp = PageObject.create_blank_page(width=W, height=H)
                tmp.merge_page(reader.pages[idx])
                T = (Transformation().translate(-cl, -cb).scale(s, s).translate(x, y2))
                tmp.add_transformation(T)
                page.merge_page(tmp)
                y = y2 - gap

            header, footer = page_meta[batch[0]]
            overlay = PdfReader(io.BytesIO(make_stamp_overlay_bytes(W, H, header, footer)))
            page.merge_page(overlay.pages[0])
            writer.add_page(page)

    buf = io.BytesIO(); writer.write(buf)
    return buf.getvalue()

# ----------- UI -----------
excel_file = st.file_uploader("Plik Excel (ZLECENIE, ilość palet, przewoźnik):", type=["xlsx", "xlsm", "xls"])
pdf_file   = st.file_uploader("Plik PDF:", type=["pdf"])
max_per_sheet = st.slider("Maks. stron na kartkę", 1, 6, 3, 1)

if st.button("GENERUJ PDF", type="primary", disabled=not (excel_file and pdf_file)):
    try:
        result = annotate_pdf(pdf_file.read(), excel_file.read(), max_per_sheet)
        fname = f"zlecenia_{{datetime.now().strftime('%Y%m%d')}}.pdf"
        st.success("Gotowe! Pobierz poniżej.")
        st.download_button("Pobierz wynik", data=result, file_name=fname, mime="application/pdf")
    except Exception as e:
        st.error(f"Błąd: {{e}}")
