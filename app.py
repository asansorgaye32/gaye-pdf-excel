import io
import os
import re
import shutil
import tempfile
import zipfile
from lxml import etree as ET
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import fitz
import pytesseract
from PIL import Image
from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "assets" / "2020 ├ûRNEK PROFORMA HAZIRLAYANLI 4.08.2025(1).xlsm"
LOOKUP_PATH = APP_DIR / "assets" / "2020_PROFORMA_OTOMATIK_KODLU.xlsm"
MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
MAX_PDF_BYTES = 20 * 1024 * 1024
TURKEY_TZ = ZoneInfo("Europe/Istanbul")

app = FastAPI(title="MMO / Royal Cert PDF → Excel API", version="1.0.0")


def require_token(token: str | None):
    expected = os.getenv("APP_API_TOKEN", "").strip()
    if not expected:
        raise HTTPException(status_code=503, detail="Sunucu APP_API_TOKEN ayarı yapılmamış.")
    if not token or token != expected:
        raise HTTPException(status_code=401, detail="Yetkisiz istek. Uygulama token'ını kontrol edin.")


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract embedded text; if pages are scanned, OCR them with Turkish language data."""
    chunks = []
    try:
        with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
            for page in doc:
                chunks.append(page.get_text("text") or "")
            text = "\n".join(chunks).strip()
            if len(re.sub(r"\s+", "", text)) >= 100:
                return text

            # OCR fallback for image-only or nearly empty PDFs.
            ocr_chunks = []
            for page in doc:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.open(io.BytesIO(pix.tobytes("png")))
                try:
                    ocr_chunks.append(pytesseract.image_to_string(image, lang="tur+eng"))
                except Exception:
                    ocr_chunks.append(pytesseract.image_to_string(image, lang="eng"))
            ocr_text = "\n".join(ocr_chunks).strip()
            if len(ocr_text) > len(text):
                text = ocr_text
            return text
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"PDF okunamadı: {exc}")


def _first(patterns, text, flags=re.IGNORECASE | re.MULTILINE):
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return re.sub(r"\s+", " ", match.group(1)).strip(" :-\t\r\n")
    return ""


def _section_for_line(line: str) -> str:
    s = line.casefold()
    if any(x in s for x in ("güvensiz", "guvensiz", "kırmızı", "kirmizi", "red")):
        return "red"
    if any(x in s for x in ("kusurlu", "sarı", "sari", "yellow")):
        return "yellow"
    if any(x in s for x in ("hafif kusurlu", "mavi", "blue")):
        return "blue"
    if any(x in s for x in ("uygun", "yeşil", "yesil", "green")):
        return "green"
    return "unknown"


def analyse_without_ai(text: str, report_type_hint: str) -> dict:
    """Deterministic parser for common MMO/Royal Cert text reports; no paid AI API."""
    lines = [re.sub(r"\s+", " ", x).strip() for x in text.splitlines() if x.strip()]
    flat = "\n".join(lines)
    building_raw = _first([
        r"BİNA SORUMLUSUNA İLİŞKİN BİLGİLER[\s\S]{0,500}?ADRES\s*[:：]\s*([^\n]+)",
        r"ADRES\s*[:：]\s*([^\n]+)",
        r"(?:Bina\s*(?:sorumlusu\s*)?olan\s*adres|Bina\s*Adı\s*/\s*Adres|Bina\s*Adı\s*[-:]?\s*Adres)\s*[:：]?\s*([^\n]+)",
        r"(?:Bina\s*adı\s*ve\s*adres)\s*[:：]?\s*([^\n]+)",
    ], flat)
    manager = _first([
    r"Bina\s*sorumlusuna\s*ilişkin\s*bilgiler[\s\S]{0,500}?ADI\s*VE\s*SOYADI\s*[:：]\s*([^\n]+)",
    r"Bina\s*sorumlusu\s*(?:Adı\s*soyadı|adı\s*soyadı|adı)\s*[:：]?\s*([^\n]+)",
], flat)
    phone = _first([
    r"BİNA SORUMLUSUNA İLİŞKİN BİLGİLER[\s\S]{0,800}?TELEFON NUMARASI\s*[:：]\s*([+()\d\s-]{7,})",
    r"TELEFON\s*NUMARASI\s*[:：]\s*([+()\d\s-]{7,})",
    r"Bina\s*sorumlusu\s*telefon(?:u)?\s*[:：]?\s*([+()\d\s-]{7,})",
], flat)
    followup = _first([r"Periyodik\s*takip\s*kontrol\s*sonucu\s*değerlendirilmesi[^\n]{0,100}?([0-3]?\d[./-][01]?\d[./-](?:20)?\d{2})", r"(?:Bir\s*sonraki\s*)?(?:periyodik|takip)\s*kontrol\s*tarihi\s*[:：]?\s*([0-3]?\d[./-][01]?\d[./-](?:20)?\d{2})"], flat)
    label = _first([r"İstenilen\s*etiket(?:i|i\s*rengi)?\s*[:：]?\s*([^\n]+)"], flat)
    current_label = _first([r"(?:Kontrol\s*sonucu|Mevcut\s*etiket|Kontrol\s*etiketi)\s*[:：]?\s*([^\n]+)"], flat)
    year = _first([r"(?:Montaj|Tesis|İmalat)\s*yılı\s*[:：]?\s*(\d{4})"], flat)
    stops = _first([r"Durak\s*sayısı\s*[:：]?\s*(\d+)", r"Kat\s*ve\s*Durak\s*sayısı\s*[:：]?\s*[^\n]*?(\d+)\s*durak"], flat)
    person = _first([r"Beyan\s*yükü\s*\(?(?:kg)?\)?\s*[:：]?\s*(\d+)\s*(?:kişi|kişilik)", r"(?:Kişi\s*sayısı|Kapasite)\s*[:：]?\s*(\d+)\s*(?:kişi)?"], flat)
    capacity = _first([r"Beyan\s*yükü[^\n:]*[:：]?\s*(?:\d+\s*(?:kişi|kişilik)?\s*[/,-]?\s*)?(\d{2,5})\s*kg", r"Anma\s*yükü\s*[:：]?\s*(\d{2,5})\s*kg"], flat)
    standard = _first([r"(TS\s*EN\s*81\s*[-.]\s*(?:1|20|80)(?:\s*\+\s*A3)?)"], flat)

    # Infer report type only from visible terms.
    low = flat.casefold()
    report_type = "mmo" if "makina mühendisleri odası" in low or "mmo" in low else ("royal_cert" if "royal cert" in low or "royalcert" in low else "other")
    deficiencies = []
    seen = set()
    section = "unknown"
    for line in lines:
        detected = _section_for_line(line)
        if detected != "unknown":
            section = detected
        # Accept codes only on lines that look like a finding/defect row, avoiding standards references.
        if re.search(r"TS\s*EN|madde\s+\d|standart", line, re.I):
            continue
        codes = re.findall(r"(?<!\d)(\d+(?:\.\d+)+(?:-\*{1,2})?)(?!\d)", line)
        if not codes:
            continue
        looks_like_finding = bool(re.search(r"kusur|uygunsuz|eksik|güven|kabin|kapı|fren|durak|etiket|kontrol|madde|proforma|\b[123]\.\d", line, re.I))
        if not looks_like_finding:
            continue
        for code in codes:
            code = normalize_code(code)
            if code not in seen:
                seen.add(code)
                deficiencies.append({"code": code, "section": section})

    return {
        "report_type": report_type,
        "building_name": building_raw.split("-", 1)[0].strip() if "-" in building_raw else building_raw,
        "raw_address": building_raw,
        "address": building_raw.split("-", 1)[1].strip() if "-" in building_raw else "",
        "manager": manager,
        "phone": phone,
        "followup_date": followup,
        "requested_label": label,
        "current_label": current_label,
        "installation_year": year,
        "floor_count": stops,
        "person_count": person,
        "capacity_kg": capacity,
        "standard": standard,
        "deficiencies": deficiencies,
    }

def normalize_code(code: str) -> str:
    return re.sub(r"\s+", "", str(code or "")).strip(".,;:()[]")


def normalize_standard(raw: str) -> str:
    s = (raw or "").upper().replace("–", "-").replace(" ", "")
    # User rule: 81.1 / 81.80 are represented as 81.1; then A3; then 81.20.
    if "81-1+A3" in s or "81.1+A3" in s or "81.A3" in s or "81-A3" in s:
        return "TS EN 81.a3"
    if "81-20" in s or "81.20" in s:
        return "TS EN 81.20"
    if "81-80" in s or "81.80" in s or "81-1" in s or "81.1" in s:
        return "TS EN 81.1"
    return (raw or "").strip()

def normalise_label(value: str) -> str:
    s = (value or "").strip().upper()
    if "KIRMIZI" in s or "RED" in s or "GÜVENSİZ" in s or "GUVENSIZ" in s:
        return "KIRMIZI"
    if "MAVİ" in s or "MAVI" in s or "BLUE" in s or "HAFİF" in s or "HAFIF" in s:
        return "MAVİ"
    if "SARI" in s or "YELLOW" in s or "KUSURLU" in s:
        return "SARI"
    if "YEŞİL" in s or "YESIL" in s or "GREEN" in s or "UYGUN" in s:
        return "YEŞİL"
    return ""


def find_codes_in_source(text: str) -> set[str]:
    # This is only a sanity check against AI-invented codes; AI decides which occurrences
    # are actual deficiency list entries rather than references inside explanatory text.
    return {normalize_code(m.group(0)) for m in re.finditer(r"(?<!\d)\d+(?:\.\d+)+(?:-\*{1,2})?(?!\d)", text)}


# The workbook is edited at the OOXML package level instead of being saved with
# openpyxl. This leaves the original GAYE Asansör logo, VBA project, VML, drawings,
# relationships, custom extensions and other Excel package parts untouched.
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def get_lookup() -> dict[str, str]:
    if not LOOKUP_PATH.exists():
        raise HTTPException(status_code=500, detail="Proforma kod eşleştirme Excel'i bulunamadı.")
    wb = load_workbook(LOOKUP_PATH, read_only=True, data_only=True, keep_vba=True)
    try:
        if "MADDE_LISTESI" not in wb.sheetnames:
            raise HTTPException(status_code=500, detail="Otomatik kodlu Excel'de MADDE_LISTESI sayfası yok.")
        sheet = wb["MADDE_LISTESI"]
        result = {}
        for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
            code, description = row
            if code is not None:
                result[normalize_code(str(code))] = str(description or "").strip()
        return result
    finally:
        wb.close()


def qname(tag: str) -> str:
    return "{" + NS["m"] + "}" + tag


_CELL_CACHE = {}
_ROW_CACHE = {}
_SHEET_DATA = None


def index_sheet(root):
    global _CELL_CACHE, _ROW_CACHE, _SHEET_DATA
    _SHEET_DATA = root.find("m:sheetData", NS)
    _ROW_CACHE = {int(row.attrib["r"]): row for row in _SHEET_DATA.findall("m:row", NS)}
    _CELL_CACHE = {cell.attrib["r"]: cell for row in _SHEET_DATA.findall("m:row", NS) for cell in row.findall("m:c", NS)}


def get_cell(root, address: str):
    global _CELL_CACHE, _ROW_CACHE, _SHEET_DATA
    cell = _CELL_CACHE.get(address)
    if cell is not None:
        return cell
    row_num = int(re.search(r"\d+", address).group())
    row = _ROW_CACHE.get(row_num)
    if row is None:
        row = ET.Element(qname("row"), {"r": str(row_num)})
        inserted = False
        for i, existing in enumerate(list(_SHEET_DATA)):
            if int(existing.attrib.get("r", "0")) > row_num:
                _SHEET_DATA.insert(i, row); inserted = True; break
        if not inserted: _SHEET_DATA.append(row)
        _ROW_CACHE[row_num] = row
    cell = ET.Element(qname("c"), {"r": address})
    col_letters = re.match(r"[A-Z]+", address).group()
    col_num = 0
    for ch in col_letters: col_num = col_num * 26 + ord(ch) - 64
    inserted = False
    for i, existing in enumerate(row.findall("m:c", NS)):
        ex_col = re.match(r"[A-Z]+", existing.attrib.get("r", "A")).group()
        ex_num = 0
        for ch in ex_col: ex_num = ex_num * 26 + ord(ch) - 64
        if ex_num > col_num:
            row.insert(row.index(existing), cell); inserted = True; break
    if not inserted: row.append(cell)
    _CELL_CACHE[address] = cell
    return cell


def set_inline(root, address: str, value):
    cell = get_cell(root, address)
    for child in list(cell):
        if child.tag in (qname("v"), qname("f"), qname("is")):
            cell.remove(child)
    if value is None or str(value) == "":
        cell.attrib.pop("t", None)
        return
    cell.set("t", "inlineStr")
    inline = ET.SubElement(cell, qname("is"))
    text_el = ET.SubElement(inline, qname("t"))
    text_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_el.text = str(value)


def set_formula(root, address: str, formula: str):
    cell = get_cell(root, address)
    for child in list(cell):
        if child.tag in (qname("v"), qname("f"), qname("is")):
            cell.remove(child)
    cell.attrib.pop("t", None)
    f = ET.SubElement(cell, qname("f")); f.text = formula
    ET.SubElement(cell, qname("v"))


def copy_cell_style(root, source: str, target: str):
    src = get_cell(root, source)
    dst = get_cell(root, target)
    if "s" in src.attrib: dst.set("s", src.attrib["s"])
    else: dst.attrib.pop("s", None)


def split_building_address(data: dict):
    raw = (data.get("raw_address") or data.get("building_name") or data.get("address") or "").strip()
    # The expected format is e.g. "Özpetek A Blok - [address]".
    for sep in (" - ", " – ", " — "):
        if sep in raw:
            left, right = raw.split(sep, 1)
            if left.strip() and right.strip():
                data["building_name"] = left.strip()
                data["address"] = right.strip()
                return
    # Also handle hyphen without surrounding spaces if it is clearly a separator.
    match = re.match(r"^(.+?)\s*-\s*(.+)$", raw)
    if match:
        data["building_name"], data["address"] = match.group(1).strip(), match.group(2).strip()
    else:
        data["building_name"] = (data.get("building_name") or raw).strip()
        data["address"] = (data.get("address") or "").strip()


def fill_workbook(data: dict, source_text: str) -> bytes:
    if not TEMPLATE_PATH.exists():
        raise HTTPException(
            status_code=500,
            detail=(
                f"Şablon bulunamadı. Aranan yol: {TEMPLATE_PATH} | "
                f"assets klasörü mevcut: {(APP_DIR / 'assets').exists()}"
            ),
        )
    try:
        split_building_address(data)
        lookup = get_lookup()
        source_codes = find_codes_in_source(source_text)
        deficiencies, seen = [], set()
        for item in data.get("deficiencies", []):
            code = normalize_code(item.get("code", ""))
            section = item.get("section", "unknown")
            if not code or code not in source_codes or code in seen:
                continue
            if section not in ("red", "yellow"):
                continue
            seen.add(code)
            deficiencies.append({"code": code, "section": section})

        # Current label is based on whichever red/yellow deficiency appears first.
        label = ""
        for item in deficiencies:
            if item["section"] == "red":
                label = "KIRMIZI"; break
            if item["section"] == "yellow":
                label = "SARI"; break
        if not label:
            label = normalise_label(data.get("current_label", ""))

        today = datetime.now(TURKEY_TZ).replace(tzinfo=None).strftime("%d.%m.%Y")
        with zipfile.ZipFile(TEMPLATE_PATH, "r") as zin:
            package = {name: zin.read(name) for name in zin.namelist()}
        sheet_path = "xl/worksheets/sheet1.xml"
        if sheet_path not in package:
            raise ValueError("Excel şablonunun ana sayfası bulunamadı.")
        root = ET.fromstring(package[sheet_path])
        index_sheet(root)

        # Header: only cell contents are changed; original styles and all other
        # package parts (including logo image and VBA) remain byte-for-byte intact.
        headers = {
            "C2": data.get("building_name") or "",
            "C3": data.get("address") or "",
            "C4": data.get("manager") or "",
            "C5": "GAYE",
            "C6": data.get("phone") or "",
            "C7": "Mehmet GÖKDEMİR",
            "I2": today,
            "I3": "",  # Teklif No intentionally blank.
            "I4": "X",
            "F4": data.get("followup_date") or "",
            "F5": normalise_label(data.get("requested_label", "")) or "MAVİ",
            "I5": data.get("installation_year") or "",
            "F6": data.get("floor_count") or "",
            "I6": normalize_standard(data.get("standard", "")),
            "F7": " / ".join(x for x in [(data.get("person_count") or "").strip(), (data.get("capacity_kg") or "").strip() + " Kg" if (data.get("capacity_kg") or "").strip() else ""] if x),
            "I7": label,
        }
        for address, value in headers.items(): set_inline(root, address, value)

        # Keep the template's preformatted rows. Clear only the item value cells.
        # Never insert/delete rows, which is a common source of workbook repairs.
        for address, cell in list(_CELL_CACHE.items()):
            match = re.match(r"([A-Z]+)(\d+)$", address)
            if not match or not (9 <= int(match.group(2)) <= 3624) or match.group(1) not in ("A", "B", "C", "D", "E", "G", "I"):
                continue
            for child in list(cell):
                if child.tag in (qname("v"), qname("f"), qname("is")):
                    cell.remove(child)
            cell.attrib.pop("t", None)

        first_row = 9
        for idx, item in enumerate(deficiencies):
            row = first_row + idx
            set_inline(root, f"A{row}", "1")
            set_inline(root, f"C{row}", lookup.get(item["code"], "Proforma yok"))
            set_inline(root, f"G{row}", "0")
            set_formula(root, f"I{row}", f"A{row}*G{row}")

        labour_row = first_row + len(deficiencies)
        set_inline(root, f"A{labour_row}", "1")
        set_inline(root, f"C{labour_row}", "İşçilik")
        set_inline(root, f"G{labour_row}", "0")
        set_formula(root, f"I{labour_row}", f"A{labour_row}*G{labour_row}")

        total_row = labour_row + 1
        # Reuse original total-row styles for the moving totals area.
        for offset, source_row in enumerate((12, 13, 14)):
            dest = total_row + offset
            for col in ("G", "I"):
                copy_cell_style(root, f"{col}{source_row}", f"{col}{dest}")
        set_inline(root, f"G{total_row}", "TOPLAM :")
        set_formula(root, f"I{total_row}", f"SUM(I{first_row}:I{labour_row})")
        set_inline(root, f"G{total_row+1}", "%20 KDV :")
        set_formula(root, f"I{total_row+1}", f"I{total_row}*20%")
        set_inline(root, f"G{total_row+2}", "G.TOPLAM :")
        set_formula(root, f"I{total_row+2}", f"I{total_row}+I{total_row+1}")

        package[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        # Request a full recalc on open, without changing any other workbook parts.
        wb_root = ET.fromstring(package["xl/workbook.xml"])
        calc = wb_root.find("m:calcPr", NS)
        if calc is None:
            calc = ET.SubElement(wb_root, qname("calcPr"))
        calc.set("calcMode", "auto"); calc.set("fullCalcOnLoad", "1"); calc.set("forceFullCalc", "1")
        package["xl/workbook.xml"] = ET.tostring(wb_root, encoding="utf-8", xml_declaration=True)
        out = io.BytesIO()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, content in package.items(): zout.writestr(name, content)
        return out.getvalue()
    except HTTPException:
        raise
    except Exception as exc:
        print("EXCEL HATASI:", repr(exc), flush=True)
        raise HTTPException(status_code=500, detail=f"Excel oluşturulamadı: {exc}")


@app.get("/health")
def health():
    return {"status": "ok", "service": "mmo-royalcert-excel"}


@app.post("/generate-excel")
async def generate_excel(
    pdf: UploadFile = File(...),
    reportType: str = Form("auto"),
    x_app_token: str | None = Header(default=None, alias="X-App-Token"),
):
    require_token(x_app_token)
    if not pdf.filename or not pdf.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Lütfen PDF dosyası gönderin.")
    pdf_bytes = await pdf.read()
    if not pdf_bytes:
        raise HTTPException(status_code=400, detail="PDF dosyası boş.")
    if len(pdf_bytes) > MAX_PDF_BYTES:
        raise HTTPException(status_code=413, detail="PDF en fazla 20 MB olabilir.")
    source_text = extract_pdf_text(pdf_bytes)
    if len(source_text.strip()) < 50:
        raise HTTPException(status_code=422, detail="PDF'den okunabilir metin çıkarılamadı. Daha net PDF deneyin.")
    data = analyse_without_ai(source_text, reportType)
    print("PDF ANALIZI TAMAMLANDI:", data, flush=True)
    excel_bytes = fill_workbook(data, source_text)
    print("EXCEL OLUSTURULDU, BYTE SAYISI:", len(excel_bytes), flush=True)
    base_name = Path(pdf.filename).stem
    safe = re.sub(r"[^A-Za-z0-9._-]+", "_", base_name)[:70] or "rapor"
    filename = f"{safe}_doldurulmus.xlsm"
    return StreamingResponse(
        io.BytesIO(excel_bytes),
        media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
