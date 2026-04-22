"""
add_page.py - Bổ sung số trang vào trích nguồn trong file Excel
=================================================================
Chiến lược tìm trang (3 tầng):
  1. Exact match  - tìm chuỗi chính xác trong text đã extract từ PDF
  2. Fuzzy match  - tìm theo keyword dài nhất trích từ đoạn văn
  3. AI fallback  - gọi Vertex AI (Gemini) khi 2 cách trên thất bại

Hỗ trợ PDF scan (ảnh): tự động OCR bằng Tesseract nếu PDF không có text layer.
Yêu cầu OCR: pip install pytesseract pdf2image  +  tesseract-ocr-vie

Cách dùng:
  python add_page.py --excel data.xlsx --pdf book.pdf
  python add_page.py --excel data.xlsx --pdf book.pdf --workers 8
  python add_page.py --excel data.xlsx --pdf book.pdf --no-ai      # chỉ dùng text search
  python add_page.py --excel data.xlsx --pdf book.pdf --ocr-cache  # lưu cache OCR để tái dùng
"""

import os
import sys
import re
import time
import json
import argparse
import unicodedata
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# ── openpyxl (làm đẹp) ───────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    os.system("pip install openpyxl --quiet")
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

# ── pypdf (đọc PDF) ───────────────────────────────────────────────────────────
try:
    from pypdf import PdfReader
except ImportError:
    os.system("pip install pypdf --quiet")
    from pypdf import PdfReader

# ── OCR dependencies (tuỳ chọn) ──────────────────────────────────────────────
try:
    import pytesseract
    from pdf2image import convert_from_path
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# ── callAPI (Vertex AI) ───────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(BASE_DIR)

from dotenv import load_dotenv
load_dotenv(os.path.join(BASE_DIR, 'API', '.env'))

try:
    from API.callAPI import VertexClient, get_vertex_ai_credentials
    HAS_VERTEX = True
except ImportError:
    print("⚠️  Không tìm thấy callAPI.py — sẽ chỉ dùng text search (--no-ai).")
    HAS_VERTEX = False


# ═══════════════════════════════════════════════════════════════════════════════
# 0. OCR CHO PDF SCAN
# ═══════════════════════════════════════════════════════════════════════════════

def get_poppler_path() -> str | None:
    """
    Tìm Poppler trên Windows:
    1. Kiểm tra PATH hệ thống
    2. Kiểm tra thư mục poppler-* cạnh file script này
    3. Tự động tải về nếu không tìm thấy
    Trả về đường dẫn thư mục bin/ của Poppler, hoặc None (Linux/Mac không cần).
    """
    if sys.platform != "win32":
        return None  # Linux/Mac: poppler trong PATH hệ thống là đủ

    import shutil
    # Kiểm tra PATH hệ thống trước
    if shutil.which("pdftoppm"):
        return None  # Đã có trong PATH, không cần chỉ định

    # Tìm thư mục poppler-* cạnh script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for entry in os.listdir(script_dir):
        candidate = os.path.join(script_dir, entry, "Library", "bin")
        if os.path.isdir(candidate) and os.path.exists(os.path.join(candidate, "pdftoppm.exe")):
            print(f"   ✅ Tìm thấy Poppler: {candidate}")
            return candidate
        candidate2 = os.path.join(script_dir, entry, "bin")
        if os.path.isdir(candidate2) and os.path.exists(os.path.join(candidate2, "pdftoppm.exe")):
            print(f"   ✅ Tìm thấy Poppler: {candidate2}")
            return candidate2

    # Tự động tải Poppler về
    print("   📥 Không tìm thấy Poppler — đang tự động tải về...")
    import urllib.request, zipfile

    # Poppler Windows binary (phiên bản ổn định từ github.com/oschwartz10612/poppler-windows)
    url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip"
    zip_path = os.path.join(script_dir, "poppler_win.zip")
    extract_dir = os.path.join(script_dir, "poppler_win")

    try:
        print(f"   Đang tải: {url}")
        urllib.request.urlretrieve(url, zip_path)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)
        os.remove(zip_path)

        # Tìm thư mục bin trong extracted
        for root, dirs, files in os.walk(extract_dir):
            if "pdftoppm.exe" in files:
                print(f"   ✅ Đã cài Poppler tại: {root}")
                return root
    except Exception as e:
        print(f"   ❌ Tự động tải Poppler thất bại: {e}")
        print()
        print("   👉 Cài thủ công:")
        print("      1. Tải file zip tại: https://github.com/oschwartz10612/poppler-windows/releases")
        print("      2. Giải nén vào cùng thư mục với add_page.py")
        print("      3. Chạy lại chương trình")
        sys.exit(1)

    print("   ❌ Không tìm thấy pdftoppm.exe trong file tải về.")
    sys.exit(1)


def get_tesseract_path() -> None:
    """Kiểm tra Tesseract và ngôn ngữ Việt, in hướng dẫn nếu thiếu."""
    import shutil
    # if sys.platform == "win32":
    # Các vị trí cài đặt phổ biến của Tesseract trên Windows
    common_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
    ]
    for p in common_paths:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            # Kiểm tra ngôn ngữ vie
            langs = pytesseract.get_languages()
            if "vie" not in langs:
                print("   ⚠️  Tesseract thiếu gói ngôn ngữ tiếng Việt!")
                print("      Tải thêm tại: https://github.com/tesseract-ocr/tessdata/blob/main/vie.traineddata")
                print(f"      Đặt file vie.traineddata vào: {os.path.dirname(p)}\\tessdata\\")
                sys.exit(1)
            return
    # Không tìm thấy
    print("   ❌ Không tìm thấy Tesseract OCR!")
    print("   👉 Tải và cài đặt tại: https://github.com/UB-Mannheim/tesseract/wiki")
    print("      Khi cài: tick chọn 'Additional language data' → Vietnamese")
    sys.exit(1)


def _ocr_one_page(args_tuple):
    """Hàm OCR một trang, dùng được với multiprocessing (phải là top-level)."""
    idx, img_bytes, tess_cmd = args_tuple
    try:
        import pytesseract
        from PIL import Image
        import io
        if tess_cmd:
            pytesseract.pytesseract.tesseract_cmd = tess_cmd
        img = Image.open(io.BytesIO(img_bytes))
        text = pytesseract.image_to_string(img, lang="vie+eng", config="--psm 6")
        return idx, text
    except Exception as e:
        return idx, ""


def ocr_pdf(pdf_path: str, cache_path: str | None = None, dpi: int = 150, workers: int = 0) -> list[str]:
    if cache_path and os.path.exists(cache_path):
        print(f"⚡ Tìm thấy cache OCR: {cache_path} — đang load...")
        with open(cache_path, "r", encoding="utf-8") as f:
            return json.load(f)

    get_tesseract_path()
    tess_cmd = pytesseract.pytesseract.tesseract_cmd if sys.platform == "win32" else None
    poppler_path = get_poppler_path()

    import multiprocessing
    from concurrent.futures import ProcessPoolExecutor, as_completed
    from pypdf import PdfReader
    
    n_workers = workers if workers > 0 else min(multiprocessing.cpu_count(), 8)
    total_pages = len(PdfReader(pdf_path).pages)
    
    print(f"🔍 PDF scan → OCR song song {n_workers} luồng (DPI={dpi})...")
    
    pages_text = [""] * total_pages
    chunk_size = 10 # Xử lý 30 trang mỗi đợt để không tràn RAM
    
    with ProcessPoolExecutor(max_workers=n_workers) as executor:
        for start_page in range(1, total_pages + 1, chunk_size):
            end_page = min(start_page + chunk_size - 1, total_pages)
            print(f"   Đang đọc PDF từ trang {start_page} đến {end_page}...")
            
            # Chỉ extract một phần PDF
            convert_kwargs = dict(dpi=dpi, grayscale=True, first_page=start_page, last_page=end_page)
            if poppler_path:
                convert_kwargs["poppler_path"] = poppler_path
            
            images = convert_from_path(pdf_path, **convert_kwargs)
            
            import io
            tasks = []
            for i, img in enumerate(images):
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                # start_page đếm từ 1, index mảng đếm từ 0
                tasks.append((start_page + i - 1, buf.getvalue(), tess_cmd))
            
            # Chạy OCR cho chunk hiện tại
            futures = {executor.submit(_ocr_one_page, t): t[0] for t in tasks}
            for future in as_completed(futures):
                idx, text = future.result()
                pages_text[idx] = text
                
            print(f"   ✅ Đã OCR xong đến trang {end_page}/{total_pages}")

    if cache_path:
        with open(cache_path, "w", encoding="utf-8") as f:
            json.dump(pages_text, f, ensure_ascii=False)
            
    return pages_text


# ═══════════════════════════════════════════════════════════════════════════════
# 1. XÂY DỰNG PDF INDEX
# ═══════════════════════════════════════════════════════════════════════════════

def normalize(text: str) -> str:
    """Chuẩn hoá unicode, bỏ dấu cách thừa, lowercase."""
    text = unicodedata.normalize("NFC", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def extract_printed_page_number(raw_text: str) -> str | None:
    """
    Lấy số trang in (printed page number) từ text đã extract.
    Thường nằm ở dòng đầu tiên hoặc cuối trang, là số nguyên đơn lẻ.
    Xử lý trường hợp PDF in số đôi như '4747', '9797' (do header/footer bị ghép).
    """
    lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
    for line in lines[:3] + lines[-2:]:           # kiểm tra đầu & cuối trang
        # Trường hợp số đôi: '4747' → '47', '9797' → '97'
        m = re.fullmatch(r"(\d{1,4})\1", line)
        if m:
            return m.group(1)
        # Số đơn thường
        if re.fullmatch(r"\d{1,4}", line):
            return line
    return None


def build_pdf_index(pdf_path: str, ocr_cache: str | None = None, dpi: int = 200, ocr_workers: int = 0) -> list[dict]:
    """
    Trả về list các dict:
      { "pdf_idx": int, "printed_page": str|None, "text": str, "text_norm": str }
    Tự động phát hiện PDF scan và fallback sang OCR.
    """
    print(f"📖 Đang đọc và lập index PDF: {pdf_path}")
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)

    # ── Thử extract text thông thường ─────────────────────────────────────────
    raw_texts = []
    for page in reader.pages:
        raw_texts.append(page.extract_text() or "")

    pages_with_text = sum(1 for t in raw_texts if len(t.strip()) > 30)
    text_ratio = pages_with_text / total_pages if total_pages else 0
    print(f"   → {total_pages} trang tổng | {pages_with_text} trang có text ({text_ratio*100:.0f}%)")

    # ── Nếu < 20% trang có text → PDF scan, cần OCR ───────────────────────────
    if text_ratio < 0.2:
        print("   ⚠️  Phát hiện PDF scan (ít/không có text layer) → chuyển sang OCR...")
        raw_texts = ocr_pdf(pdf_path, cache_path=ocr_cache, dpi=dpi, workers=ocr_workers)

    # ── Build index ────────────────────────────────────────────────────────────
    index = []
    for i, raw in enumerate(raw_texts):
        if not raw.strip():
            continue
        printed = extract_printed_page_number(raw)
        index.append({
            "pdf_idx":      i,
            "printed_page": printed,
            "text":         raw,
            "text_norm":    normalize(raw),
        })

    print(f"   → Đã index {len(index)} trang có nội dung.")
    return index


# ═══════════════════════════════════════════════════════════════════════════════
# 2. TÌM TRANG BẰNG TEXT SEARCH (tầng 1 & 2)
# ═══════════════════════════════════════════════════════════════════════════════

MIN_FRAGMENT = 25          # độ dài tối thiểu của chuỗi dùng để tìm kiếm
FUZZY_MIN_WORDS = 5        # số từ tối thiểu để fuzzy match
FUZZY_HIT_RATIO = 0.55     # tỉ lệ từ khớp tối thiểu


def _best_page_label(page_info: dict) -> str:
    """Trả về nhãn trang: printed_page nếu có, không thì dùng pdf_idx+1."""
    return page_info["printed_page"] or str(page_info["pdf_idx"] + 1)


def find_page_exact(excerpt_norm: str, index: list[dict]) -> str | None:
    """Tìm chuỗi chính xác (sau chuẩn hoá)."""
    # Dùng đoạn dài nhất có thể để tránh false positive
    fragment = excerpt_norm[:300] if len(excerpt_norm) > 300 else excerpt_norm
    if len(fragment) < MIN_FRAGMENT:
        return None
    for page in index:
        if fragment in page["text_norm"]:
            return _best_page_label(page)
    # Thử với đoạn ngắn hơn (50 ký tự đầu)
    short = excerpt_norm[:80]
    if len(short) >= MIN_FRAGMENT:
        for page in index:
            if short in page["text_norm"]:
                return _best_page_label(page)
    return None


def find_page_fuzzy(excerpt: str, index: list[dict]) -> str | None:
    """
    Tìm theo từ khoá: lấy các từ dài (≥4 ký tự) từ đoạn đầu excerpt,
    đếm xem trang nào có số từ khớp cao nhất.
    """
    excerpt_norm = normalize(excerpt)
    words = [w for w in excerpt_norm.split() if len(w) >= 4]
    # Chỉ lấy 20 từ đầu để tăng độ chính xác
    words = words[:20]
    if len(words) < FUZZY_MIN_WORDS:
        return None

    best_page = None
    best_ratio = 0.0
    for page in index:
        hits = sum(1 for w in words if w in page["text_norm"])
        ratio = hits / len(words)
        if ratio > best_ratio:
            best_ratio = ratio
            best_page = page

    if best_ratio >= FUZZY_HIT_RATIO and best_page is not None:
        return _best_page_label(best_page)
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# 3. TÌM TRANG BẰNG AI (tầng 3 - fallback)
# ═══════════════════════════════════════════════════════════════════════════════

MAX_RETRIES = 3


def find_page_ai(client, pdf_path: str, excerpt: str) -> str | None:
    """Gọi Gemini qua Vertex AI để tìm số trang của đoạn trích."""
    prompt = f"""Bạn là trợ lý tra cứu tài liệu độ chính xác cao.
Nhiệm vụ: Tìm số trang in (số in trên sách, không phải số thứ tự trang PDF) chứa đoạn văn bản dưới đây trong file PDF đính kèm.

Đoạn văn cần tìm:
\"\"\"
{excerpt[:600]}
\"\"\"

Chỉ trả về JSON duy nhất, không giải thích:
{{"page": "số_trang_hoặc_null"}}

Nếu không tìm thấy thì trả về: {{"page": null}}
"""
    import json

    for attempt in range(MAX_RETRIES):
        try:
            raw = client.send_data_to_AI(
                prompt=prompt,
                file_paths=[pdf_path],
                temperature=0.0,
                max_output_tokens=128,
                use_file_api=True,
                response_mime_type="application/json",
            )
            text = str(raw).strip()
            # Bóc markdown fence nếu có
            text = re.sub(r"^```json\s*|^```\s*|```$", "", text, flags=re.MULTILINE).strip()
            data = json.loads(text)
            page = data.get("page")
            if page and str(page).lower() not in ("null", "none", ""):
                return str(page)
            return None
        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                time.sleep(2 ** attempt)
            else:
                print(f"    ⚠️  AI lỗi sau {MAX_RETRIES} lần thử: {e}")
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# 4. LOGIC CHÍNH CHO TỪNG Ô TRÍCH DẪN
# ═══════════════════════════════════════════════════════════════════════════════

def update_trich_nguon(current: str, page: str) -> str:
    """
    Chèn 'tr. X' vào chuỗi trích nguồn.
    - Nếu rỗng          → "(tr. X)"
    - Nếu đã có tr.     → không thay đổi
    - Nếu kết thúc bằng ')' → chèn trước dấu ')'
    - Còn lại           → thêm " (tr. X)" vào cuối
    """
    val = str(current).strip()
    if not val or val.lower() == "nan":
        return f"(tr. {page})"
    # Đã có số trang → bỏ qua
    if re.search(r"tr\.\s*\d", val, re.IGNORECASE):
        return val
    if val.endswith(")"):
        return val[:-1] + f", tr. {page})"
    return val + f" (tr. {page})"


def process_row(
    row_idx: int,
    row: pd.Series,
    pdf_index: list[dict],
    client,               # VertexClient hoặc None
    pdf_path: str,
    use_ai: bool,
) -> tuple[int, dict]:
    """Xử lý một dòng, trả về (row_idx, {col_name: new_value})."""
    updates = {}
    stats = {"exact": 0, "fuzzy": 0, "ai": 0, "miss": 0}

    for col_name in row.index:
        if not str(col_name).startswith("Nội dung trích dẫn"):
            continue
        suffix = str(col_name).replace("Nội dung trích dẫn", "").strip()
        trich_nguon_col = f"Trích nguồn {suffix}"
        if trich_nguon_col not in row.index:
            continue

        excerpt = str(row.get(col_name, "")).strip()
        if not excerpt or excerpt.lower() == "nan":
            continue

        excerpt_norm = normalize(excerpt)

        # Tầng 1: exact
        page = find_page_exact(excerpt_norm, pdf_index)
        if page:
            stats["exact"] += 1
        else:
            # Tầng 2: fuzzy
            page = find_page_fuzzy(excerpt, pdf_index)
            if page:
                stats["fuzzy"] += 1
            elif use_ai and client:
                # Tầng 3: AI
                page = find_page_ai(client, pdf_path, excerpt)
                if page:
                    stats["ai"] += 1
                else:
                    stats["miss"] += 1
            else:
                stats["miss"] += 1

        if page:
            current_val = row.get(trich_nguon_col, "")
            updates[trich_nguon_col] = update_trich_nguon(str(current_val), page)

    return row_idx, updates, stats


# ═══════════════════════════════════════════════════════════════════════════════
# 5. LÀM ĐẸP EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def apply_beautiful_format(excel_path: str):
    print("🎨 Đang format Excel...")
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    thin = Border(
        left=Side(border_style="thin", color="CCCCCC"),
        right=Side(border_style="thin", color="CCCCCC"),
        top=Side(border_style="thin", color="CCCCCC"),
        bottom=Side(border_style="thin", color="CCCCCC"),
    )

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin
    ws.row_dimensions[1].height = 36

    for row_idx in range(2, ws.max_row + 1):
        bg = "FFFFFF" if row_idx % 2 == 0 else "F5F7FA"
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin
            is_center = col_idx in (1, 6)
            cell.alignment = Alignment(
                horizontal="center" if is_center else "left",
                vertical="top", wrap_text=True,
            )
            # Cột Trích nguồn (chẵn từ cột 8 trở đi) → in nghiêng
            cell.font = Font(italic=(col_idx > 6 and col_idx % 2 == 0), size=9 if col_idx > 6 else 10)

    fixed = {1: 5, 2: 28, 3: 35, 4: 38, 5: 38, 6: 18}
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        if col_idx in fixed:
            ws.column_dimensions[letter].width = fixed[col_idx]
        else:
            ws.column_dimensions[letter].width = 60 if col_idx % 2 != 0 else 42

    ws.freeze_panes = "A2"
    wb.save(excel_path)
    print("✨ Format xong!")


# ═══════════════════════════════════════════════════════════════════════════════
# 6. MAIN
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="Bổ sung số trang vào trích nguồn Excel")
    parser.add_argument("--excel",     required=True,  help="Đường dẫn file Excel (.xlsx)")
    parser.add_argument("--pdf",       required=True,  help="Đường dẫn file PDF sách")
    parser.add_argument("--workers",   type=int, default=4, help="Số luồng song song (default: 4)")
    parser.add_argument("--no-ai",     action="store_true", help="Bỏ qua AI, chỉ dùng text search")
    parser.add_argument("--ocr-cache", action="store_true",
                        help="Lưu kết quả OCR vào file .json để tái dùng (tiết kiệm thời gian)")
    parser.add_argument("--dpi",       type=int, default=200,
                        help="DPI khi render PDF scan để OCR (default: 200, tăng lên 300 nếu chữ nhỏ)")
    args = parser.parse_args()

    # ── Kiểm tra file ──────────────────────────────────────────────────────────
    for path in (args.excel, args.pdf):
        if not os.path.exists(path):
            print(f"❌ Không tìm thấy file: {path}")
            sys.exit(1)

    use_ai = not args.no_ai and HAS_VERTEX

    # ── Cache path cho OCR ─────────────────────────────────────────────────────
    ocr_cache_path = None
    if args.ocr_cache:
        ocr_cache_path = os.path.splitext(args.pdf)[0] + "_ocr_cache.json"

    # ── Build PDF index ────────────────────────────────────────────────────────
    pdf_index = build_pdf_index(args.pdf, ocr_cache=ocr_cache_path, dpi=args.dpi, ocr_workers=args.workers)
    if not pdf_index:
        print("❌ Không đọc được nội dung từ PDF dù đã thử OCR. Kiểm tra lại file.")
        sys.exit(1)

    # ── Khởi tạo Vertex client ─────────────────────────────────────────────────
    client = None
    if use_ai:
        print("🤖 Khởi tạo Vertex AI client...")
        try:
            creds = get_vertex_ai_credentials()
            client = VertexClient(
                project_id=os.getenv("PROJECT_ID"),
                creds=creds,
                model_name="gemini-3.1-pro-preview",
                region="global",
            )
            print("   ☁️  Upload PDF lên File API để cache...")
            client.upload_files_cached([args.pdf])
        except Exception as e:
            print(f"   ⚠️  Không khởi tạo được AI client: {e}. Chạy chỉ với text search.")
            client = None

    # ── Đọc Excel ─────────────────────────────────────────────────────────────
    print(f"\n📂 Đọc Excel: {args.excel}")
    df = pd.read_excel(args.excel)
    total = len(df)
    print(f"   → {total} dòng dữ liệu.")

    # ── Xử lý đa luồng ────────────────────────────────────────────────────────
    print(f"\n⚡ Bắt đầu xử lý với {args.workers} luồng song song...")
    all_updates = {}
    total_stats = {"exact": 0, "fuzzy": 0, "ai": 0, "miss": 0}

    with ThreadPoolExecutor(max_workers=args.workers) as executor:
        futures = {
            executor.submit(process_row, idx, row, pdf_index, client, args.pdf, use_ai): idx
            for idx, row in df.iterrows()
        }
        done = 0
        for future in as_completed(futures):
            done += 1
            orig_idx = futures[future]
            try:
                row_idx, updates, stats = future.result()
                all_updates[row_idx] = updates
                for k in total_stats:
                    total_stats[k] += stats[k]
                found = len(updates)
                tag = f"exact:{stats['exact']} fuzzy:{stats['fuzzy']} ai:{stats['ai']} miss:{stats['miss']}"
                print(f"   [{done:>3}/{total}] Dòng {orig_idx+2}: {found} trang tìm được  ({tag})")
            except Exception as e:
                print(f"   ❌ [{done}/{total}] Lỗi dòng {orig_idx+2}: {e}")

    # ── Cập nhật DataFrame ─────────────────────────────────────────────────────
    for idx, updates in all_updates.items():
        for col, val in updates.items():
            df.at[idx, col] = val

    # ── Lưu file ───────────────────────────────────────────────────────────────
    base, ext = os.path.splitext(args.excel)
    output_path = base + "_with_pages" + ext
    df.to_excel(output_path, index=False)
    apply_beautiful_format(output_path)

    # ── Tổng kết ───────────────────────────────────────────────────────────────
    total_found = total_stats["exact"] + total_stats["fuzzy"] + total_stats["ai"]
    total_all   = total_found + total_stats["miss"]
    pct = (total_found / total_all * 100) if total_all else 0
    print(f"""
╔══════════════════════════════════════════════╗
║  ✅ HOÀN TẤT  
║  📄 Kết quả: {output_path}
║  ─────────────────────────────────────────
║  Tổng trích dẫn xử lý : {total_all:>5}
║  Tìm được trang        : {total_found:>5}  ({pct:.1f}%)
║    - Exact match       : {total_stats['exact']:>5}
║    - Fuzzy match       : {total_stats['fuzzy']:>5}
║    - AI fallback       : {total_stats['ai']:>5}
║  Không tìm được        : {total_stats['miss']:>5}
╚══════════════════════════════════════════════╝
""")


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()  # bắt buộc trên Windows khi dùng ProcessPoolExecutor
    main()