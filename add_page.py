"""
add_page.py - Bổ sung số trang vào trích nguồn trong file Excel
=================================================================
Chiến lược tìm trang (3 tầng):
  1. Exact match  - tìm chuỗi chính xác trong text đã extract từ PDF
  2. Fuzzy match  - tìm theo keyword dài nhất trích từ đoạn văn
  3. AI fallback  - gọi Vertex AI (Gemini) khi 2 cách trên thất bại

Cách dùng:
  python add_page.py --excel data.xlsx --pdf book.pdf
  python add_page.py --excel data.xlsx --pdf book.pdf --workers 8
  python add_page.py --excel data.xlsx --pdf book.pdf --no-ai   # chỉ dùng text search
"""

import os
import sys
import re
import time
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

# ── callAPI (Vertex AI) ───────────────────────────────────────────────────────
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
try:
    from callAPI import VertexClient, get_vertex_ai_credentials
    HAS_VERTEX = True
except ImportError:
    print("⚠️  Không tìm thấy callAPI.py — sẽ chỉ dùng text search (--no-ai).")
    HAS_VERTEX = False


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


def build_pdf_index(pdf_path: str) -> list[dict]:
    """
    Trả về list các dict:
      { "pdf_idx": int, "printed_page": str|None, "text": str, "text_norm": str }
    """
    print(f"📖 Đang đọc và lập index PDF: {pdf_path}")
    reader = PdfReader(pdf_path)
    index = []
    for i, page in enumerate(reader.pages):
        raw = page.extract_text() or ""
        if not raw.strip():
            continue
        printed = extract_printed_page_number(raw)
        index.append({
            "pdf_idx":     i,
            "printed_page": printed,
            "text":        raw,
            "text_norm":   normalize(raw),
        })
    print(f"   → Đã index {len(index)} trang có nội dung ({len(reader.pages)} trang tổng).")
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
    parser.add_argument("--excel",   required=True,  help="Đường dẫn file Excel (.xlsx)")
    parser.add_argument("--pdf",     required=True,  help="Đường dẫn file PDF sách")
    parser.add_argument("--workers", type=int, default=4, help="Số luồng song song (default: 4)")
    parser.add_argument("--no-ai",   action="store_true",  help="Bỏ qua AI, chỉ dùng text search")
    args = parser.parse_args()

    # ── Kiểm tra file ──────────────────────────────────────────────────────────
    for path in (args.excel, args.pdf):
        if not os.path.exists(path):
            print(f"❌ Không tìm thấy file: {path}")
            sys.exit(1)

    use_ai = not args.no_ai and HAS_VERTEX

    # ── Build PDF index ────────────────────────────────────────────────────────
    pdf_index = build_pdf_index(args.pdf)
    if not pdf_index:
        print("❌ Không đọc được nội dung từ PDF (có thể là ảnh scan). Dừng.")
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
    main()