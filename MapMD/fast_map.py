"""
fast_map.py — Bản Tối Thượng (Hỗ trợ chạy nhiều file)
(Đa luồng + Mục nhỏ nhất + CHIA NHỎ ĐOẠN DÀI + Cache riêng từng file + Metadata)
"""

import json
import os
import re
import time
import sys
import hashlib
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import List, Dict, Optional
import glob

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(BASE_DIR)

from dotenv import load_dotenv
load_dotenv(os.path.join(BASE_DIR, 'API', '.env'))

try:
    from API.callAPI import VertexClient, get_vertex_ai_credentials
except ImportError as e:
    print("❌ Import lỗi:", e)
    sys.exit(1)

MAX_WORKERS = 7
MAX_RETRIES = 3
API_DELAY   = 1

MAX_CHUNK_WORDS = 800 

CACHE_LOCK = threading.Lock()

# ══════════════════════════════════════════════════════════════
# 1. HELPERS LẤY DỮ LIỆU TỪ TOC VÀ MD (Giữ nguyên)
# ══════════════════════════════════════════════════════════════
def extract_page_number_from_body(body: str) -> Optional[int]:
    page_candidates = [
        int(m.group(1))
        for m in re.finditer(r'(?:^|\n)\s*(\d{1,4})\s*(?:\n|$)', body)
        if int(m.group(1)) >= 3
    ]
    if not page_candidates:
        return None
    return page_candidates[-1] 

def _extract_book_info_from_content(text: str) -> Dict:
    info = {"ten_tac_gia": "Không rõ", "ten_sach": "Không rõ", "nxb": "Không rõ", "noi_xb": "Hà Nội", "nam_xb": "Không rõ"}
    lines = text.replace('\r\n', '\n').split('\n')[:100] 

    for line in lines:
        stripped = line.strip()
        if stripped.startswith('# ') and not stripped.startswith('## '):
            if not any(kw in stripped.upper() for kw in ['MỤC LỤC', 'LỜI NÓI ĐẦU']):
                info['ten_sach'] = stripped.lstrip('# ').strip()
                break

    for line in lines:
        stripped = line.strip().strip('*')
        if re.search(r'chủ\s+biên', stripped, re.IGNORECASE):
            author = re.sub(r'\s*\(?\s*chủ\s+biên\s*\)?\s*', '', stripped, flags=re.IGNORECASE)
            author = author.strip('*').strip()
            if author: 
                info['ten_tac_gia'] = author
            break

    for line in lines:
        stripped = line.strip().strip('*')
        m = re.search(r'(NHÀ XUẤT BẢN\s+.+)', stripped, re.IGNORECASE)
        if m:
            info['nxb'] = m.group(1).strip().strip('*').strip()
            break

    for line in lines:
        stripped = line.strip()
        m = re.search(r'(\d{4})\s*[/|-]\s*CXB', stripped, re.IGNORECASE)
        if m:
            info['nam_xb'] = m.group(1)
            break
        m_fallback = re.search(r'(?:năm|Năm)\s+(\d{4})', stripped)
        if m_fallback:
            info['nam_xb'] = m_fallback.group(1)

    return info

def _collect_leaves(toc_data: List, result: List):
    if not isinstance(toc_data, list): return
    for chapter in toc_data:
        if not isinstance(chapter, dict): continue
        ch_num = chapter.get("chapter_number", "")
        ch_title = chapter.get("chapter_title", "")
        for lesson in chapter.get("lessons", []):
            if not isinstance(lesson, dict): continue
            ls_num = lesson.get("lesson_number", "")
            ls_title = lesson.get("lesson_title", "")
            for section in lesson.get("sections", []):
                if not isinstance(section, dict): continue
                sc_num = section.get("section_number", "")
                sc_title = section.get("section_title", "")
                subsecs = section.get("subsections", [])
                if subsecs:
                    for sub in subsecs:
                        if not isinstance(sub, dict): continue
                        sub_title = sub.get("subsection_title", "")
                        cache_key = f"{ch_num}|{ls_num}|{sc_num}|{sub_title}"
                        result.append({"cache_key": cache_key, "label": f"{ch_title} > {ls_title} > {sc_title} > {sub_title}"})
                else:
                    cache_key = f"{ch_num}|{ls_num}|{sc_num}|"
                    result.append({"cache_key": cache_key, "label": f"{ch_title} > {ls_title} > {sc_title}"})

def _build_toc_list_for_prompt(leaves: List) -> str:
    lines = []
    for i, leaf in enumerate(leaves):
        lines.append(f"{i+1}. [{leaf['cache_key']}] {leaf['label']}")
    return "\n".join(lines)

def load_prompt_file(prompt_path: str) -> str:
    if not os.path.exists(prompt_path):
        print(f"⚠️ Lỗi: Không tìm thấy file prompt tại {prompt_path}")
        sys.exit(1)
    with open(prompt_path, "r", encoding="utf-8") as f:
        return f.read()

# ══════════════════════════════════════════════════════════════
# 2. CACHE & CHIA ĐOẠN ĐẾN TẬN RĂNG (Giữ nguyên)
# ══════════════════════════════════════════════════════════════
def load_cache(path: str) -> Dict:
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_cache(cache_data: Dict, path: str):
    with CACHE_LOCK:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

def _split_long_text_into_chunks(text: str, max_words=MAX_CHUNK_WORDS) -> List[str]:
    words = text.split()
    if len(words) <= max_words:
        return [text]
    
    chunks = []
    current_chunk = []
    current_count = 0
    
    paragraphs = text.split('\n')
    for p in paragraphs:
        p_words = len(p.split())
        if p_words == 0: continue
        
        if current_count + p_words > max_words and current_chunk:
            chunks.append('\n'.join(current_chunk))
            current_chunk = [p]
            current_count = p_words
        else:
            current_chunk.append(p)
            current_count += p_words
            
    if current_chunk:
        chunks.append('\n'.join(current_chunk))
        
    return chunks

def segment_granular_md(text: str) -> List[Dict]:
    lines = text.split('\n')
    sections = []
    heading_stack = []
    current_text = []

    _re_md_heading = re.compile(r'^(#{1,6})\s+(.+)$')
    _re_plain_heading = re.compile(r'^((?:[IVXLCDM]{1,5}|\d{1,2}|[a-z])[\.\/\)]\s+.+)$', re.IGNORECASE)

    def add_section(h_stack, c_text):
        body = "\n".join(c_text).strip()
        if body and len(body.split()) >= 15:
            full_heading_path = " > ".join([h[1] for h in h_stack]) if h_stack else "Phần mở đầu"
            
            chunks = _split_long_text_into_chunks(body)
            for idx, chunk in enumerate(chunks):
                page_num = extract_page_number_from_body(chunk)
                suffix = f" (phần {idx+1})" if len(chunks) > 1 else ""
                
                sections.append({
                    "heading": f"{full_heading_path}{suffix}", 
                    "text": chunk, 
                    "page": page_num,
                    "original_heading": full_heading_path
                })

    for line in lines:
        stripped = line.strip()
        if not stripped: continue

        m_md = _re_md_heading.match(stripped)
        m_plain = _re_plain_heading.match(stripped)
        
        is_heading, level, h_text = False, 99, ""
        if m_md and len(stripped) <= 150:
            is_heading, level, h_text = True, len(m_md.group(1)), m_md.group(2).strip()
        elif m_plain and len(stripped) <= 150:
            is_heading, level, h_text = True, 7, m_plain.group(1).strip()

        if is_heading:
            add_section(heading_stack, current_text)
            
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()
            heading_stack.append((level, h_text))
            current_text = []
        else:
            current_text.append(line)

    add_section(heading_stack, current_text)
    return sections

# ══════════════════════════════════════════════════════════════
# 3. GỌI API ĐA LUỒNG (Giữ nguyên)
# ══════════════════════════════════════════════════════════════
def process_single_chunk(client, heading: str, chunk_text: str, page: Optional[int], toc_list_str: str, prompt_template: str, book_info: Dict, cache_key: str, cache_dict: Dict, cache_path: str, is_resume: bool) -> List[Dict]:
    if is_resume and cache_key in cache_dict:
        return cache_dict[cache_key], True 
        
    format_kwargs = {
        "ten_tac_gia": book_info.get("ten_tac_gia", "Không rõ"),
        "ten_sach": book_info.get("ten_sach", "Không rõ"),
        "nxb": book_info.get("nxb", "Không rõ"),
        "noi_xb": book_info.get("noi_xb", "Hà Nội"),
        "nam_xb": book_info.get("nam_xb", "Không rõ"),
        "heading": heading,
        "excerpt": chunk_text,
        "toc_list": toc_list_str
    }
    
    page_str = f", tr. {page}" if page else ""
    trich_nguon_chuan = f"({format_kwargs['ten_tac_gia']}, {format_kwargs['ten_sach']}, {format_kwargs['nxb']}, {format_kwargs['noi_xb']}, {format_kwargs['nam_xb']}{page_str})"

    try:
        prompt = prompt_template.format(**format_kwargs)
    except Exception:
        prompt = prompt_template
        for key, value in format_kwargs.items():
            prompt = prompt.replace(f"{{{key}}}", str(value))
    
    mapped_results = []
    
    for attempt in range(MAX_RETRIES):
        try:
            raw = client.send_data_to_AI(
                prompt=prompt,
                file_paths=None,
                temperature=0.0,
                top_p=0.8,
                max_output_tokens=63520,
                use_file_api=False,
            )
            text_resp = str(raw).strip()
            
            if "```json" in text_resp:
                text_resp = re.search(r"```json\s*([\s\S]*?)\s*```", text_resp).group(1)
            elif "```" in text_resp:
                text_resp = re.search(r"```\s*([\s\S]*?)\s*```", text_resp).group(1)
            
            s = text_resp.find('{')
            e = text_resp.rfind('}') + 1
            if s != -1 and e > s:
                text_resp = text_resp[s:e]
                
            data = json.loads(text_resp)
            if "mappings" in data:
                for m in data["mappings"]:
                    if str(m.get("confidence")).lower() in ["high", "medium"]:
                        m["trich_nguon"] = trich_nguon_chuan
                        mapped_results.append(m)
                break 
        except Exception as e:
            time.sleep(API_DELAY * 2)

    cache_dict[cache_key] = mapped_results
    save_cache(cache_dict, cache_path)
    
    return mapped_results, False

# ══════════════════════════════════════════════════════════════
# 4. HÀM XỬ LÝ CHO TỪNG FILE (Mới thêm)
# ══════════════════════════════════════════════════════════════
def process_single_file(ref_file: str, toc_name: str, toc_list_str: str, prompt_template: str, client: VertexClient, args):
    print(f"\n[{'='*50}]")
    print(f"📄 ĐANG XỬ LÝ FILE: {ref_file}")
    
    if not os.path.exists(ref_file):
        print(f"❌ Không tìm thấy file: {ref_file}. Bỏ qua!")
        return

    with open(ref_file, "r", encoding="utf-8") as f:
        md_text = f.read()
    
    book_info = _extract_book_info_from_content(md_text)
    if args.author:     book_info["ten_tac_gia"] = args.author
    if args.book_title: book_info["ten_sach"]    = args.book_title
    if args.publisher:  book_info["nxb"]         = args.publisher
    if args.pub_place:  book_info["noi_xb"]      = args.pub_place
    if args.pub_year:   book_info["nam_xb"]      = args.pub_year
    
    print(f"📖 Sách: {book_info['ten_sach']} | Tác giả: {book_info['ten_tac_gia']} ({book_info['nam_xb']})")

    file_name = os.path.basename(ref_file)
    cache_sig = hashlib.sha1(f"{toc_name}|{file_name}".encode()).hexdigest()[:10]
    work_dir = os.path.dirname(os.path.abspath(ref_file))
    cache_path = os.path.join(work_dir, f"fast_map_cache_{cache_sig}.json") 
    
    is_resume = not args.no_resume
    cache_dict = load_cache(cache_path) if is_resume else {}
    if cache_dict:
        print(f"🔄 Tìm thấy Cache tiến độ ({os.path.basename(cache_path)}) -> Khôi phục {len(cache_dict)} mục.")

    sections = segment_granular_md(md_text)
    print(f"-> Đã băm thành {len(sections)} chunk. Đang xử lý đa luồng ({MAX_WORKERS} luồng)...")

    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {}
        for idx, sec in enumerate(sections):
            cache_key = f"{file_name}#{idx}"
            future = executor.submit(
                process_single_chunk, client, sec["heading"], sec["text"], sec.get("page"),
                toc_list_str, prompt_template, book_info, cache_key, cache_dict, cache_path, is_resume
            )
            futures[future] = (idx, sec)
            
        completed = 0
        for future in as_completed(futures):
            completed += 1
            idx, sec = futures[future]
            try:
                mapped_data, from_cache = future.result()
                
                cache_tag = "[CACHE]" if from_cache else "[API]"
                if mapped_data:
                    status = f"✅ Map thành công ({len(mapped_data)} map)"
                    for md in mapped_data:
                        md["source_heading"] = sec["original_heading"] 
                    results.extend(mapped_data)
                else:
                    status = "⚠️ Bỏ qua"
                    
                page_info = f" | Trang {sec['page']}" if sec.get('page') else ""
                print(f"  {cache_tag} [{completed}/{len(sections)}] Mục: {sec['heading'][:30]}...{page_info} -> {status}")
                
            except Exception as exc:
                print(f"  ❌ [{completed}/{len(sections)}] Lỗi mục {sec['heading'][:20]}: {exc}")

    toc_base = os.path.splitext(toc_name)[0]
    file_base = os.path.splitext(file_name)[0]
    
    # Lưu file output cùng thư mục với file tham chiếu (ref_file)
    output_file = os.path.join(work_dir, f"map_{toc_base}_{file_base}.json")

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"🎉 Hoàn tất file {file_name}! Kết quả lưu tại: {output_file}")


# ══════════════════════════════════════════════════════════════
# 5. CHẠY CHÍNH
# ══════════════════════════════════════════════════════════════
def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--toc", required=True, help="Đường dẫn đến file TOC JSON")
    
    # Sửa: Thay --ref-file thành --ref-files, dùng nargs='+' để nhận nhiều file
    parser.add_argument("--ref-files", nargs='+', required=True, help="Danh sách các file MD (Ví dụ: file1.md file2.md hoặc data/*.md)")
    
    parser.add_argument("--prompt", default="prompt_section_map.txt")
    parser.add_argument("--no-resume", action="store_true")
    
    parser.add_argument("--author", default=None)
    parser.add_argument("--book-title", default=None)
    parser.add_argument("--publisher", default=None)
    parser.add_argument("--pub-place", default=None)
    parser.add_argument("--pub-year", default=None)
    args = parser.parse_args()

    print("\n🚀 BẮT ĐẦU CHẠY MARKDOWN MAPPER (HỖ TRỢ CHẠY NHIỀU FILE) 🚀")
    prompt_template = load_prompt_file(args.prompt)
    
    client = VertexClient(
        project_id=os.getenv("PROJECT_ID", "test-project"),
        creds=get_vertex_ai_credentials(),
        model_name="gemini-3.1-pro-preview",
        region="global"
    )

    # Đọc TOC (Chỉ đọc 1 lần để dùng chung cho mọi file)
    with open(args.toc, 'r', encoding='utf-8') as f:
        toc_data = json.load(f)
    toc_chapters = toc_data.get("chapters", toc_data) if isinstance(toc_data, dict) else toc_data
    leaves = []
    _collect_leaves(toc_chapters, leaves)
    toc_list_str = _build_toc_list_for_prompt(leaves)
    toc_name = os.path.basename(args.toc)

    # Phân tích danh sách file đầu vào (Hỗ trợ wildcard trực tiếp nếu shell không tự xử lý)
    files_to_process = []
    for file_pattern in args.ref_files:
        matched_files = glob.glob(file_pattern)
        if matched_files:
            files_to_process.extend(matched_files)
        else:
            files_to_process.append(file_pattern) # Thêm vào để sau đó check path exists và báo lỗi
    
    files_to_process = list(set(files_to_process)) # Loại bỏ file trùng lặp
    print(f"📌 Tổng số file cần xử lý: {len(files_to_process)}")

    # Duyệt qua từng file và thực thi
    for ref_file in files_to_process:
        process_single_file(ref_file, toc_name, toc_list_str, prompt_template, client, args)
        
    print(f"\n✅ ĐÃ XỬ LÝ XONG TOÀN BỘ {len(files_to_process)} FILE!")

if __name__ == "__main__":
    main()