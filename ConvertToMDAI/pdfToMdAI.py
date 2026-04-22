import json
import os
import sys
import tempfile

# Thêm thư mục cha (SachDienTu) vào sys.path để import API khi chạy trực tiếp
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from PyPDF2 import PdfReader, PdfWriter
from API.callAPIforPDF import VertexClient
from google.oauth2 import service_account
from dotenv import load_dotenv

load_dotenv(os.path.join(os.path.dirname(__file__), 'API\.env'))

PAGES_PER_CHUNK = 20

def clean_md_response(text):
    """Loại bỏ wrapper ```markdown ... ``` mà AI thường trả về."""
    if not text:
        return ""
    text = text.strip()
    # Xóa ```markdown ở đầu và ``` ở cuối
    if text.startswith("```markdown"):
        text = text[len("```markdown"):].strip()
    elif text.startswith("```md"):
        text = text[len("```md"):].strip()
    elif text.startswith("```"):
        text = text[3:].strip()
    if text.endswith("```"):
        text = text[:-3].strip()
    return text

def split_pdf(pdf_path, pages_per_chunk=PAGES_PER_CHUNK):
    """Chia PDF thành các file tạm, mỗi file có pages_per_chunk trang.
    Trả về list các tuple: (chunk_path, start_page, end_page, total_pages)
    """
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    chunks = []
    temp_dir = tempfile.mkdtemp(prefix="pdf_chunks_")

    for start in range(0, total_pages, pages_per_chunk):
        end = min(start + pages_per_chunk, total_pages)
        writer = PdfWriter()
        for i in range(start, end):
            writer.add_page(reader.pages[i])

        chunk_filename = f"chunk_{start+1:04d}_{end:04d}.pdf"
        chunk_path = os.path.join(temp_dir, chunk_filename)
        with open(chunk_path, "wb") as f:
            writer.write(f)

        chunks.append((chunk_path, start + 1, end, total_pages))

    return chunks, temp_dir

def _create_client(model):
    """Tạo VertexClient với credentials từ .env"""
    service_account_data = {
        "type": os.getenv("TYPE"),
        "project_id": os.getenv("PROJECT_ID"),
        "private_key_id": os.getenv("PRIVATE_KEY_ID"),
        "private_key": os.getenv("PRIVATE_KEY").replace('\\n', '\n'),
        "client_email": os.getenv("CLIENT_EMAIL"),
        "client_id": os.getenv("CLIENT_ID", ""),
        "auth_uri": os.getenv("AUTH_URI"),
        "token_uri": os.getenv("TOKEN_URI"),
        "auth_provider_x509_cert_url": os.getenv("AUTH_PROVIDER_X509_CERT_URL"),
        "client_x509_cert_url": os.getenv("CLIENT_X509_CERT_URL"),
        "universe_domain": os.getenv("UNIVERSE_DOMAIN")
    }

    creds = service_account.Credentials.from_service_account_info(
        service_account_data,
        scopes=["https://www.googleapis.com/auth/cloud-platform"]
    )

    return VertexClient(
        project_id=os.getenv('PROJECT_ID'),
        creds=creds,
        model_name=model
    )

SYSTEM_INSTRUCTION = """Bạn là chuyên viên số hóa tài liệu chuyên nghiệp. 
Nhiệm vụ duy nhất của bạn là chuyển đổi file PDF sang định dạng Markdown với độ chính xác tuyệt đối.
Bạn KHÔNG ĐƯỢC tóm tắt, lược bỏ, hay thay đổi bất kỳ nội dung nào. 
Output của bạn là TOÀN BỘ nội dung gốc được format lại dưới dạng Markdown chuẩn."""

PROMPT_TEMPLATE = """Chuyển đổi TOÀN BỘ nội dung của file PDF đính kèm sang định dạng Markdown. 
Đây là phần {chunk_info} của tài liệu (trang {start_page} đến {end_page}, tổng {total_pages} trang).

## QUY TẮC BẮT BUỘC:
1. **TOÀN VẸN NỘI DUNG**: Giữ nguyên 100% nội dung gốc. KHÔNG tóm tắt, KHÔNG lược bỏ, KHÔNG paraphrase. Mỗi câu, mỗi từ trong PDF đều phải xuất hiện trong output.
2. **CẤU TRÚC HEADING**: 
   - Tên sách/tiêu đề chính → `#` (H1) — chỉ dùng 1 lần duy nhất ở phần đầu tiên
   - Chương/Phần lớn → `##` (H2)
   - Mục con trong chương → `###` (H3)
   - Tiểu mục → `####` (H4)
   - Giữ đúng thứ bậc heading, không nhảy cấp
3. **BẢNG BIỂU**: Convert tất cả bảng sang Markdown table syntax (`| col1 | col2 |`). Giữ nguyên mọi dữ liệu trong bảng.
4. **HÌNH ẢNH/BIỂU ĐỒ**: Mô tả bằng `![Mô tả nội dung hình ảnh]()`. Mô tả phải chi tiết, bao gồm nội dung text trong ảnh nếu có.
5. **CHÚ THÍCH/FOOTNOTE**: Giữ nguyên tất cả chú thích cuối trang. Dùng format `[^1]` cho tham chiếu và `[^1]: nội dung` cho phần chú thích.
6. **DANH SÁCH**: 
   - Danh sách có số thứ tự → dùng `1. 2. 3.`
   - Danh sách không thứ tự → dùng `- `
7. **ĐỊNH DẠNG VĂN BẢN**:
   - Chữ in đậm → `**text**`  
   - Chữ in nghiêng → `*text*`
   - Trích dẫn → `> text`
8. **SỐ LIỆU & TÊN RIÊNG**: Giữ chính xác tuyệt đối mọi con số, ngày tháng, tên người, tên địa danh, thuật ngữ chuyên ngành.
9. **PHÂN CÁCH**: Dùng `---` để phân cách các phần lớn (giữa các chương).
10. **KHÔNG WRAP**: Trả về trực tiếp nội dung Markdown, KHÔNG bao bọc trong block code.

## OUTPUT: Chỉ trả về nội dung Markdown, không kèm lời giải thích hay ghi chú nào khác."""


def getBookMenuFromAI(file_name, pdf_path, output_folder, failed_log_path, model="gemini-2.5-pro"):
    """Xử lý PDF → Markdown. Tự động chia nhỏ file lớn thành các chunk."""
    
    # Đếm số trang
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    print(f"📖 File: {file_name} — {total_pages} trang")

    # Tạo client
    client = _create_client(model)

    # Tạo thư mục output
    os.makedirs(output_folder, exist_ok=True)
    md_path = os.path.join(output_folder, f"{file_name}.md")

    if total_pages <= PAGES_PER_CHUNK:
        # File nhỏ — xử lý 1 lần
        print(f"⏳ File nhỏ ({total_pages} trang), xử lý 1 lần...")
        _process_single(client, pdf_path, md_path, 1, total_pages, total_pages, failed_log_path, file_name)
    else:
        # File lớn — chia chunk
        chunks, temp_dir = split_pdf(pdf_path)
        total_chunks = len(chunks)
        print(f"📦 Chia thành {total_chunks} phần (mỗi phần ~{PAGES_PER_CHUNK} trang)")

        all_md_parts = []
        for idx, (chunk_path, start_page, end_page, total) in enumerate(chunks, 1):
            print(f"\n⏳ [{idx}/{total_chunks}] Đang xử lý trang {start_page}-{end_page}...")
            try:
                md_text = _process_chunk(client, chunk_path, start_page, end_page, total, idx, total_chunks)
                if md_text:
                    all_md_parts.append(md_text)
                    print(f"   ✔ Xong phần {idx}/{total_chunks} ({len(md_text)} ký tự)")
                else:
                    print(f"   ⚠️ Phần {idx}/{total_chunks} trả về rỗng!")
                    with open(failed_log_path, "a", encoding="utf-8") as log:
                        log.write(f"{file_name} - Chunk {idx} (trang {start_page}-{end_page}) - Trả về rỗng\n")
            except Exception as e:
                print(f"   ❌ Lỗi phần {idx}/{total_chunks}: {e}")
                with open(failed_log_path, "a", encoding="utf-8") as log:
                    log.write(f"{file_name} - Chunk {idx} (trang {start_page}-{end_page}) - Lỗi: {e}\n")

        # Ghép tất cả các phần
        if all_md_parts:
            merged = "\n\n".join(all_md_parts)
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(merged)
            print(f"\n✅ Hoàn thành! Đã ghép {len(all_md_parts)}/{total_chunks} phần → {md_path}")
        else:
            print(f"\n❌ Không có phần nào thành công!")

        # Dọn file tạm
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        print(f"🗑️ Đã xóa file tạm")


def _process_single(client, pdf_path, md_path, start_page, end_page, total_pages, failed_log_path, file_name):
    """Xử lý file PDF nhỏ (1 lần gọi)."""
    prompt = PROMPT_TEMPLATE.format(
        chunk_info="duy nhất",
        start_page=start_page,
        end_page=end_page,
        total_pages=total_pages
    )
    try:
        response_text = client.send_data_to_AI(
            prompt,
            file_paths=[pdf_path],
            temperature=0.1,
            top_p=0.8,
            system_instruction=SYSTEM_INSTRUCTION
        )
        cleaned_text = clean_md_response(response_text)
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(cleaned_text)
        print(f"✔ Đã lưu: {md_path}")
    except Exception as e:
        print(f"❌ Lỗi: {e}")
        with open(failed_log_path, "a", encoding="utf-8") as log:
            log.write(f"{file_name} - {pdf_path} - Lỗi: {e}\n")


def _process_chunk(client, chunk_path, start_page, end_page, total_pages, chunk_idx, total_chunks):
    """Xử lý 1 chunk PDF, trả về markdown text."""
    if chunk_idx == 1:
        chunk_info = f"đầu tiên (1/{total_chunks})"
    elif chunk_idx == total_chunks:
        chunk_info = f"cuối cùng ({chunk_idx}/{total_chunks})"
    else:
        chunk_info = f"{chunk_idx}/{total_chunks}"

    prompt = PROMPT_TEMPLATE.format(
        chunk_info=chunk_info,
        start_page=start_page,
        end_page=end_page,
        total_pages=total_pages
    )

    response_text = client.send_data_to_AI(
        prompt,
        file_paths=[chunk_path],
        temperature=0.1,
        top_p=0.8,
        system_instruction=SYSTEM_INSTRUCTION
    )
    return clean_md_response(response_text)


def scan_folder(folder):
    failed_log_path = os.path.join(os.path.dirname(folder), "FailedFile.txt")
    # Xóa log cũ nếu có
    if os.path.exists(failed_log_path):
        os.remove(failed_log_path)

    for root, dirs, files in os.walk(folder):
        for f in files:
            if f.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, f)
                file_name = os.path.splitext(os.path.basename(pdf_path))[0]

                # Lấy tên thư mục chứa file pdf
                parent_folder_name = os.path.basename(root)

                # Tạo đường dẫn thư mục đầu ra
                output_folder = os.path.join(
                    os.path.dirname(folder),
                    "SDT_Done",
                    parent_folder_name
                )

                getBookMenuFromAI(file_name, pdf_path, output_folder, failed_log_path)

if __name__ == "__main__":
    pdf_path = r"D:\\Lịch sử văn minh thế giới - Vũ Dương Ninh\\CHƯƠNG I. VĂN MINH BẮC PHI VÀ TÂY Á.pdf"
    file_name = os.path.splitext(os.path.basename(pdf_path))[0]
    output_folder = os.path.join(os.path.dirname(pdf_path), "SDT_Done", "SachDienTu")
    failed_log = os.path.join(os.path.dirname(pdf_path), "FailedFile.txt")
    getBookMenuFromAI(file_name, pdf_path, output_folder, failed_log)
