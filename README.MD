<!-- @format -->

# Hệ Thống Chuyển Đổi & Map Trích Dẫn Sách Giáo Khoa

Pipeline tự động chuyển đổi sách tham khảo (PDF) sang Markdown, map đoạn trích với mục lục SGK, bổ sung số trang, gộp Excel và xuất file Markdown cuối cùng.

---

## Mục Lục

- [Yêu Cầu Hệ Thống](#yêu-cầu-hệ-thống)
- [Cài Đặt](#cài-đặt)
- [Cấu Trúc Thư Mục](#cấu-trúc-thư-mục)
- [Quy Trình Chạy Từng Bước](#quy-trình-chạy-từng-bước)
  - [Bước 1: Chuyển PDF sang Markdown (AI)](#bước-1-chuyển-pdf-sang-markdown-ai)
  - [Bước 2: Map đoạn trích với mục lục SGK](#bước-2-map-đoạn-trích-với-mục-lục-sgk)
  - [Bước 3: Xuất Excel để kiểm tra](#bước-3-xuất-excel-để-kiểm-tra)
  - [Bước 4: Bổ sung số trang vào trích nguồn](#bước-4-bổ-sung-số-trang-vào-trích-nguồn)
  - [Bước 5: Gộp các file Excel](#bước-5-gộp-các-file-excel)
  - [Bước 6: Xuất Markdown cuối cùng](#bước-6-xuất-markdown-cuối-cùng)
- [Tổng Quan Pipeline](#tổng-quan-pipeline)
- [Xử Lý Lỗi Thường Gặp](#xử-lý-lỗi-thường-gặp)
- [Ghi Chú](#ghi-chú)

---

## Yêu Cầu Hệ Thống

| Thành phần        | Yêu cầu                                       |
| :---------------- | :-------------------------------------------- |
| **Python**        | >= 3.10                                       |
| **OS**            | Windows 10/11                                 |
| **Google Cloud**  | Service Account có quyền Vertex AI            |
| **Poppler**       | `poppler-24.08.0` (cho PDF scan ở Bước 4b)    |
| **Tesseract OCR** | Tesseract + gói ngôn ngữ `vie` (cho PDF scan) |

---

## Cài Đặt

### 1. Tạo môi trường ảo & cài thư viện

```powershell
cd d:\CheckTool\ChuyenDoi
python -m venv venv
.\venv\Scripts\Activate.ps1

pip install -r requirements.txt
```

### 2. Cài Tesseract OCR (cho PDF scan)

Chỉ cần cài nếu sử dụng Bước 4b (PDF scan/ảnh).

- Tải tại: https://github.com/UB-Mannheim/tesseract/wiki
- Khi cài: tick chọn **"Additional language data" -> Vietnamese**
- Hoặc tải thêm `vie.traineddata` từ https://github.com/tesseract-ocr/tessdata và đặt vào thư mục `tessdata/` của Tesseract

---

## Cấu Trúc Thư Mục

```
ChuyenDoi/
├── API/
│   ├── .env                          # Credentials Google Cloud
│   ├── callAPI.py                    # VertexClient chính (text generation)
│   └── callAPIforPDF.py              # VertexClient cho PDF -> MD
├── ConvertToMDAI/
│   └── pdfToMdAI.py                  # Bước 1: PDF -> Markdown (AI)
├── MapMD/
│   ├── fast_map.py                   # Bước 2: Map đoạn trích <-> mục lục SGK
│   ├── prompt_section_map.txt        # Prompt template cho AI mapping
│   └── export_ex.py                  # Bước 3: Xuất kết quả map -> Excel
├── MergeExcel/
│   ├── add_page.py                   # Bước 4a: Thêm trang (PDF kỹ thuật số)
│   ├── add_page_v2.py                # Bước 4b: Thêm trang (PDF scan/ảnh)
│   └── merge.py                      # Bước 5: Gộp nhiều file Excel
├── ConvertToMD/
│   └── export_md.py                  # Bước 6: Excel -> Markdown cuối cùng
├── requirements.txt                  # File chứa danh sách thư viện cần cài
└── Mục lục SGK                       # Thư mục chứa file JSON mục lục SGK
```

---

## Quy Trình Chạy Từng Bước

**Lưu ý chung:** Tất cả lệnh chạy từ thư mục gốc `d:\CheckTool\ChuyenDoi`. Kích hoạt venv trước khi chạy.

```powershell
cd d:\CheckTool\ChuyenDoi
.\venv\Scripts\Activate.ps1
```

---

### Bước 1: Chuyển PDF sang Markdown (AI)

**Mục đích:** Chuyển đổi file PDF sách tham khảo thành file Markdown, giữ nguyên 100% nội dung gốc.

**Script:** `ConvertToMDAI/pdfToMdAI.py`

**Cơ chế hoạt động:**

- PDF nhỏ (dưới hoặc bằng 20 trang): xử lý 1 lần duy nhất
- PDF lớn (trên 20 trang): tự động chia thành các chunk 20 trang, gọi AI từng chunk rồi ghép lại thành file MD hoàn chỉnh
- Sử dụng model `gemini-2.5-pro` qua Vertex AI

**Cách chạy:**

Sửa phần `if __name__` trong file `pdfToMdAI.py`, chỉ định đường dẫn PDF đầu vào và thư mục đầu ra.

**Kết quả:** Các file `.md` được lưu trong thư mục `SDT_Done/` cùng cấp với thư mục PDF đầu vào.

---

### Bước 2: Map đoạn trích với mục lục SGK

**Mục đích:** Đọc file MD (đã tạo ở Bước 1), chia thành các đoạn nhỏ theo heading, sau đó dùng AI map từng đoạn với mục nhỏ nhất trong mục lục SGK (file JSON). Kết quả là file JSON chứa danh sách mapping giữa đoạn trích và mục lục.

**Script:** `MapMD/fast_map.py`
**Prompt:** `MapMD/prompt_section_map.txt`

**Chạy với 1 file MD:**

```powershell
python MapMD/fast_map.py ^
  --toc "Lich su 11-xinhe-10.3.25.json" ^
  --ref-files "D:\DuongDan\SDT_Done\TenSach.md" ^
  --prompt MapMD/prompt_section_map.txt
```

**Chạy với nhiều file MD cùng lúc:**

```powershell
python MapMD/fast_map.py --toc "Lich su 11-xinhe-10.3.25.json" --ref-files "D:\DuongDan\SDT_Done\*.md" --prompt MapMD/prompt_section_map.txt
```

**Tuỳ chọn ghi đè metadata sách (nếu AI parse sai):**

Trong một số trường hợp AI trích xuất sai thông tin sách (tác giả, tên, NXB...), có thể ghi đè bằng tham số dòng lệnh:

```powershell
python MapMD/fast_map.py ^
  --toc "Lich su 11-xinhe-10.3.25.json" ^
  --ref-files "D:\DuongDan\SDT_Done\TenSach.md" ^
  --prompt MapMD/prompt_section_map.txt ^
  --author "Nguyễn Văn A" ^
  --book-title "Tên Sách" ^
  --publisher "NXB Giáo Dục" ^
  --pub-place "Hà Nội" ^
  --pub-year "2023"
```

**Bảng tham số:**

| Tham số       | Mô tả                                                |
| ------------- | ---------------------------------------------------- |
| `--toc`       | File JSON mục lục SGK                                |
| `--ref-files` | Danh sách file MD tham khảo (hỗ trợ wildcard `*.md`) |
| `--prompt`    | File prompt template cho AI                          |
| `--no-resume` | Chạy lại từ đầu, bỏ qua cache đã lưu trước đó        |

**Kết quả:** File `map_<toc>_<ref>.json` được lưu cùng thư mục với file MD đầu vào. File này chứa danh sách các mapping đã tìm được.

---

### Bước 3: Xuất Excel để kiểm tra

**Mục đích:** Chuyển kết quả mapping (file JSON từ Bước 2) thành file Excel có format và màu sắc, dễ kiểm tra bằng tay.

**Script:** `MapMD/export_ex.py`

```powershell
python MapMD/export_ex.py ^
  --toc "Lich su 11-xinhe-10.3.25.json" ^
  --map "D:\DuongDan\SDT_Done\map_Lich su 11_TenSach.json" ^
  --output "D:\DuongDan\KetQua_TenSach.xlsx"
```

**Kết quả:** File Excel gồm 2 sheet:

- **Trích dẫn**: Bảng đầy đủ với các cột STT, Chương, Bài, Mục, Tiểu mục, Số lượng đoạn trích, Nội dung trích dẫn, Trích nguồn
- **Thống kê**: Tổng quan số liệu mapping (bao nhiêu mục đã map, bao nhiêu còn thiếu)

**Sau bước này:** Mở file Excel, kiểm tra nội dung trích dẫn. Nếu trích nguồn thiếu số trang thì chuyển sang Bước 4.

---

### Bước 4: Bổ sung số trang vào trích nguồn

**Mục đích:** Tự động tìm và bổ sung `tr. X` vào cột trích nguồn trong file Excel, dựa trên nội dung đoạn trích và file PDF gốc.

**Chiến lược tìm trang (3 tầng):**

1. **Exact match** -- tìm chuỗi chính xác (sau khi chuẩn hoá unicode) trong text PDF. Ưu tiên khớp đoạn dài (300 ký tự), nếu không tìm được thì giảm xuống 80 ký tự.
2. **Fuzzy match** -- trích 20 từ dài nhất (>= 4 ký tự) từ đoạn đầu excerpt, tìm trang có tỉ lệ từ khớp >= 55%.
3. **AI fallback** -- gọi Gemini qua Vertex AI, gửi kèm file PDF, yêu cầu trả về số trang in chứa đoạn trích. Chỉ dùng khi 2 cách trên thất bại.

**Tham số `--output-dir`:** Cả 2 script đều hỗ trợ tham số `--output-dir` để chỉ định thư mục lưu file kết quả. Thư mục sẽ được tạo tự động nếu chưa tồn tại. Nếu không truyền tham số này, file kết quả sẽ được lưu cùng thư mục với file Excel đầu vào.

#### 4a. Với PDF kỹ thuật số (file có text layer)

**Script:** `MergeExcel/add_page.py`

Script này đọc text trực tiếp từ PDF bằng thư viện `pypdf`, không cần OCR. Phù hợp với các file PDF được tạo từ phần mềm soạn thảo (Word, InDesign...).

```powershell
python MergeExcel/add_page.py ^
  --excel "D:\DuongDan\KetQua_TenSach.xlsx" ^
  --pdf "D:\DuongDan\TenSach.pdf" ^
  --output-dir "D:\DuongDan\output_with_page" ^
  --workers 4
```

Nếu chỉ muốn dùng text search mà không gọi AI:

```powershell
python MergeExcel/add_page.py ^
  --excel "D:\DuongDan\KetQua_TenSach.xlsx" ^
  --pdf "D:\DuongDan\TenSach.pdf" ^
  --output-dir "D:\DuongDan\output_with_page" ^
  --no-ai
```

**Bảng tham số `add_page.py`:**

| Tham số        | Mô tả                                                                |
| -------------- | -------------------------------------------------------------------- |
| `--excel`      | (bắt buộc) Đường dẫn file Excel đầu vào (.xlsx)                      |
| `--pdf`        | (bắt buộc) Đường dẫn file PDF sách gốc                               |
| `--output-dir` | Thư mục lưu file kết quả (tạo tự động). Mặc định: cùng thư mục Excel |
| `--workers`    | Số luồng song song (mặc định: 4)                                     |
| `--no-ai`      | Bỏ qua AI fallback, chỉ dùng exact + fuzzy match                     |

#### 4b. Với PDF scan (file ảnh gộp thành PDF)

**Script:** `MergeExcel/add_page_v2.py`

Script này tương tự `add_page.py` nhưng bổ sung khả năng xử lý PDF scan. Khi phát hiện PDF có ít hơn 20% trang chứa text layer, script tự động chuyển sang OCR bằng Tesseract.

**Yêu cầu bổ sung:**

- **Poppler** (`poppler-24.08.0`): Script tự động tìm thư mục `poppler-24*` ở thư mục gốc project (`d:\CheckTool\ChuyenDoi\`). Nếu chưa có, script sẽ tự tải về. Chỉ cần tải 1 lần duy nhất.
- **Tesseract OCR**: Đã cài với gói ngôn ngữ Vietnamese (xem phần Cài Đặt).

```powershell
python MergeExcel/add_page_v2.py ^
  --excel "D:\DuongDan\KetQua_TenSach.xlsx" ^
  --pdf "D:\DuongDan\TenSach_Scan.pdf" ^
  --output-dir "D:\DuongDan\output_with_page" ^
  --workers 4 ^
  --ocr-cache
```

**Bảng tham số `add_page_v2.py`:**

| Tham số        | Mô tả                                                                |
| -------------- | -------------------------------------------------------------------- |
| `--excel`      | (bắt buộc) Đường dẫn file Excel đầu vào (.xlsx)                      |
| `--pdf`        | (bắt buộc) Đường dẫn file PDF sách gốc                               |
| `--output-dir` | Thư mục lưu file kết quả (tạo tự động). Mặc định: cùng thư mục Excel |
| `--workers`    | Số luồng song song (mặc định: 4)                                     |
| `--no-ai`      | Bỏ qua AI fallback, chỉ dùng text search + OCR                       |
| `--ocr-cache`  | Lưu kết quả OCR vào file JSON để tái dùng lần sau                    |
| `--dpi`        | DPI khi render PDF scan (mặc định: 200, tăng lên 300 nếu chữ nhỏ)    |

**Kết quả:** File `*_with_pages.xlsx` được lưu vào thư mục `--output-dir` (hoặc cùng thư mục Excel nếu không truyền). File này chứa trích nguồn đã bổ sung số trang dạng `tr. X`.

---

### Bước 5: Gộp các file Excel thành 1 file tổng hợp

**Mục đích:** Gộp tất cả file Excel `_with_pages.xlsx` (từ Bước 4) thành 1 file Excel duy nhất. Các dòng trích dẫn từ nhiều sách tham khảo khác nhau được gộp vào cùng 1 bảng.

**Script:** `MergeExcel/merge.py`

**Gộp từ folder:**

```powershell
python MergeExcel/merge.py ^
  --folder "D:\DuongDan\output_with_page" ^
  --output "LichSu11.xlsx"
```

**Gộp từ danh sách file chỉ định:**

```powershell
python MergeExcel/merge.py ^
  --files "file1_with_pages.xlsx" "file2_with_pages.xlsx" "file3_with_pages.xlsx" ^
  --output "LichSu11.xlsx"
```

**Thêm file mới vào file tổng hợp đã có (không ghi đè dữ liệu cũ):**

```powershell
python MergeExcel/merge.py ^
  --base "LichSu11.xlsx" ^
  --files "file_moi_with_pages.xlsx" ^
  --output "LichSu11_v2.xlsx"
```

**Kết quả:** File Excel tổng hợp chứa đầy đủ trích dẫn từ tất cả sách tham khảo.

---

### Bước 6: Xuất Markdown cuối cùng

**Mục đích:** Chuyển file Excel tổng hợp (từ Bước 5) thành các file Markdown riêng theo từng Bài. Mỗi file chứa đoạn trích đã map với mục nhỏ trong SGK, kèm trích nguồn đầy đủ.

**Script:** `ConvertToMD/export_md.py`

```powershell
python ConvertToMD/export_md.py ^
  --input "LichSu11.xlsx" ^
  --output "./export_11" ^
  --malop "C11" ^
  --toc "Lich su 11-xinhe-10.3.25.json"
```

**Bảng tham số:**

| Tham số    | Mô tả                                               |
| ---------- | --------------------------------------------------- |
| `--input`  | File Excel tổng hợp (từ Bước 5)                     |
| `--output` | Thư mục xuất file MD                                |
| `--malop`  | Mã lớp: `C10`, `C11`, `C12`                         |
| `--toc`    | File JSON mục lục SGK (để map chính xác Chương/Bài) |

**Kết quả:** Các file MD được đặt tên theo format:

```
LICHSU_KNTT_C11_1_1_material.md    # Chương 1, Bài 1
LICHSU_KNTT_C11_1_2_material.md    # Chương 1, Bài 2
LICHSU_KNTT_C11_2_3_material.md    # Chương 2, Bài 3
...
```

Mỗi file có cấu trúc:

```markdown
# TÊN CHƯƠNG

## TÊN BÀI

### Tên Mục

**Tên Tiểu mục**

> "Đoạn trích nguyên văn từ sách tham khảo..."
> **(Tác giả, Tên sách, NXB, Nơi XB, Năm, tr. X)**
```

---

## Tổng Quan Pipeline

```
+---------------------------------------------------------+
|  BUOC 1: PDF -> Markdown                                |
|  pdfToMdAI.py                                           |
|  Input:  Sách tham khảo (*.pdf)                         |
|  Output: File *.md (toàn bộ nội dung sách)              |
+--------------+------------------------------------------+
               |
               v
+---------------------------------------------------------+
|  BUOC 2: Map đoạn trích <-> Mục lục SGK                |
|  fast_map.py + prompt_section_map.txt                   |
|  Input:  *.md + TOC JSON                                |
|  Output: map_*.json (danh sách mapping)                 |
+--------------+------------------------------------------+
               |
               v
+---------------------------------------------------------+
|  BUOC 3: JSON -> Excel (kiểm tra)                      |
|  export_ex.py                                           |
|  Input:  map_*.json + TOC JSON                          |
|  Output: *.xlsx (có màu sắc, thống kê)                 |
+--------------+------------------------------------------+
               |   Kiểm tra thủ công
               v
+---------------------------------------------------------+
|  BUOC 4: Bổ sung số trang                               |
|  add_page.py (PDF kỹ thuật số)                          |
|  add_page_v2.py (PDF scan -- cần poppler-24 + Tesseract)|
|  Input:  *.xlsx + *.pdf                                 |
|  Output: *_with_pages.xlsx -> --output-dir              |
+--------------+------------------------------------------+
               |
               v
+---------------------------------------------------------+
|  BUOC 5: Gộp Excel                                      |
|  merge.py                                               |
|  Input:  Folder chứa *_with_pages.xlsx                  |
|  Output: LichSu11.xlsx (file tổng hợp)                 |
+--------------+------------------------------------------+
               |
               v
+---------------------------------------------------------+
|  BUOC 6: Excel -> Markdown cuối cùng                    |
|  export_md.py                                           |
|  Input:  LichSu11.xlsx + TOC JSON                       |
|  Output: LICHSU_KNTT_C11_*_material.md                  |
+---------------------------------------------------------+
```

---

## Xử Lý Lỗi Thường Gặp

| Lỗi                          | Nguyên nhân                   | Cách xử lý                                              |
| ---------------------------- | ----------------------------- | ------------------------------------------------------- |
| `Import lỗi: callAPI`        | Chưa cài thư viện             | `pip install google-genai google-auth python-dotenv`    |
| `Không tìm thấy PRIVATE_KEY` | File `.env` chưa cấu hình     | Kiểm tra file `API/.env`                                |
| `Không tìm thấy Poppler`     | Chưa tải hoặc sai tên thư mục | Script tự tải, hoặc đặt `poppler-24.08.0` ở thư mục gốc |
| `Tesseract thiếu vie`        | Chưa cài gói ngôn ngữ Việt    | Tải `vie.traineddata` vào thư mục `tessdata/`           |
| `PDF scan -> trả về rỗng`    | Chữ nhỏ, DPI thấp             | Tăng `--dpi 300`                                        |
| `AI trả về rỗng`             | Safety filter chặn            | Kiểm tra nội dung, thử lại                              |

---

## Ghi Chú

1. **Cache mapping**: `fast_map.py` tự lưu cache tiến độ. Nếu bị gián đoạn, chạy lại sẽ tiếp tục từ chỗ dừng. Dùng `--no-resume` để chạy lại từ đầu.
2. **Output tự động**: `add_page.py` và `add_page_v2.py` hỗ trợ `--output-dir` để lưu file kết quả vào thư mục riêng, không cần di chuyển file thủ công.
3. **Model AI**: Hiện tại sử dụng `gemini-3.1-pro-preview` (mapping, add_page) và `gemini-2.5-pro` (PDF->MD).
4. **Đa luồng**: `fast_map.py` mặc định 7 luồng, `add_page.py` mặc định 4 luồng -- có thể điều chỉnh qua `--workers`.
5. **Poppler**: Script `add_page_v2.py` tự tìm `poppler-24*` ở thư mục gốc project. Nếu chưa có, sẽ tự tải về lần đầu tiên chạy. Các lần sau tái sử dụng.
