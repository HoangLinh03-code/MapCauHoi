"""
export_md.py - Chuyển đổi dữ liệu Lịch Sử từ Excel sang Markdown
=============================================================================
Tính năng:
  - Sử dụng file JSON Mục lục để map 100% chính xác [SỐ CHƯƠNG] và [SỐ BÀI].
  - TỰ ĐỘNG ĐẶT TÊN FILE THEO FORM: LICHSU_KNTT_[MÃ LỚP]_[CHƯƠNG]_[BÀI]
  - Fallback thông minh: Nếu bài nào không có trong JSON, tự động lùi về dùng Regex.
"""

import os
import sys
import argparse
import re
import pandas as pd
import json

# ══════════════════════════════════════════════════════════════════════════════
# XỬ LÝ CHUỖI & ĐỌC JSON MỤC LỤC
# ══════════════════════════════════════════════════════════════════════════════

def normalize_text(text: str) -> str:
    """Chuẩn hóa chuỗi (chuyển chữ thường, xóa khoảng trắng thừa) để so khớp."""
    if not text or pd.isna(text): return ""
    return re.sub(r'\s+', ' ', str(text).lower().strip())

def build_toc_mapping(json_path: str) -> dict:
    """
    Đọc JSON mục lục và quét để tạo map:
    { "bài 1: liên hợp quốc": {"chuong": "1", "bai": "1"} }
    """
    mapping = {}
    if not json_path or not os.path.exists(json_path):
        return mapping

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Mẹo: Quét toàn bộ chuỗi text trong JSON (bất chấp cấu trúc JSON là gì)
        # Giữ đúng thứ tự từ trên xuống dưới của file mục lục
        strings = re.findall(r'"([^"]*)"', content)

        current_chuong_num = "0"
        for s in strings:
            s = s.strip()
            if not s: continue

            # Phát hiện Chủ đề / Chương
            chuong_match = re.search(r'(?i)^(?:chủ đề|chương|phần)\s*(\d+)', s)
            if chuong_match:
                current_chuong_num = chuong_match.group(1)
                continue

            # Phát hiện Bài
            bai_match = re.search(r'(?i)^bài\s*(\d+)', s)
            if bai_match:
                bai_num = bai_match.group(1)
                norm_key = normalize_text(s)
                mapping[norm_key] = {
                    "chuong": current_chuong_num,
                    "bai": bai_num
                }

        print(f"  🗺️ Đã map thành công {len(mapping)} Bài từ file JSON Mục lục.")
    except Exception as e:
        print(f"  ⚠️ Lỗi khi đọc JSON mục lục: {e}")

    return mapping

# ══════════════════════════════════════════════════════════════════════════════
# REGEX FALLBACK (Dự phòng nếu JSON thiếu)
# ══════════════════════════════════════════════════════════════════════════════

def get_chuong_number(text: str, seq_id: int) -> str:
    if not text or pd.isna(text): return str(seq_id)
    match = re.search(r'(?i)(?:chủ đề|chương|phần)\s*(\d+)', str(text).strip())
    return match.group(1) if match else str(seq_id)

def get_bai_number(text: str) -> str:
    if not text or pd.isna(text): return "0"
    text = str(text).strip()
    match = re.search(r'(?i)bài\s*(\d+)', text)
    if match: return match.group(1)
    
    matches = re.findall(r'\d+', text)
    if matches:
        for m in matches:
            if int(m) < 100: return m  # Tránh năm 1945, 1986...
        return matches[0]
    return "0"

def format_quote_text(text: str) -> str:
    text = str(text).strip()
    if not text: return ""
    return f"> {text}" if (text.startswith('"') and text.endswith('"')) else f'> "{text}"'

def format_source_text(text: str) -> str:
    text = str(text).strip()
    if not text: return ""
    if text.startswith("(") and text.endswith(")"): text = text[1:-1]
    return f"> **({text})**"

# ══════════════════════════════════════════════════════════════════════════════
# LOGIC RENDER MARKDOWN
# ══════════════════════════════════════════════════════════════════════════════

def export_to_markdown(df: pd.DataFrame, output_dir: str, ma_lop: str, toc_mapping: dict):
    df = df.fillna("")
    quote_cols = [c for c in df.columns if str(c).startswith("Nội dung trích dẫn ")]
    
    chu_de_list = []
    for cd in df["Chủ đề / Chương"]:
        cd_str = str(cd).strip()
        if cd_str and cd_str not in chu_de_list:
            chu_de_list.append(cd_str)
    chu_de_map = {name: i + 1 for i, name in enumerate(chu_de_list)}
    
    grouped = df.groupby("Bài", sort=False)
    exported_count = 0
    
    for bai, group in grouped:
        bai_str = str(bai).strip()
        if not bai_str: continue
            
        chu_de_first = str(group.iloc[0].get("Chủ đề / Chương", "")).strip()
        
        # KIỂM TRA MAPPING TỪ JSON TRƯỚC
        norm_bai = normalize_text(bai_str)
        if toc_mapping and norm_bai in toc_mapping:
            chuong_num = toc_mapping[norm_bai]["chuong"]
            bai_num = toc_mapping[norm_bai]["bai"]
        else:
            # Fallback dùng Regex dự phòng
            seq_id = chu_de_map.get(chu_de_first, 0)
            chuong_num = get_chuong_number(chu_de_first, seq_id)
            bai_num = get_bai_number(bai_str)
        
        file_name = f"LICHSU_KNTT_{ma_lop}_{chuong_num}_{bai_num}_material.md"
        file_path = os.path.join(output_dir, file_name)
        
        md_lines = []
        current_chu_de, current_muc, current_tieu_muc = None, None, None
        
        for _, row in group.iterrows():
            chu_de = str(row.get("Chủ đề / Chương", "")).strip()
            muc = str(row.get("Mục", "")).strip()
            tieu_muc = str(row.get("Tiểu mục", "")).strip()
            
            # 1. Chủ đề & Bài
            if chu_de and chu_de != current_chu_de:
                md_lines.append(f"# {chu_de.upper()}\n")
                md_lines.append(f"## {bai_str.upper()}\n")
                current_chu_de = chu_de
            if not current_chu_de:
                md_lines.append(f"## {bai_str.upper()}\n")
                current_chu_de = "Đã in bài"

            # 2. Mục
            if muc and muc != current_muc:
                md_lines.append(f"### {muc}\n")
                current_muc, current_tieu_muc = muc, None
                
            # 3. Tiểu mục
            if tieu_muc and tieu_muc != current_tieu_muc:
                md_lines.append(f"**{tieu_muc}**\n")
                current_tieu_muc = tieu_muc
                
            # 4. Trích dẫn
            has_quotes = False
            for col in quote_cols:
                idx = col.replace("Nội dung trích dẫn ", "")
                nd = str(row.get(col, "")).strip()
                ng = str(row.get(f"Trích nguồn {idx}", "")).strip()
                
                if nd or ng:
                    has_quotes = True
                    if nd: md_lines.append(format_quote_text(nd))
                    if ng: md_lines.append(format_source_text(ng))
                    md_lines.append("")
            if not has_quotes: md_lines.append("")

        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines))
            
        print(f"  ✅ Đã tạo: {file_name}")
        exported_count += 1
        
    return exported_count

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="File Excel/CSV")
    parser.add_argument("--output", default="./export_markdown", help="Thư mục xuất file")
    parser.add_argument("--malop", default="C10", help="Mã lớp (VD: C10, C11, C12)")
    parser.add_argument("--toc", default="", help="Đường dẫn file JSON Mục lục (Tùy chọn)")
    
    args = parser.parse_args()

    print("=" * 60)
    print("  XUẤT MARKDOWN (MAPPING TỪ FILE JSON GỐC)")
    print("=" * 60)

    os.makedirs(args.output, exist_ok=True)
    
    toc_mapping = build_toc_mapping(args.toc) if args.toc else {}

    print(f"\n📖 Đang đọc dữ liệu từ: {args.input}...")
    try:
        if args.input.endswith(".csv"):
            df = pd.read_csv(args.input, dtype=str)
        else:
            try:
                df = pd.read_excel(args.input, sheet_name="DuLieu", dtype=str)
            except ValueError:
                df = pd.read_excel(args.input, dtype=str)
    except Exception as e:
        print(f"❌ Lỗi đọc file: {e}")
        sys.exit(1)

    print(f"\n⚙️  Đang xuất Markdown...")
    count = export_to_markdown(df, args.output, args.malop, toc_mapping)

    print(f"""
╔══════════════════════════════════════════════════════╗
║  ✅ HOÀN TẤT
║  📂 Thư mục lưu      : {args.output}
║  📄 Tổng số file .md : {count}
╚══════════════════════════════════════════════════════╝
""")

if __name__ == "__main__":
    main()