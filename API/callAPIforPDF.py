import os
import base64
from google import genai
from google.genai import types

class VertexClient:
    def __init__(self, project_id, creds, model_name="gemini-2.5-pro", location="us-central1"):
        # 1. Giữ nguyên cấu hình đã test thành công ở test_connect
        self.model_name = model_name
        self.location = location
        
        print(f"⚡ [VertexClient] Kết nối: {model_name} | Region: {location}")
        
        self.client = genai.Client(
            vertexai=True,
            project=project_id,
            location=location,
            credentials=creds
        )

    def send_data_to_AI(self, prompt, file_paths=None, temperature=0.1, top_p=0.8, system_instruction=None):
        contents = []

        # 2. Xử lý PDF: Cách đóng gói an toàn nhất cho SDK mới
        if file_paths:
            for file_path in file_paths:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as f:
                        file_bytes = f.read()
                    
                    # Dùng types.Part.from_bytes là chuẩn nhất
                    # Nhưng để chắc chắn, ta kiểm tra xem file có rỗng không
                    if len(file_bytes) == 0:
                        print(f"⚠️ File rỗng: {file_path}")
                        continue
                        
                    pdf_part = types.Part.from_bytes(
                        data=file_bytes,
                        mime_type="application/pdf"
                    )
                    contents.append(pdf_part)

        # 3. Thêm Prompt Text
        contents.append(types.Part.from_text(text=prompt))

        # 4. Config (có hỗ trợ system_instruction)
        config_kwargs = {
            "temperature": temperature,
            "top_p": top_p
        }
        if system_instruction:
            config_kwargs["system_instruction"] = system_instruction
        
        config = types.GenerateContentConfig(**config_kwargs)

        try:
            print(f"⏳ Đang gửi request tới {self.model_name}...")
            
            # Gửi request
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=[types.Content(role="user", parts=contents)],
                config=config
            )
            
            return response.text

        except Exception as e:
            err_msg = str(e)
            print(f"❌ Lỗi API: {err_msg}")
            
            # Phân tích lỗi cụ thể giúp bạn
            if "400" in err_msg:
                if "loading the file" in err_msg or "mime" in err_msg:
                    print("👉 Gợi ý: Model này có thể đang kén file PDF. Thử convert sang ảnh hoặc text.")
                elif "not supported" in err_msg:
                    print("👉 Gợi ý: Model 'Preview' này có thể chưa hỗ trợ Multimodal (chỉ nhận Text).")
            return None