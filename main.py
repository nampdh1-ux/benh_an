import os
import io
import json
import random
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional, List, Dict, Any

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import google.generativeai as genai
from fpdf import FPDF
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# ==============================================================================
# KHỞI TẠO ỨNG DỤNG FASTAPI
# ==============================================================================
app = FastAPI(title="Bệnh Án Lâm Sàng API")

# Cấu hình CORS để cho phép Frontend gọi API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Trong thực tế nên điền domain frontend của bạn
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==============================================================================
# HÀM CẤU HÌNH CÁC BIẾN MÔI TRƯỜNG & AI MODEL
# ==============================================================================
def get_feature_model(model_name: str = "gemini-1.5-flash"):
    """Lấy Gemini Model từ biến môi trường của Render."""
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("Chưa cấu hình GEMINI_API_KEY")
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(model_name)

# ==============================================================================
# BẢO MẬT & XÁC THỰC DANH TÍNH (OTP + GMAIL)
# ==============================================================================
class OTPRequest(BaseModel):
    email: str

def send_otp_email(target_email: str, otp_code: str) -> bool:
    sender_mail = os.getenv("SENDER_EMAIL")
    sender_pass = os.getenv("SENDER_APP_PASSWORD")
    if not (sender_mail and sender_pass):
        return False
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_mail
        msg['To'] = target_email
        msg['Subject'] = f"🔑 Mã xác thực truy cập Bệnh án Lâm sàng: {otp_code}"
        body = f"Xin chào,\n\nMã xác thực (OTP) dùng để đăng nhập vào Ứng dụng Bệnh án Lâm sàng của bạn là:\n\n👉  {otp_code}  👈\n\nMã có hiệu lực cho phiên hiện tại."
        msg.attach(MIMEText(body, 'plain'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(sender_mail, sender_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Lỗi gửi email: {e}")
        return False

# ==============================================================================
# CÁC HÀM XỬ LÝ NÉN ẢNH VÀ ĐỊNH DẠNG TEXT
# ==============================================================================
def optimize_lab_image(photo_bytes: bytes, max_dimension=1600, quality=85) -> Image.Image:
    try:
        img = Image.open(io.BytesIO(photo_bytes))
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        width, height = img.size
        if max(width, height) > max_dimension:
            if width > height:
                new_width = max_dimension
                new_height = int(height * (max_dimension / width))
            else:
                new_height = max_dimension
                new_width = int(width * (max_dimension / height))
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
        
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=quality, optimize=True)
        buffer.seek(0)
        return Image.open(buffer)
    except Exception:
        return Image.open(io.BytesIO(photo_bytes))

# ==============================================================================
# LỚP TẠO FILE PDF (Dùng font chuẩn để tránh lỗi trên Render)
# ==============================================================================
class BenhAnPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.loai_ba = "Nội khoa / Tiền phẫu"

    def header(self):
        if self.page_no() == 1:
            self.set_font("Helvetica", "B", 15)
            title = "BENH AN HAU PHAU" if self.loai_ba == "Hau phau" else "BENH AN LAM SANG"
            self.cell(0, 8, title, align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font("Helvetica", "", 8)
            self.cell(0, 4, f"Thoi gian tao: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(3)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "", 8)
        self.cell(0, 10, f"Trang {self.page_no()}/{{nb}}", align="C")

    def add_section_header(self, title):
        self.set_font("Helvetica", "B", 11)
        self.set_fill_color(225, 235, 245)
        # Loại bỏ dấu tiếng việt hoặc dùng encode an toàn do FPDF mặc định không hỗ trợ UTF-8 tốt nếu ko có font
        safe_title = title.encode("latin-1", "replace").decode("latin-1")
        self.cell(0, 7, safe_title, fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_body_text(self, text):
        self.set_font("Helvetica", "", 9.5)
        safe_text = str(text).encode("latin-1", "replace").decode("latin-1") if str(text).strip() else "Chua ghi nhan thong tin."
        self.multi_cell(0, 5, safe_text)
        self.ln(2)

# ==============================================================================
# ROUTE API ENDPOINTS
# ==============================================================================

@app.get("/")
async def health_check():
    return {"status": "ok", "message": "API Bệnh Án Lâm Sàng đang hoạt động"}

@app.post("/api/send-otp")
async def handle_send_otp(req: OTPRequest):
    GMAIL_REGEX = r"^[a-zA-Z0-9](\.?[a-zA-Z0-9_-]){5,29}@gmail\.com$"
    if not re.match(GMAIL_REGEX, req.email):
        return JSONResponse({"status": "error", "message": "Địa chỉ Gmail không hợp lệ!"}, status_code=400)
    
    otp = str(random.randint(100000, 999999))
    success = send_otp_email(req.email, otp)
    if success:
        return {"status": "success", "message": f"Đã gửi OTP tới {req.email}", "otp_hash": otp} # Thực tế nên hash OTP
    return JSONResponse({"status": "error", "message": "Không thể gửi email OTP, vui lòng kiểm tra cấu hình SMTP."}, status_code=500)


@app.post("/api/ai-suggest")
async def ai_suggest(prompt_type: str = Form(...), context: str = Form(...)):
    try:
        model = get_feature_model()
    except ValueError as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    prompts = {
        "summary": f"Dựa vào thông tin lâm sàng sau, hãy tóm tắt bệnh án ngắn gọn, đúng chuẩn y khoa:\n{context}",
        "reasoning": f"Dựa vào tóm tắt và chẩn đoán sơ bộ sau, hãy viết luận đoán lâm sàng chi tiết:\n{context}",
        "treatment": f"Đề xuất hướng điều trị chi tiết và kế hoạch theo dõi cho bệnh án sau:\n{context}"
    }
    
    selected_prompt = prompts.get(prompt_type, f"Phân tích bệnh án y khoa:\n{context}")
    
    try:
        response = model.generate_content(selected_prompt)
        return {"status": "success", "result": response.text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi AI Gemini: {str(e)}")


@app.post("/api/ocr-lab")
async def ocr_lab(file: UploadFile = File(...)):
    try:
        model = get_feature_model()
    except ValueError as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    content = await file.read()
    image = optimize_lab_image(content)
    
    prompt = "Hãy trích xuất toàn bộ kết quả xét nghiệm/cận lâm sàng từ ảnh này dưới dạng văn bản cấu trúc danh sách ngắn gọn."
    try:
        response = model.generate_content([prompt, image])
        return {"status": "success", "extracted_text": response.text}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi OCR: {str(e)}")


@app.post("/api/export-pdf")
async def export_pdf(payload: Dict[str, Any]):
    pdf = BenhAnPDF()
    pdf.loai_ba = payload.get("loai_benh_an", "Noi khoa / Tien phau")
    pdf.add_page()
    
    pdf.add_section_header("I. HANH CHINH")
    pdf.add_body_text(f"Ho va ten: {payload.get('ho_ten', '')} | Tuoi: {payload.get('tuoi', '')} | Gioi tinh: {payload.get('gioi_tinh', '')}")
    pdf.add_body_text(f"Khoa phong: {payload.get('khoa_phong', '')} | Dia chi: {payload.get('dia_chi', '')}")
    
    pdf.add_section_header("II. LY DO VAO VIEN")
    pdf.add_body_text(payload.get("ly_do_vao_vien", ""))
    
    pdf.add_section_header("III. BENH SU")
    pdf.add_body_text(payload.get("benh_su", ""))
    
    pdf.add_section_header("IV. TOM TAT BENH AN")
    pdf.add_body_text(payload.get("tom_tat", ""))
    
    pdf.add_section_header("V. CHAN DOAN SO BO")
    pdf.add_body_text(payload.get("chan_doan_so_bo", ""))

    buffer = io.BytesIO(pdf.output())
    buffer.seek(0)
    
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=benh_an_lam_sang.pdf"}
    )