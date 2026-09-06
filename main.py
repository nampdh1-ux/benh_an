import hashlib
import io
import json
import os
import random
import re
import smtplib
import tempfile
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from fpdf import FPDF
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import google.generativeai as genai

app = FastAPI(title="Bệnh Án Lâm Sàng Win2K")
templates = Jinja2Templates(directory="templates")

# Cấu hình biến môi trường hoặc file cục bộ
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "123456")
AUTH_SECRET_KEY = os.getenv("AUTH_SECRET_KEY", "clinical_secret_2026")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Bộ nhớ tạm cho mã OTP phiên đăng nhập
otp_storage = {}

def get_ai_model(model_name="gemini-3.1-flash-lite"):
    if not GEMINI_API_KEY:
        return None
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel(model_name)

def generate_auth_token(email: str) -> str:
    raw_str = f"{email}_{AUTH_SECRET_KEY}"
    return hashlib.sha256(raw_str.encode("utf-8")).hexdigest()

def send_otp_email(target_email: str, otp_code: str) -> bool:
    if not (SENDER_EMAIL and SENDER_APP_PASSWORD):
        return False
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = target_email
        msg['Subject'] = f"🔑 Mã xác thực truy cập Bệnh án Lâm sàng: {otp_code}"
        body = f"Mã xác thực OTP đăng nhập của bạn là:\n\n👉 {otp_code} 👈"
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

# --- GIAO DIỆN CHÍNH ---
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

# --- AUTHENTICATION ENDPOINTS ---
@app.post("/api/auth/send-otp")
async def api_send_otp(email: str = Form(...)):
    email = email.strip().lower()
    if not re.match(r"^[a-zA-Z0-9](\.?[a-zA-Z0-9_-]){5,29}@gmail\.com$", email):
        raise HTTPException(status_code=400, detail="Vui lòng nhập Gmail hợp lệ!")
    otp = str(random.randint(100000, 999999))
    otp_storage[email] = otp
    if send_otp_email(email, otp):
        return {"success": True, "message": f"Đã gửi OTP đến {email}"}
    raise HTTPException(status_code=500, detail="Lỗi kết nối máy chủ gửi email.")

@app.post("/api/auth/login")
async def api_login(email: str = Form(...), otp: str = Form(...), password: str = Form(...)):
    email = email.strip().lower()
    if email not in otp_storage or otp_storage[email] != otp.strip():
        raise HTTPException(status_code=400, detail="Mã OTP không chính xác!")
    if password.strip() != APP_PASSWORD:
        raise HTTPException(status_code=400, detail="Mật khẩu nội bộ không chính xác!")
    token = generate_auth_token(email)
    otp_storage.pop(email, None)
    return {"success": True, "token": token, "email": email}

# --- AI DIAGNOSIS & LOGIC ---
@app.post("/api/ai/cdpb")
async def api_ai_cdpb(request: Request):
    data = await request.json()
    model = get_ai_model("gemini-3.1-flash-lite")
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY")
    prompt = f"""
    Bạn là bác sĩ lâm sàng giàu kinh nghiệm. Dựa vào ca bệnh dưới đây, hãy đưa ra danh sách Chẩn đoán phân biệt và Biện luận lâm sàng:
    {json.dumps(data, ensure_ascii=False)}
    Xuất ra đúng 2 khối:
    [CHAN_DOAN_PHAN_BIET]
    ...
    [BIEN_LUAN_SO_BO]
    ...
    """
    resp = model.generate_content(prompt).text
    cdpb, bl = "", ""
    if "[CHAN_DOAN_PHAN_BIET]" in resp and "[BIEN_LUAN_SO_BO]" in resp:
        parts = resp.split("[BIEN_LUAN_SO_BO]")
        cdpb = parts[0].replace("[CHAN_DOAN_PHAN_BIET]", "").strip()
        bl = parts[1].strip()
    else:
        cdpb = resp.strip()
    return {"chan_doan_phan_biet": cdpb, "bien_luan": bl}

# --- OCR PHIẾU CẬN LÂM SÀNG ---
@app.post("/api/ocr/batch")
async def api_ocr_batch(files: List[UploadFile] = File(...)):
    model = get_ai_model("gemini-3.1-flash-lite")
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY")
    results = []
    ocr_prompt = "Bạn là bác sĩ xét nghiệm. Đọc phiếu này và trả về JSON chuẩn có 2 khóa: 'ket_qua' (chuỗi chỉ số xuống dòng) và 'phien_giai' (biện luận chỉ số bất thường)."
    for file in files:
        contents = await file.read()
        img = Image.open(io.BytesIO(contents))
        resp = model.generate_content([ocr_prompt, img]).text.strip()
        if resp.startswith("```json"): resp = resp[7:]
        if resp.startswith("```"): resp = resp[3:]
        if resp.endswith("```"): resp = resp[:-3]
        try:
            results.append(json.loads(resp.strip()))
        except Exception:
            results.append({"ket_qua": "Không phân tích được", "phien_giai": "-"})
    return {"results": results}

# --- XUẤT FILE PDF ---
class SimpleMedicalPDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "BENH AN LAM SANG", align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(3)

@app.post("/api/export/pdf")
async def api_export_pdf(request: Request):
    data = await request.json()
    pdf = SimpleMedicalPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)
    pdf.multi_cell(0, 6, f"Ho ten: {data.get('ho_ten', '').upper()} - Tuoi: {data.get('tuoi')} - Gioi tinh: {data.get('gioi_tinh')}")
    pdf.multi_cell(0, 6, f"Ly do vao vien: {data.get('ly_do_vao_vien', '')}")
    pdf.multi_cell(0, 6, f"Chan doan so bo: {data.get('chan_doan_so_bo', '')}")
    pdf.multi_cell(0, 6, f"Chan doan xac dinh: {data.get('chan_doan_xac_dinh', '')}")
    pdf_bytes = bytes(pdf.output())
    return Response(content=pdf_bytes, media_type="application/pdf", headers={"Content-Disposition": "inline; filename=benh_an.pdf"})

# --- XUẤT FILE POWERPOINT ---
@app.post("/api/export/pptx")
async def api_export_pptx(request: Request):
    data = await request.json()
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf = box.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "BỆNH ÁN LÂM SÀNG"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(10, 36, 106)
    p2 = tf.add_paragraph()
    p2.text = f"Bệnh nhân: {data.get('ho_ten', '').upper()} | {data.get('tuoi')} tuổi"
    p2.font.size = Pt(20)
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return Response(content=pptx_io.getvalue(), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers={"Content-Disposition": "attachment; filename=benh_an.pptx"})
