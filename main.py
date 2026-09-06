import hashlib
import io
import json
import os
import random
import re
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Any, Dict, List

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from fpdf import FPDF
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pydantic import BaseModel
import google.generativeai as genai

app = FastAPI(title="Benh An Lam Sang Win2K")
templates = Jinja2Templates(directory="templates")

# Cấu hình biến môi trường
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "123456")
AUTH_SECRET_KEY = os.getenv("AUTH_SECRET_KEY", "clinical_secret_2026")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Bộ nhớ tạm mã OTP
otp_storage: Dict[str, str] = {}

def get_ai_model(model_name: str = "gemini-2.5-flash"):
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
        msg['Subject'] = f"🔑 Ma xac thuc OTP Benh an Lam sang: {otp_code}"
        body = f"Ma xac thuc OTP cua ban la:\n\n👉 {otp_code} 👈"
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

# --- CÁC PYDANTIC SCHEMA AN TOÀN TUYỆT ĐỐI (TRÁNH LỖI TUPLE) ---
class SendOTPRequest(BaseModel):
    email: str

class LoginRequest(BaseModel):
    email: str
    otp: str
    password: str

# --- GIAO DIỆN CHÍNH ---
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")

# --- AUTHENTICATION ---
@app.post("/api/auth/send-otp")
async def api_send_otp(req: SendOTPRequest):
    email = req.email.strip().lower()
    if not re.match(r"^[a-zA-Z0-9](\.?[a-zA-Z0-9_-]){5,29}@gmail\.com$", email):
        raise HTTPException(status_code=400, detail="Vui long nhap Gmail hop le (@gmail.com)!")
    otp = str(random.randint(100000, 999999))
    otp_storage[email] = otp
    if send_otp_email(email, otp):
        return {"success": True, "message": f"Da gui OTP toi {email}"}
    raise HTTPException(status_code=500, detail="Loi ket noi gui email.")

@app.post("/api/auth/login")
async def api_login(req: LoginRequest):
    email = req.email.strip().lower()
    if email not in otp_storage or otp_storage[email] != req.otp.strip():
        raise HTTPException(status_code=400, detail="Ma OTP khong chinh xac!")
    if req.password.strip() != APP_PASSWORD:
        raise HTTPException(status_code=400, detail="Mat khau khong dung!")
    token = generate_auth_token(email)
    otp_storage.pop(email, None)
    return {"success": True, "token": token, "email": email}

# --- AI CHẨN ĐOÁN ---
@app.post("/api/ai/cdpb")
async def api_ai_cdpb(payload: Dict[str, Any]):
    model = get_ai_model("gemini-2.5-flash")
    if not model:
        raise HTTPException(status_code=500, detail="Chua cau hinh GEMINI_API_KEY")
    prompt = f"""
    Ban la bac si lam sang. Phan tich ca benh sau va dua ra:
    1. Danh sach Chan doan phan biet
    2. Bien luan lam sang
    
    Du lieu:
    {json.dumps(payload, ensure_ascii=False)}

    Xuat dung 2 the:
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

# --- OCR PHIẾU XÉT NGHIỆM ---
@app.post("/api/ocr/batch")
async def api_ocr_batch(files: List[UploadFile] = File(...)):
    model = get_ai_model("gemini-2.5-flash")
    if not model:
        raise HTTPException(status_code=500, detail="Chua cau hinh GEMINI_API_KEY")
    results = []
    ocr_prompt = "Doc phieu xet nghiem va tra ve JSON co 2 key: 'ket_qua' (chi so) va 'phien_giai' (bien luan)."
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
            results.append({"ket_qua": "Khong phan tich duoc", "phien_giai": "-"})
    return {"results": results}

# --- XUẤT PDF ---
class SimpleMedicalPDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "BENH AN LAM SANG", align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(3)

@app.post("/api/export/pdf")
async def api_export_pdf(payload: Dict[str, Any]):
    pdf = SimpleMedicalPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)
    pdf.multi_cell(0, 6, f"Ho ten: {str(payload.get('ho_ten', '')).upper()} - Tuoi: {payload.get('tuoi')} - Gioi tinh: {payload.get('gioi_tinh')}")
    pdf.multi_cell(0, 6, f"Ly do vao vien: {payload.get('ly_do_vao_vien', '')}")
    pdf.multi_cell(0, 6, f"Chan doan so bo: {payload.get('chan_doan_so_bo', '')}")
    pdf.multi_cell(0, 6, f"Chan doan xac dinh: {payload.get('chan_doan_xac_dinh', '')}")
    pdf_bytes = bytes(pdf.output())
    return Response(content=pdf_bytes, media_type="application/pdf", headers={"Content-Disposition": "inline; filename=benh_an.pdf"})

# --- XUẤT POWERPOINT ---
@app.post("/api/export/pptx")
async def api_export_pptx(payload: Dict[str, Any]):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf = box.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "BENH AN LAM SANG"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(10, 36, 106)
    p2 = tf.add_paragraph()
    p2.text = f"Benh nhan: {str(payload.get('ho_ten', '')).upper()} | {payload.get('tuoi')} tuoi"
    p2.font.size = Pt(20)
    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return Response(content=pptx_io.getvalue(), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers={"Content-Disposition": "attachment; filename=benh_an.pptx"})