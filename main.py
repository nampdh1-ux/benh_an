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
from typing import Any, Dict, List, Optional

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.templating import Jinja2Templates
from fpdf import FPDF
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel
import google.generativeai as genai

app = FastAPI(title="Bệnh Án Lâm Sàng - FastAPI Backend")
templates = Jinja2Templates(directory="templates")

# ==============================================================================
# 1. CẤU HÌNH BIẾN MÔI TRƯỜNG & AUTHENTICATION
# ==============================================================================
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "123456")
AUTH_SECRET_KEY = os.getenv("AUTH_SECRET_KEY", "clinical_secret_2026")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

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
        msg['Subject'] = f"🔑 Mã xác thực OTP Bệnh án Lâm sàng: {otp_code}"
        body = f"Mã xác thực OTP của bạn là:\n\n👉 {otp_code} 👈"
        msg.attach(MIMEText(body, 'plain'))
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception:
        return False

# Pydantic Schemas
class SendOTPRequest(BaseModel):
    email: str

class LoginRequest(BaseModel):
    email: str
    otp: str
    password: str

# ==============================================================================
# 2. ROUTE GIAO DIỆN CHÍNH (HTML)
# ==============================================================================
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")

# ==============================================================================
# 3. ROUTES AUTHENTICATION
# ==============================================================================
@app.post("/api/auth/send-otp")
async def api_send_otp(req: SendOTPRequest):
    email = req.email.strip().lower()
    if not re.match(r"^[a-zA-Z0-9](\.?[a-zA-Z0-9_-]){5,29}@gmail\.com$", email):
        raise HTTPException(status_code=400, detail="Vui lòng nhập Gmail hợp lệ!")
    otp = str(random.randint(100000, 999999))
    otp_storage[email] = otp
    if send_otp_email(email, otp):
        return {"success": True, "message": f"Đã gửi OTP tới {email}"}
    raise HTTPException(status_code=500, detail="Lỗi kết nối gửi email.")

@app.post("/api/auth/login")
async def api_login(req: LoginRequest):
    email = req.email.strip().lower()
    if email not in otp_storage or otp_storage[email] != req.otp.strip():
        raise HTTPException(status_code=400, detail="Mã OTP không chính xác!")
    if req.password.strip() != APP_PASSWORD:
        raise HTTPException(status_code=400, detail="Mật khẩu nội bộ không đúng!")
    token = generate_auth_token(email)
    otp_storage.pop(email, None)
    return {"success": True, "token": token, "email": email}

# ==============================================================================
# 4. ROUTES AI (CHẨN ĐOÁN, ĐIỀU TRỊ, TIÊN LƯỢNG, PHẢN BIỆN)
# ==============================================================================
def get_benh_su_text(payload: Dict[str, Any]) -> str:
    if payload.get("loai_benh_an") == "Hậu phẫu":
        return f"- Trước mổ: {payload.get('bs_truoc_mo', '')}\n- Trong mổ: {payload.get('bs_trong_mo', '')}\n- Sau mổ: {payload.get('bs_sau_mo', '')}"
    return payload.get("benh_su", "")

@app.post("/api/ai/cdpb")
async def api_ai_cdpb(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model: raise HTTPException(status_code=500, detail="Thiếu API Key")
    
    benh_su_str = get_benh_su_text(payload)
    context = (
        f"Loại: {payload.get('loai_benh_an')}\n"
        f"Tuổi: {payload.get('tuoi')}, Giới: {payload.get('gioi_tinh')}\n"
        f"Lý do: {payload.get('ly_do_vao_vien')}\n"
        f"Bệnh sử: {benh_su_str}\n"
        f"Khám: {payload.get('kham_toan_than')}\n"
        f"Sơ bộ: {payload.get('chan_doan_so_bo')}"
    )
    prompt = f"""
    Bạn là bác sĩ lâm sàng. Dựa vào ca bệnh ({context}), đưa ra:
    1. Chẩn đoán phân biệt
    2. Biện luận chẩn đoán sơ bộ
    Trả về ĐÚNG 2 thẻ: [CHAN_DOAN_PHAN_BIET] và [BIEN_LUAN_SO_BO]
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

@app.post("/api/ai/treatment")
async def api_ai_treatment(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model: raise HTTPException(status_code=500, detail="Thiếu API Key")
    context = f"Loại: {payload.get('loai_benh_an')}\nTuổi: {payload.get('tuoi')}, {payload.get('gioi_tinh')}\nTiền sử: {payload.get('ts_noi_khoa')}\nChẩn đoán: {payload.get('chan_doan_xac_dinh')}"
    prompt = f"Bạn là bác sĩ điều trị. Xây dựng phác đồ cho ca bệnh ({context}). Yêu cầu trả về đúng 3 tag: [MUC_TIEU], [DIEU_TRI_CU_THE], [THEO_DOI]."
    txt = model.generate_content(prompt).text
    mt, ct, td = "", "", ""
    if "[MUC_TIEU]" in txt and "[DIEU_TRI_CU_THE]" in txt and "[THEO_DOI]" in txt:
        p1 = txt.split("[DIEU_TRI_CU_THE]")
        mt = p1[0].replace("[MUC_TIEU]", "").strip()
        p2 = p1[1].split("[THEO_DOI]")
        ct = p2[0].strip()
        td = p2[1].strip()
    return {"dt_muc_tieu": mt, "dt_cu_the": ct, "dt_theo_doi": td}

@app.post("/api/ai/prognosis")
async def api_ai_prognosis(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model: raise HTTPException(status_code=500, detail="Thiếu API Key")
    context = f"Loại: {payload.get('loai_benh_an')}\nChẩn đoán: {payload.get('chan_doan_xac_dinh')}\nĐiều trị: {payload.get('dt_cu_the')}"
    prompt = f"Bạn là bác sĩ lâm sàng. Đưa ra TIÊN LƯỢNG và TƯ VẤN cho ca bệnh ({context}). Trả về 2 tag: [TIEN_LUONG] và [TU_VAN]."
    res_text = model.generate_content(prompt).text
    tl, tv = "", ""
    if "[TIEN_LUONG]" in res_text and "[TU_VAN]" in res_text:
        parts = res_text.split("[TU_VAN]")
        tl = parts[0].replace("[TIEN_LUONG]", "").strip()
        tv = parts[1].strip()
    return {"tien_luong": tl, "tu_van": tv}

@app.post("/api/ai/critique")
async def api_ai_critique(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model: raise HTTPException(status_code=500, detail="Thiếu API Key")
    phong_cach = payload.get("phong_cach", "Học thuật & Hướng dẫn")
    context = f"Bệnh sử: {get_benh_su_text(payload)}\nChẩn đoán SB: {payload.get('chan_doan_so_bo')}\nChẩn đoán XĐ: {payload.get('chan_doan_xac_dinh')}"
    prompt = f"""Bạn là Giảng viên lâm sàng. Nhận xét ca bệnh ({context}) theo phong cách {phong_cach}. 
    Trả về ĐÚNG định dạng JSON: {{"nhan_xet_tong_the": "...", "danh_sach_cau_hoi": [{{"chu_de": "...", "cau_hoi": "...", "goi_y_tra_loi": "..."}}]}}"""
    res_pb = model.generate_content(prompt).text.strip()
    if res_pb.startswith("```json"): res_pb = res_pb[7:]
    elif res_pb.startswith("```"): res_pb = res_pb[3:]
    if res_pb.endswith("```"): res_pb = res_pb[:-3]
    try:
        return json.loads(res_pb.strip())
    except Exception:
        return {"nhan_xet_tong_the": "Lỗi phân tích cú pháp từ AI.", "danh_sach_cau_hoi": []}

@app.post("/api/ocr/batch")
async def api_ocr_batch(files: List[UploadFile] = File(...)):
    model = get_ai_model()
    if not model: raise HTTPException(status_code=500, detail="Chưa cấu hình API Key")
    results = []
    ocr_prompt = "Đọc phiếu xét nghiệm và trả về JSON có 2 key: 'ket_qua' (chỉ số) và 'phien_giai' (biện luận)."
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

# ==============================================================================
# 5. XUẤT FILE PDF
# ==============================================================================
def format_bullet_points(text: str) -> str:
    if not text or not str(text).strip(): return "Chưa ghi nhận thông tin."
    lines = str(text).strip().split("\n")
    formatted_lines = []
    for line in lines:
        cleaned = line.strip()
        if cleaned:
            if not cleaned.startswith("-") and not cleaned.startswith("*"): formatted_lines.append(f"- {cleaned}")
            else: formatted_lines.append(cleaned)
    return "\n".join(formatted_lines)

class BenhAnPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Fallback font an toàn nếu không tìm thấy ttf
        try:
            if os.path.exists("Roboto-Regular.ttf") and os.path.exists("Roboto-Bold.ttf"):
                self.add_font("Roboto", "", "Roboto-Regular.ttf")
                self.add_font("Roboto-Bold", "", "Roboto-Bold.ttf")
                self.font_normal = "Roboto"
                self.font_bold = "Roboto-Bold"
            else:
                self.font_normal = "Helvetica"
                self.font_bold = "Helvetica"
        except:
            self.font_normal = "Helvetica"
            self.font_bold = "Helvetica"

    def header(self):
        if self.page_no() == 1:
            self.set_font(self.font_bold, "", 15)
            self.cell(0, 8, "BENH AN LAM SANG", align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(5)

    def add_section_header(self, title):
        self.set_font(self.font_bold, "", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, title, fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_body_text(self, text):
        self.set_font(self.font_normal, "", 9.5)
        self.multi_cell(0, 5, str(text).strip() if str(text).strip() else "Chua ghi nhan thong tin.")
        self.ln(2)

@app.post("/api/export/pdf")
async def api_export_pdf(data: Dict[str, Any]):
    pdf = BenhAnPDF()
    pdf.add_page()
    
    pdf.add_section_header("I. HANH CHINH")
    hc_text = (
        f"- Ho ten: {str(data.get('ho_ten', '')).upper()} | Tuoi: {data.get('tuoi')} | Gioi: {data.get('gioi_tinh')}\n"
        f"- Khoa phong: {data.get('khoa_phong')}\n"
        f"- Ngay vao vien: {data.get('ngay_vao_vien')}"
    )
    pdf.add_body_text(hc_text)

    pdf.add_section_header("II. LY DO VAO VIEN")
    pdf.add_body_text(data.get('ly_do_vao_vien', ''))

    pdf.add_section_header("III. BENH SU")
    pdf.add_body_text(get_benh_su_text(data))

    pdf.add_section_header("IV. KHAM LAM SANG")
    pdf.add_body_text(format_bullet_points(data.get('kham_toan_than', '')))

    pdf.add_section_header("V. CHAN DOAN SO BO")
    pdf.add_body_text(data.get('chan_doan_so_bo', ''))

    pdf.add_section_header("VI. CHAN DOAN XAC DINH")
    pdf.add_body_text(data.get('chan_doan_xac_dinh', ''))

    pdf.add_section_header("VII. DIEU TRI & TIEN LUONG")
    pdf.add_body_text(format_bullet_points(data.get('dt_cu_the', '')))
    pdf.add_body_text(format_bullet_points(data.get('tien_luong', '')))

    pdf_bytes = bytes(pdf.output())
    return Response(content=pdf_bytes, media_type="application/pdf", headers={"Content-Disposition": "inline; filename=benhan.pdf"})

# ==============================================================================
# 6. XUẤT FILE POWERPOINT
# ==============================================================================
@app.post("/api/export/pptx")
async def api_export_pptx(data: Dict[str, Any]):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # Slide bìa
    slide = prs.slides.add_slide(blank_layout)
    box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
    tf = box.text_frame
    p1 = tf.paragraphs[0]
    p1.text = "BỆNH ÁN LÂM SÀNG"
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(13, 71, 161)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf.add_paragraph()
    p2.text = f"Bệnh nhân: {str(data.get('ho_ten', '')).upper()} | {data.get('tuoi')} tuổi"
    p2.font.size = Pt(20)
    p2.alignment = PP_ALIGN.CENTER

    # Slide Bệnh sử
    slide2 = prs.slides.add_slide(blank_layout)
    box2 = slide2.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(6.0))
    tf2 = box2.text_frame
    tf2.word_wrap = True
    p_title = tf2.paragraphs[0]
    p_title.text = "BỆNH SỬ"
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(13, 71, 161)
    
    p_content = tf2.add_paragraph()
    p_content.text = get_benh_su_text(data)
    p_content.font.size = Pt(16)

    # Slide Chẩn đoán
    slide3 = prs.slides.add_slide(blank_layout)
    box3 = slide3.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(6.0))
    tf3 = box3.text_frame
    tf3.word_wrap = True
    p3_title = tf3.paragraphs[0]
    p3_title.text = "CHẨN ĐOÁN"
    p3_title.font.size = Pt(24)
    p3_title.font.bold = True
    
    p3_c1 = tf3.add_paragraph()
    p3_c1.text = f"Sơ bộ: {data.get('chan_doan_so_bo', '')}"
    p3_c1.font.size = Pt(16)
    
    p3_c2 = tf3.add_paragraph()
    p3_c2.text = f"Xác định: {data.get('chan_doan_xac_dinh', '')}"
    p3_c2.font.size = Pt(16)
    p3_c2.font.color.rgb = RGBColor(180, 0, 0)

    pptx_io = io.BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return Response(content=pptx_io.getvalue(), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers={"Content-Disposition": "attachment; filename=benh_an.pptx"})
