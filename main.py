import hashlib
import io
import json
import os
import random
import re
import smtplib
import unicodedata
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
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel
import google.generativeai as genai

app = FastAPI(title="Bệnh Án Lâm Sàng Win2K")
templates = Jinja2Templates(directory="templates")

# Cấu hình biến môi trường
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
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        return genai.GenerativeModel(model_name)
    except Exception:
        return None

def strip_vietnamese_accents(text: str) -> str:
    """Loại bỏ dấu tiếng Việt dự phòng khi xuất PDF bằng font chuẩn Latinh"""
    if not text:
        return ""
    text = unicodedata.normalize('NFD', str(text))
    text = re.sub(r'[\u0300-\u036f]', '', text)
    text = text.replace('đ', 'd').replace('Đ', 'D')
    return text

def format_bullet_points(text: str) -> str:
    if not text or not str(text).strip():
        return "Chua ghi nhan thong tin."
    lines = str(text).strip().split("\n")
    formatted_lines = []
    for line in lines:
        cleaned = line.strip()
        if cleaned:
            if not cleaned.startswith("-") and not cleaned.startswith("*"):
                formatted_lines.append(f"- {cleaned}")
            else:
                formatted_lines.append(cleaned)
    return "\n".join(formatted_lines)

def get_benh_su_text(payload: Dict[str, Any]) -> str:
    if payload.get("loai_benh_an") == "Hậu phẫu":
        return f"- Truoc mo: {payload.get('bs_truoc_mo', '')}\n- Trong mo: {payload.get('bs_trong_mo', '')}\n- Sau mo: {payload.get('bs_sau_mo', '')}"
    return payload.get("benh_su", "")

# --- GIAO DIỆN CHÍNH ---
@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")

# --- AI ENDPOINTS ---
@app.post("/api/ai/cdpb")
async def api_ai_cdpb(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY trên Render!")
    
    benh_su_str = get_benh_su_text(payload)
    context = (
        f"Loại bệnh án: {payload.get('loai_benh_an')}\n"
        f"Bệnh nhân: {payload.get('tuoi')} tuổi, Giới tính: {payload.get('gioi_tinh')}\n"
        f"Lý do vào viện: {payload.get('ly_do_vao_vien')}\n"
        f"Bệnh sử: {benh_su_str}\n"
        f"Khám toàn thân: {payload.get('kham_toan_than')}\n"
        f"Chẩn đoán sơ bộ: {payload.get('chan_doan_so_bo')}"
    )
    prompt = f"""
    Bạn là bác sĩ lâm sàng thực thụ. Dựa vào ca bệnh dưới đây ({context}):
    1. Đưa ra danh sách Chẩn đoán phân biệt theo thứ tự ưu tiên.
    2. Biện luận chẩn đoán sơ bộ chặt chẽ.
    Trả về ĐÚNG 2 thẻ:
    [CHAN_DOAN_PHAN_BIET]
    ...
    [BIEN_LUAN_SO_BO]
    ...
    """
    try:
        resp = model.generate_content(prompt).text
        cdpb, bl = "", ""
        if "[CHAN_DOAN_PHAN_BIET]" in resp and "[BIEN_LUAN_SO_BO]" in resp:
            parts = resp.split("[BIEN_LUAN_SO_BO]")
            cdpb = parts[0].replace("[CHAN_DOAN_PHAN_BIET]", "").strip()
            bl = parts[1].strip()
        else:
            cdpb = resp.strip()
        return {"chan_doan_phan_biet": cdpb, "bien_luan": bl}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi gọi AI: {str(e)}")

@app.post("/api/ai/treatment")
async def api_ai_treatment(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY!")
    context = f"Loại: {payload.get('loai_benh_an')}\nChẩn đoán: {payload.get('chan_doan_xac_dinh')}\nTiền sử: {payload.get('ts_noi_khoa')}"
    prompt = f"Bạn là bác sĩ điều trị. Xây dựng phác đồ cho ca bệnh ({context}). Trả về ĐÚNG 3 thẻ: [MUC_TIEU], [DIEU_TRI_CU_THE], [THEO_DOI]."
    try:
        txt = model.generate_content(prompt).text
        mt, ct, td = "", "", ""
        if "[MUC_TIEU]" in txt and "[DIEU_TRI_CU_THE]" in txt and "[THEO_DOI]" in txt:
            p1 = txt.split("[DIEU_TRI_CU_THE]")
            mt = p1[0].replace("[MUC_TIEU]", "").strip()
            p2 = p1[1].split("[THEO_DOI]")
            ct = p2[0].strip()
            td = p2[1].strip()
        return {"dt_muc_tieu": mt, "dt_cu_the": ct, "dt_theo_doi": td}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi AI: {str(e)}")

@app.post("/api/ai/prognosis")
async def api_ai_prognosis(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY!")
    context = f"Chẩn đoán: {payload.get('chan_doan_xac_dinh')}\nĐiều trị: {payload.get('dt_cu_the')}"
    prompt = f"Bạn là bác sĩ lâm sàng. Đưa ra TIÊN LƯỢNG và TƯ VẤN cho ca bệnh ({context}). Trả về 2 thẻ: [TIEN_LUONG] và [TU_VAN]."
    try:
        res_text = model.generate_content(prompt).text
        tl, tv = "", ""
        if "[TIEN_LUONG]" in res_text and "[TU_VAN]" in res_text:
            parts = res_text.split("[TU_VAN]")
            tl = parts[0].replace("[TIEN_LUONG]", "").strip()
            tv = parts[1].strip()
        return {"tien_luong": tl, "tu_van": tv}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi AI: {str(e)}")

@app.post("/api/ai/critique")
async def api_ai_critique(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY!")
    phong_cach = payload.get("phong_cach", "Học thuật & Hướng dẫn")
    context = f"Bệnh sử: {get_benh_su_text(payload)}\nChẩn đoán SB: {payload.get('chan_doan_so_bo')}\nChẩn đoán XĐ: {payload.get('chan_doan_xac_dinh')}"
    prompt = f"""Bạn là Giảng viên lâm sàng. Nhận xét ca bệnh ({context}) theo phong cách {phong_cach}.
    Trả về ĐÚNG định dạng JSON thuần không bọc code markdown: {{"nhan_xet_tong_the": "...", "danh_sach_cau_hoi": [{{"chu_de": "...", "cau_hoi": "...", "goi_y_tra_loi": "..."}}]}}"""
    try:
        res_pb = model.generate_content(prompt).text.strip()
        if res_pb.startswith("```json"): res_pb = res_pb[7:]
        if res_pb.startswith("```"): res_pb = res_pb[3:]
        if res_pb.endswith("```"): res_pb = res_pb[:-3]
        return json.loads(res_pb.strip())
    except Exception:
        return {"nhan_xet_tong_the": "Giảng viên đã ghi nhận ca bệnh. Cần chú ý kiểm tra thêm cận lâm sàng và các chỉ số sinh hiệu.", "danh_sach_cau_hoi": []}

@app.post("/api/ocr/batch")
async def api_ocr_batch(files: List[UploadFile] = File(...)):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY!")
    results = []
    ocr_prompt = "Đọc phiếu xét nghiệm và trả về JSON có 2 key: 'ket_qua' (liệt kê chỉ số dạng xuống dòng) và 'phien_giai' (biện luận ngắn gọn)."
    for file in files:
        try:
            contents = await file.read()
            img = Image.open(io.BytesIO(contents))
            resp = model.generate_content([ocr_prompt, img]).text.strip()
            if resp.startswith("```json"): resp = resp[7:]
            if resp.startswith("```"): resp = resp[3:]
            if resp.endswith("```"): resp = resp[:-3]
            results.append(json.loads(resp.strip()))
        except Exception:
            results.append({"ket_qua": "Khong the phan tich anh", "phien_giai": "-"})
    return {"results": results}

# --- XUẤT VĂN BẢN PDF CHỐNG LỖI 500 ---
class RobustPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.use_unicode = False
        if os.path.exists("Roboto-Regular.ttf") and os.path.exists("Roboto-Bold.ttf"):
            try:
                self.add_font("Roboto", "", "Roboto-Regular.ttf")
                self.add_font("Roboto-Bold", "", "Roboto-Bold.ttf")
                self.font_reg = "Roboto"
                self.font_bld = "Roboto-Bold"
                self.use_unicode = True
            except Exception:
                self.font_reg = "Helvetica"
                self.font_bld = "Helvetica"
        else:
            self.font_reg = "Helvetica"
            self.font_bld = "Helvetica"

    def clean(self, text: Any) -> str:
        s = str(text or "")
        return s if self.use_unicode else strip_vietnamese_accents(s)

    def header(self):
        if self.page_no() == 1:
            self.set_font(self.font_bld, "B" if not self.use_unicode else "", 15)
            self.cell(0, 8, self.clean("BENH AN LAM SANG"), align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(4)

    def add_sec(self, title: str):
        self.set_font(self.font_bld, "B" if not self.use_unicode else "", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, self.clean(title), fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_txt(self, text: str):
        self.set_font(self.font_reg, "", 9.5)
        self.multi_cell(0, 5, self.clean(text) if str(text).strip() else "Chua ghi nhan.")
        self.ln(2)

@app.post("/api/export/pdf")
async def api_export_pdf(payload: Dict[str, Any]):
    try:
        pdf = RobustPDF()
        pdf.add_page()
        
        pdf.add_sec("I. PHAN HANH CHINH")
        hc = (
            f"- Ho ten: {str(payload.get('ho_ten', '')).upper()} | Tuoi: {payload.get('tuoi')} | Gioi: {payload.get('gioi_tinh')}\n"
            f"- Khoa phong: {payload.get('khoa_phong')} | Dan toc: {payload.get('dan_tok')}\n"
            f"- Ngay vao vien: {payload.get('ngay_vao_vien')} | Nguoi lam BA: {payload.get('sinh_vien')}"
        )
        pdf.add_txt(hc)

        pdf.add_sec("II. LY DO VAO VIEN")
        pdf.add_txt(payload.get("ly_do_vao_vien", ""))

        pdf.add_sec("III. BENH SU")
        pdf.add_txt(get_benh_su_text(payload))

        pdf.add_sec("IV. TIEN SU")
        ts = f"- Noi khoa: {payload.get('ts_noi_khoa', '')}\n- Ngoai khoa: {payload.get('ts_ngoai_khoa', '')}\n- Ban than: {payload.get('ts_loi_song', '')}\n- Gia dinh: {payload.get('ts_gia_dinh', '')}"
        pdf.add_txt(ts)

        pdf.add_sec("V. THAM KHAM LAM SANG")
        pdf.add_txt(format_bullet_points(payload.get("kham_toan_than", "")))
        if payload.get("loai_benh_an") == "Hậu phẫu":
            pdf.add_txt(f"- Ngay hau phau: {payload.get('ngay_hau_phau', '')}\n- Vet mo: {payload.get('kham_vet_mo', '')}\n- Dan luu: {payload.get('kham_dan_luu', '')}")

        pdf.add_sec("VI. CHAN DOAN SO BO & PHAN BIET")
        pdf.add_txt(f"- Sơ bộ: {payload.get('chan_doan_so_bo', '')}\n- Phân biệt: {payload.get('chan_doan_phan_biet', '')}\n- Biện luận: {payload.get('bien_luan', '')}")

        pdf.add_sec("VII. TOM TAT BENH AN")
        pdf.add_txt(payload.get("tom_tat", ""))

        pdf.add_sec("VIII. CHAN DOAN XAC DINH")
        pdf.add_txt(payload.get("chan_doan_xac_dinh", ""))

        pdf.add_sec("IX. DIEU TRI & TIEN LUONG")
        dt = f"- Muc tieu: {payload.get('dt_muc_tieu', '')}\n- Cu the: {payload.get('dt_cu_the', '')}\n- Theo doi: {payload.get('dt_theo_doi', '')}\n- Tien luong: {payload.get('tien_luong', '')}"
        pdf.add_txt(dt)

        pdf_output = pdf.output()
        pdf_bytes = bytes(pdf_output) if isinstance(pdf_output, (bytes, bytearray)) else pdf_output.encode("latin-1", errors="ignore")
        return Response(content=pdf_bytes, media_type="application/pdf", headers={"Content-Disposition": "inline; filename=benhan.pdf"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo PDF: {str(e)}")

@app.post("/api/export/pptx")
async def api_export_pptx(payload: Dict[str, Any]):
    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]
        
        slide = prs.slides.add_slide(blank)
        box = slide.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(3.5))
        tf = box.text_frame
        p1 = tf.paragraphs[0]
        p1.text = "BỆNH ÁN LÂM SÀNG"
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = RGBColor(13, 71, 161)
        p1.alignment = PP_ALIGN.CENTER
        
        p2 = tf.add_paragraph()
        p2.text = f"Bệnh nhân: {str(payload.get('ho_ten', '')).upper()} | {payload.get('tuoi')} tuổi"
        p2.font.size = Pt(20)
        p2.alignment = PP_ALIGN.CENTER

        pptx_io = io.BytesIO()
        prs.save(pptx_io)
        pptx_io.seek(0)
        return Response(content=pptx_io.getvalue(), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation", headers={"Content-Disposition": "attachment; filename=benhan.pptx"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi tạo PPTX: {str(e)}")
