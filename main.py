import hashlib
import io
import json
import os
import random
import re
import smtplib
import urllib.request
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

app = FastAPI(title="Benh An Lam Sang Win2K")
templates = Jinja2Templates(directory="templates")

# Cấu hình biến môi trường
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "123456")
AUTH_SECRET_KEY = os.getenv("AUTH_SECRET_KEY", "clinical_secret_2026")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Tự động tải font Unicode hỗ trợ tiếng Việt nếu chưa có trên server Render
FONT_REGULAR_PATH = "DejaVuSans.ttf"
FONT_BOLD_PATH = "DejaVuSans-Bold.ttf"

def ensure_unicode_fonts():
    base_url = "https://raw.githubusercontent.com/dejavu-fonts/dejavu-fonts/master/ttf/"
    if not os.path.exists(FONT_REGULAR_PATH):
        try:
            urllib.request.urlretrieve(base_url + "DejaVuSans.ttf", FONT_REGULAR_PATH)
        except Exception as e:
            print(f"Không thể tải font Regular: {e}")
    if not os.path.exists(FONT_BOLD_PATH):
        try:
            urllib.request.urlretrieve(base_url + "DejaVuSans-Bold.ttf", FONT_BOLD_PATH)
        except Exception as e:
            print(f"Không thể tải font Bold: {e}")

ensure_unicode_fonts()

def get_ai_model():
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None
    try:
        genai.configure(api_key=api_key)
        # Ưu tiên các dòng flash hỗ trợ tốt và ổn định
        return genai.GenerativeModel("gemini-2.5-flash")
    except Exception:
        return None

def format_bullet_points(text: str) -> str:
    if not text or not str(text).strip():
        return "Chưa ghi nhận thông tin."
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
        return f"- Trước mổ: {payload.get('bs_truoc_mo', '')}\n- Trong mổ: {payload.get('bs_trong_mo', '')}\n- Sau mổ: {payload.get('bs_sau_mo', '')}"
    return payload.get("benh_su", "")

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")

# --- AI ENDPOINTS ---
@app.post("/api/ai/cdpb")
async def api_ai_cdpb(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY trong Environment Variables của Render!")
    
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
    Bạn là bác sĩ lâm sàng thực thụ. Dựa vào ca bệnh:
    {context}
    
    Hãy đưa ra:
    1. Danh sách Chẩn đoán phân biệt
    2. Biện luận chẩn đoán sơ bộ
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
        raise HTTPException(status_code=500, detail=f"Lỗi AI: {str(e)}")

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
    Trả về ĐÚNG định dạng JSON thuần: {{"nhan_xet_tong_the": "...", "danh_sach_cau_hoi": [{{"chu_de": "...", "cau_hoi": "...", "goi_y_tra_loi": "..."}}]}}"""
    try:
        res_pb = model.generate_content(prompt).text.strip()
        if res_pb.startswith("```json"): res_pb = res_pb[7:]
        if res_pb.startswith("```"): res_pb = res_pb[3:]
        if res_pb.endswith("```"): res_pb = res_pb[:-3]
        return json.loads(res_pb.strip())
    except Exception as e:
        return {"nhan_xet_tong_the": f"Lỗi phản biện: {str(e)}", "danh_sach_cau_hoi": []}

@app.post("/api/ocr/batch")
async def api_ocr_batch(files: List[UploadFile] = File(...)):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cấu hình GEMINI_API_KEY!")
    results = []
    ocr_prompt = "Đọc phiếu xét nghiệm và trả về JSON có 2 key: 'ket_qua' (chỉ số xét nghiệm) và 'phien_giai' (biện luận chỉ số bất thường)."
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
            results.append({"ket_qua": "Không thể phân tích ảnh", "phien_giai": "-"})
    return {"results": results}

# --- XUẤT VĂN BẢN PDF UNICODE TIẾNG VIỆT ĐẦY ĐỦ DẤU ---
class UnicodePDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if os.path.exists(FONT_REGULAR_PATH) and os.path.exists(FONT_BOLD_PATH):
            self.add_font("DejaVu", "", FONT_REGULAR_PATH)
            self.add_font("DejaVu", "B", FONT_BOLD_PATH)
            self.font_family_name = "DejaVu"
        else:
            self.font_family_name = "Helvetica"

    def header(self):
        if self.page_no() == 1:
            self.set_font(self.font_family_name, "B", 15)
            self.cell(0, 8, "BỆNH ÁN LÂM SÀNG", align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font(self.font_family_name, "", 9)
            self.cell(0, 4, f"Thời gian lập: {datetime.now().strftime('%d/%m/%Y %H:%M')}", align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(4)

    def add_sec(self, title: str):
        self.set_font(self.font_family_name, "B", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, str(title), fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def add_txt(self, text: str):
        self.set_font(self.font_family_name, "", 9.5)
        self.multi_cell(0, 5, str(text).strip() if str(text).strip() else "Chưa ghi nhận thông tin.")
        self.ln(2)

@app.post("/api/export/pdf")
async def api_export_pdf(payload: Dict[str, Any]):
    try:
        pdf = UnicodePDF()
        pdf.add_page()
        
        pdf.add_sec("I. PHẦN HÀNH CHÍNH")
        hc = (
            f"- Họ và tên: {str(payload.get('ho_ten', '')).upper()} | Tuổi: {payload.get('tuoi')} | Giới: {payload.get('gioi_tinh')}\n"
            f"- Dân tộc: {payload.get('dan_tok')} | Nghề nghiệp: {payload.get('nghe_nghiep')}\n"
            f"- Khoa phòng: {payload.get('khoa_phong')}\n"
            f"- Ngày vào viện: {payload.get('ngay_vao_vien')} | Người làm BA: {payload.get('sinh_vien')}"
        )
        pdf.add_txt(hc)

        pdf.add_sec("II. LÝ DO VÀO VIỆN")
        pdf.add_txt(payload.get("ly_do_vao_vien", ""))

        pdf.add_sec("III. BỆNH SỬ")
        pdf.add_txt(get_benh_su_text(payload))

        pdf.add_sec("IV. TIỀN SỬ")
        ts = f"- Nội khoa: {payload.get('ts_noi_khoa', '')}\n- Ngoại khoa: {payload.get('ts_ngoai_khoa', '')}\n- Lối sống/Thói quen: {payload.get('ts_loi_song', '')}\n- Gia đình: {payload.get('ts_gia_dinh', '')}"
        pdf.add_txt(ts)

        pdf.add_sec("V. THĂM KHÁM LÂM SÀNG")
        pdf.add_txt(format_bullet_points(payload.get("kham_toan_than", "")))
        if payload.get("loai_benh_an") == "Hậu phẫu":
            pdf.add_txt(f"- Ngày hậu phẫu: {payload.get('ngay_hau_phau', '')}\n- Vết mổ: {payload.get('kham_vet_mo', '')}\n- Ống dẫn lưu: {payload.get('kham_dan_luu', '')}")

        pdf.add_sec("VI. CHẨN ĐOÁN SƠ BỘ & PHÂN BIỆT")
        pdf.add_txt(f"- Sơ bộ: {payload.get('chan_doan_so_bo', '')}\n- Phân biệt: {payload.get('chan_doan_phan_biet', '')}\n- Biện luận: {payload.get('bien_luan', '')}")

        pdf.add_sec("VII. TÓM TẮT BỆNH ÁN")
        pdf.add_txt(payload.get("tom_tat", ""))

        pdf.add_sec("VIII. CHẨN ĐOÁN XÁC ĐỊNH")
        pdf.add_txt(payload.get("chan_doan_xac_dinh", ""))

        pdf.add_sec("IX. ĐIỀU TRỊ & TIÊN LƯỢNG")
        dt = f"- Mục tiêu: {payload.get('dt_muc_tieu', '')}\n- Cụ thể: {payload.get('dt_cu_the', '')}\n- Theo dõi: {payload.get('dt_theo_doi', '')}\n- Tiên lượng: {payload.get('tien_luong', '')}"
        pdf.add_txt(dt)

        pdf_bytes = bytes(pdf.output())
        return Response(content=pdf_bytes, media_type="application/pdf", headers={"Content-Disposition": "inline; filename=benhan.pdf"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi kết xuất PDF: {str(e)}")

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
