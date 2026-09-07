import hashlib
import io
import json
import os
import random
import re
import smtplib
import unicodedata
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

app = FastAPI(title="Bệnh Án Lâm Sàng Win2K")
templates = Jinja2Templates(directory="templates")

# Cấu hình biến môi trường
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "")
SENDER_APP_PASSWORD = os.getenv("SENDER_APP_PASSWORD", "")
ADMIN_EMAIL = os.getenv("ADMIN_EMAIL", "")
APP_PASSWORD = os.getenv("APP_PASSWORD", "123456")
AUTH_SECRET_KEY = os.getenv("AUTH_SECRET_KEY", "clinical_secret_2026")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

# --- TẢI FONT UNICODE TIẾNG VIỆT TỪ NGUỒN CHÍNH THỨC ---
FONT_REGULAR = "Roboto-Regular.ttf"
FONT_BOLD = "Roboto-Bold.ttf"

def download_fonts_if_missing():
    # Sử dụng link raw chính thức từ GitHub Google Fonts
    urls = {
        FONT_REGULAR: "https://raw.githubusercontent.com/google/fonts/main/apache/roboto/Roboto-Regular.ttf",
        FONT_BOLD: "https://raw.githubusercontent.com/google/fonts/main/apache/roboto/Roboto-Bold.ttf"
    }
    for filename, url in urls.items():
        if not os.path.exists(filename) or os.path.getsize(filename) < 10000:
            try:
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=15) as resp, open(filename, 'wb') as f:
                    f.write(resp.read())
                print(f"✅ Đã tải thành công font: {filename}")
            except Exception as e:
                print(f"⚠️ Chưa thể tải {filename}: {e}")

download_fonts_if_missing()

def get_ai_model(model_name: str = "gemini-3.1-flash-lite"):
    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    if not api_key:
        return None
    try:
        genai.configure(api_key=api_key)
        return genai.GenerativeModel(model_name)
    except Exception as e:
        print(f"Lỗi khởi tạo AI: {e}")
        return None

def strip_accents(text: Any) -> str:
    """Khử dấu an toàn dự phòng khi bắt buộc dùng font Latinh"""
    if not text:
        return ""
    text = unicodedata.normalize('NFD', str(text))
    text = re.sub(r'[\u0300-\u036f]', '', text)
    return text.replace('đ', 'd').replace('Đ', 'D')

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
        raise HTTPException(status_code=500, detail="Chưa cài đặt GEMINI_API_KEY trong Environment của Render!")
    
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
    1. Danh sách Chẩn đoán phân biệt (Differential Diagnosis)
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
        raise HTTPException(status_code=500, detail=f"Lỗi Gemini: {str(e)}")

@app.post("/api/ai/treatment")
async def api_ai_treatment(payload: Dict[str, Any]):
    model = get_ai_model()
    if not model:
        raise HTTPException(status_code=500, detail="Chưa cài đặt GEMINI_API_KEY!")
    context = f"Loại: {payload.get('loai_benh_an')}\nChẩn đoán: {payload.get('chan_doan_xac_dinh')}\nTiền sử: {payload.get('ts_noi_khoa')}"
    prompt = f"Bạn là bác sĩ điều trị. Xây dựng phác đồ cho ca bệnh, trả lời ngắn gọn, thẳng vấn đề, ở dạng xuống dòng, chữ đầu viết hoa đơn giản. ({context}). Trả về ĐÚNG 3 thẻ: [MUC_TIEU], [DIEU_TRI_CU_THE], [THEO_DOI]."
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
        raise HTTPException(status_code=500, detail="Chưa cài đặt GEMINI_API_KEY!")
    context = f"Chẩn đoán: {payload.get('chan_doan_xac_dinh')}\nĐiều trị: {payload.get('dt_cu_the')}"
    prompt = f"Bạn là bác sĩ lâm sàng. Đưa ra TIÊN LƯỢNG và TƯ VẤN cho ca bệnh, vào thẳng vấn đề, đơn giản, dạng xuống dòng. ({context}). Trả về 2 thẻ: [TIEN_LUONG] và [TU_VAN]."
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
        raise HTTPException(status_code=500, detail="Chưa cài đặt GEMINI_API_KEY!")
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
        raise HTTPException(status_code=500, detail="Chưa cài đặt GEMINI_API_KEY!")
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

# --- BỘ TẠO PDF TIẾNG VIỆT ĐẢO TRẬT TỰ CHUẨN XÁC NỘI KHOA VS HẬU PHẪU ---
class RobustUnicodePDF(FPDF):
    def __init__(self, loai_ba="Nội khoa / Tiền phẫu", *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.loai_ba = loai_ba
        self.use_unicode = False
        if os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD):
            try:
                self.add_font("Roboto", "", FONT_REGULAR)
                self.add_font("Roboto", "B", FONT_BOLD)
                self.font_family_name = "Roboto"
                self.use_unicode = True
            except Exception:
                self.font_family_name = "Helvetica"
        else:
            self.font_family_name = "Helvetica"

    def clean_text(self, text: Any) -> str:
        s = str(text or "")
        if self.use_unicode:
            return s
        return strip_accents(s)

    def header(self):
        if self.page_no() == 1:
            self.set_font(self.font_family_name, "B" if not self.use_unicode else "", 15)
            tieu_de = "BỆNH ÁN HẬU PHẪU" if self.loai_ba == "Hậu phẫu" else "BỆNH ÁN LÂM SÀNG"
            self.cell(0, 8, self.clean_text(tieu_de), align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_font(self.font_family_name, "", 9)
            self.cell(0, 4, self.clean_text(f"Thời gian lập: {datetime.now().strftime('%d/%m/%Y %H:%M')}"), align="C", new_x="LMARGIN", new_y="NEXT")
            self.ln(4)

    def add_sec(self, title: str):
        self.set_font(self.font_family_name, "B" if not self.use_unicode else "", 11)
        self.set_fill_color(225, 235, 245)
        self.cell(0, 7, self.clean_text(title), fill=True, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    # 1. Cập nhật hàm in tiêu đề phụ: ép kiểu chữ đậm ("B")
    def add_subsec(self, title: str):
        # Thiết lập font chữ in đậm và tăng nhẹ kích thước lên 10pt
        self.set_font(self.font_family_name, "B", 10)
        self.cell(0, 6, self.clean_text(title), new_x="LMARGIN", new_y="NEXT")

    # 2. Đảm bảo hàm in nội dung luôn trả về font chữ thường (Regular - "")
    def add_txt(self, text: str):
        self.set_font(self.font_family_name, "", 9.5)
        self.multi_cell(0, 5, self.clean_text(text) if str(text).strip() else self.clean_text("Chưa ghi nhận thông tin."))
        self.ln(2)

    def render_table_cls(self, cls_rows):
        col_w = (self.w - self.l_margin - self.r_margin) / 2.0
        line_h = 5.0
        self.set_font(self.font_family_name, "B" if not self.use_unicode else "", 9.5)
        self.set_fill_color(230, 235, 245)
        if self.get_y() > 260:
            self.add_page()
        self.cell(col_w, 7, self.clean_text("KẾT QUẢ CẬN LÂM SÀNG"), border=1, align="C", fill=True)
        self.cell(col_w, 7, self.clean_text("PHIÊN GIẢI / BIỆN GIẢI"), border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
        
        self.set_font(self.font_family_name, "", 9)
        for kq, pg in cls_rows:
            txt_kq = format_bullet_points(kq) if kq else "-"
            txt_pg = format_bullet_points(pg) if pg else "-"
            
            # Tính chiều cao dòng
            nb_l = len(self.multi_cell(col_w - 4, line_h, self.clean_text(txt_kq), dry_run=True, output="LINES"))
            nb_r = len(self.multi_cell(col_w - 4, line_h, self.clean_text(txt_pg), dry_run=True, output="LINES"))
            row_h = max(max(nb_l, nb_r) * line_h + 4, 8)
            
            if self.get_y() + row_h > 275:
                self.add_page()
                self.set_font(self.font_family_name, "B" if not self.use_unicode else "", 9.5)
                self.set_fill_color(230, 235, 245)
                self.cell(col_w, 7, self.clean_text("KẾT QUẢ CẬN LÂM SÀNG"), border=1, align="C", fill=True)
                self.cell(col_w, 7, self.clean_text("PHIÊN GIẢI / BIỆN GIẢI"), border=1, align="C", fill=True, new_x="LMARGIN", new_y="NEXT")
                self.set_font(self.font_family_name, "", 9)

            curr_x = self.get_x()
            curr_y = self.get_y()
            self.rect(curr_x, curr_y, col_w, row_h)
            self.rect(curr_x + col_w, curr_y, col_w, row_h)

            self.set_xy(curr_x + 2, curr_y + 2)
            self.multi_cell(col_w - 4, line_h, self.clean_text(txt_kq))
            self.set_xy(curr_x + col_w + 2, curr_y + 2)
            self.multi_cell(col_w - 4, line_h, self.clean_text(txt_pg))
            self.set_xy(curr_x, curr_y + row_h)
        self.ln(3)

@app.post("/api/export/pdf")
async def api_export_pdf(payload: Dict[str, Any]):
    try:
        if not os.path.exists(FONT_REGULAR):
            download_fonts_if_missing()

        loai_ba = str(payload.get("loai_benh_an", "Nội khoa / Tiền phẫu")).strip()
        is_hau_phau = (loai_ba == "Hậu phẫu")
        
        pdf = RobustUnicodePDF(loai_ba=loai_ba)
        pdf.add_page()

        # I. HÀNH CHÍNH
        pdf.add_sec("I. PHẦN HÀNH CHÍNH")
        hc = (
            f"- Họ và tên: {str(payload.get('ho_ten', '')).upper()}   |   Tuổi: {payload.get('tuoi')}   |   Giới tính: {payload.get('gioi_tinh')}\n"
            f"- Dân tộc: {payload.get('dan_tok')}   |   Nghề nghiệp: {payload.get('nghe_nghiep')}\n"
            f"- Khoa phòng: {payload.get('khoa_phong')}   |   Địa chỉ: {payload.get('dia_chi', '')}\n"
            f"- Ngày giờ vào viện: {payload.get('ngay_vao_vien')}   |   Ngày làm BA: {payload.get('ngay_lam_benh_an', '')}\n"
            f"- Người làm bệnh án: {payload.get('sinh_vien')}"
        )
        pdf.add_txt(hc)

        # II. LÝ DO VÀO VIỆN
        pdf.add_sec("II. LÝ DO VÀO VIỆN")
        pdf.add_txt(payload.get("ly_do_vao_vien", ""))

        # III. BỆNH SỬ
        pdf.add_sec("III. BỆNH SỬ")
        if is_hau_phau:
            pdf.add_subsec("1. Tình trạng trước mổ:")
            pdf.add_txt(format_bullet_points(payload.get("bs_truoc_mo", "")))
            pdf.add_subsec("2. Tình trạng trong mổ:")
            pdf.add_txt(format_bullet_points(payload.get("bs_trong_mo", "")))
            pdf.add_subsec("3. Quá trình sau mổ:")
            pdf.add_txt(format_bullet_points(payload.get("bs_sau_mo", "")))
        else:
            pdf.add_txt(payload.get("benh_su", ""))

        # IV. TIỀN SỬ
        pdf.add_sec("IV. TIỀN SỬ")
        pdf.add_subsec("1. Tiền sử nội khoa:")
        pdf.add_txt(format_bullet_points(payload.get("ts_noi_khoa", "")))
        pdf.add_subsec("2. Tiền sử ngoại khoa & Dị ứng:")
        pdf.add_txt(format_bullet_points(payload.get("ts_ngoai_khoa", "")))
        pdf.add_subsec("3. Tiền sử bản thân (Lối sống & Thói quen):")
        pdf.add_txt(format_bullet_points(payload.get("ts_loi_song", "")))
        pdf.add_subsec("4. Tiền sử gia đình:")
        pdf.add_txt(format_bullet_points(payload.get("ts_gia_dinh", "")))

        # V. THĂM KHÁM LÂM SÀNG
        pdf.add_sec("V. THĂM KHÁM LÂM SÀNG")
        if not is_hau_phau:
            pdf.add_subsec("1. Thăm khám lúc vào viện:")
            pdf.add_txt(format_bullet_points(payload.get("kham_vao_vien", "")))
            pdf.add_subsec("2. Thăm khám hiện tại - Toàn thân:")
        else:
            pdf.add_subsec(f"1. Thăm khám hiện tại ({payload.get('ngay_hau_phau', 'Hậu phẫu')}):")
            pdf.add_subsec("a. Toàn thân:")
            
        pdf.add_txt(format_bullet_points(payload.get("kham_toan_than", "")))
        
        # Sinh hiệu & BMI
        mach = payload.get("sh_mach") or "--"
        nhiet = payload.get("sh_nhiet_do") or "--"
        ha = payload.get("sh_ha") or "--"
        nt = payload.get("sh_nhip_tho") or "--"
        cn = payload.get("sh_can_nang") or "--"
        cc = payload.get("sh_chieu_cao") or "--"
        bmi = payload.get("sh_bmi") or "--"
        eval_bmi = payload.get("sh_bmi_eval") or "--"
        sh_line = f"Sinh hiệu: Mạch: {mach} ck/phút | HA: {ha} mmHg | Nhiệt độ: {nhiet} °C | Nhịp thở: {nt} l/phút\nThể trạng: Chiều cao: {cc} cm | Cân nặng: {cn} kg | BMI: {bmi} kg/m² ({eval_bmi})"
        pdf.add_txt(sh_line)

        # Nếu là Hậu phẫu: In thêm vết mổ & dẫn lưu
        if is_hau_phau:
            pdf.add_subsec("b. Vết mổ & Dẫn lưu:")
            pdf.add_txt(f"- Vết mổ: {payload.get('kham_vet_mo', '')}\n- Dẫn lưu: {payload.get('kham_dan_luu', '')}")
            pdf.add_subsec("c. Các cơ quan:")
        else:
            pdf.add_subsec("3. Thăm khám hiện tại - Các cơ quan:")

        # In 7 cơ quan theo thứ tự ưu tiên
        organs = [
            ("Tuần hoàn", "kham_tuan_hoan"),
            ("Hô hấp", "kham_ho_hap"),
            ("Tiêu hóa", "kham_tieu_hoa"),
            ("Thần kinh", "kham_than_kinh"),
            ("Thận - Tiết niệu", "kham_tiet_nieu"),
            ("Cơ xương khớp", "kham_co_xuong_khop"),
            ("Các cơ quan khác", "kham_co_quan_khac")
        ]
        fav_key = payload.get("uu_tien_co_quan", "none")
        if fav_key != "none":
            fav = [o for o in organs if o[1] == fav_key]
            others = [o for o in organs if o[1] != fav_key]
            organs = fav + others

        for name, key in organs:
            pdf.add_subsec(f"- {name}:")
            pdf.add_txt(format_bullet_points(payload.get(key, "")))

        # ĐỊNH NGHĨA HÀM KHỐI CHO CÁC MỤC LOGIC
        def sec_tom_tat(num_rom):
            pdf.add_sec(f"{num_rom}. TÓM TẮT BỆNH ÁN")
            pdf.add_txt(payload.get("tom_tat", ""))

        def sec_chan_doan_so_bo(num_sb, num_pb, num_bl):
            pdf.add_sec(f"{num_sb}. CHẨN ĐOÁN SƠ BỘ")
            pdf.add_txt(payload.get("chan_doan_so_bo", ""))
            pdf.add_sec(f"{num_pb}. CHẨN ĐOÁN PHÂN BIỆT")
            pdf.add_txt(payload.get("chan_doan_phan_biet", ""))
            if payload.get("bien_luan"):
                pdf.add_sec(f"{num_bl}. BIỆN LUẬN CHẨN ĐOÁN SƠ BỘ")
                pdf.add_txt(payload.get("bien_luan", ""))

        def sec_can_lam_sang(num_dx, num_co):
            pdf.add_sec(f"{num_dx}. ĐỀ XUẤT CẬN LÂM SÀNG")
            label_1 = "1. Đánh giá sau mổ / Biến chứng:" if is_hau_phau else "1. Phục vụ chẩn đoán xác định:"
            label_2 = "2. Theo dõi hồi phục & Chăm sóc:" if is_hau_phau else "2. Phục vụ điều trị:"
            pdf.add_subsec(label_1)
            pdf.add_txt(format_bullet_points(payload.get("cls_dx_xac_dinh", "")))
            pdf.add_subsec(label_2)
            pdf.add_txt(format_bullet_points(payload.get("cls_dx_dieu_tri", "")))
            pdf.add_subsec("3. Cận lâm sàng khác:")
            pdf.add_txt(format_bullet_points(payload.get("cls_dx_khac", "")))

            pdf.add_sec(f"{num_co}. CẬN LÂM SÀNG ĐÃ CÓ")
            cls_rows = []
            so_hang = int(payload.get("so_hang_cls", 3))
            for i in range(so_hang):
                kq = payload.get(f"cls_kq_{i}", "").strip()
                pg = payload.get(f"cls_pg_{i}", "").strip()
                if kq or pg:
                    cls_rows.append((kq, pg))
            if cls_rows:
                pdf.render_table_cls(cls_rows)
            else:
                pdf.add_txt("Chưa ghi nhận kết quả cận lâm sàng.")

        def sec_chan_doan_xac_dinh(num_xd, num_blxd):
            pdf.add_sec(f"{num_xd}. CHẨN ĐOÁN XÁC ĐỊNH")
            pdf.add_txt(format_bullet_points(payload.get("chan_doan_xac_dinh", "")))
            if payload.get("bien_luan_xac_dinh"):
                pdf.add_sec(f"{num_blxd}. BIỆN LUẬN CHẨN ĐOÁN XÁC ĐỊNH")
                pdf.add_txt(format_bullet_points(payload.get("bien_luan_xac_dinh", "")))

        # PHÂN NHÁNH TRẬT TỰ SỐ LA MÃ CHUẨN XÁC TUYỆT ĐỐI
        if not is_hau_phau:
            # TRẬT TỰ NỘI KHOA: Tóm tắt -> CĐ Sơ bộ -> CLS -> CĐ Xác định
            sec_tom_tat("VI")
            sec_chan_doan_so_bo("VII", "VIII", "IX")
            sec_can_lam_sang("X", "XI")
            sec_chan_doan_xac_dinh("XII", "XIII")
            pdf.add_sec("XIV. ĐIỀU TRỊ")
            pdf.add_subsec("1. Mục tiêu điều trị:")
            pdf.add_txt(format_bullet_points(payload.get("dt_muc_tieu", "")))
            pdf.add_subsec("2. Điều trị cụ thể:")
            pdf.add_txt(format_bullet_points(payload.get("dt_cu_the", "")))
            pdf.add_subsec("3. Theo dõi sau điều trị:")
            pdf.add_txt(format_bullet_points(payload.get("dt_theo_doi", "")))
            pdf.add_sec("XV. TIÊN LƯỢNG")
            pdf.add_txt(format_bullet_points(payload.get("tien_luong", "")))
            pdf.add_sec("XVI. TƯ VẤN")
            pdf.add_txt(format_bullet_points(payload.get("tu_van", "")))
        else:
            # TRẬT TỰ HẬU PHẪU: CĐ Sơ bộ -> CLS -> Tóm tắt -> CĐ Xác định
            sec_chan_doan_so_bo("VI", "VII", "VIII")
            sec_can_lam_sang("IX", "X")
            sec_tom_tat("XI")
            sec_chan_doan_xac_dinh("XII", "XIII")
            pdf.add_sec("XIV. ĐIỀU TRỊ HẬU PHẪU")
            pdf.add_subsec("1. Mục tiêu điều trị:")
            pdf.add_txt(format_bullet_points(payload.get("dt_muc_tieu", "")))
            pdf.add_subsec("2. Điều trị cụ thể:")
            pdf.add_txt(format_bullet_points(payload.get("dt_cu_the", "")))
            pdf.add_subsec("3. Theo dõi sau điều trị:")
            pdf.add_txt(format_bullet_points(payload.get("dt_theo_doi", "")))
            pdf.add_sec("XV. TIÊN LƯỢNG")
            pdf.add_txt(format_bullet_points(payload.get("tien_luong", "")))
            pdf.add_sec("XVI. TƯ VẤN")
            pdf.add_txt(format_bullet_points(payload.get("tu_van", "")))

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
