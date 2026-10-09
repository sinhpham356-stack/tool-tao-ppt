import io
import re
import math
import os
import urllib.request
import gc
from datetime import datetime
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image, ImageOps, ImageDraw, ImageFont

# Thư viện đọc PDF để bóc tách ảnh
try:
    from pypdf import PdfReader
except ImportError:
    pass

# --- SỬA LỖI ĐỌC ẢNH IPHONE (HEIC) ---
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

st.set_page_config(page_title="Công Cụ Chèn Ảnh PowerPoint & PDF", page_icon="⚡", layout="centered")

st.markdown("""
    <style>
        div[data-testid="stFileUploadDropzone"] > div > small {
            display: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# Khởi tạo Session State hỗ trợ nhiều Đơn vị
if "units" not in st.session_state: st.session_state.units = [{"id": 1, "name": "Đơn vị 1"}]
if "photo_cart" not in st.session_state: st.session_state.photo_cart = {1: {}} 

if "template_bytes" not in st.session_state: st.session_state.template_bytes = None 
if "cover_pos" not in st.session_state: st.session_state.cover_pos = "Dưới - Giữa"
if "ty_pos" not in st.session_state: st.session_state.ty_pos = "Trung tâm"

if "final_pdf" not in st.session_state: st.session_state.final_pdf = None
if "final_pptx" not in st.session_state: st.session_state.final_pptx = None
if "show_download_pdf" not in st.session_state: st.session_state.show_download_pdf = False
if "show_download_pptx" not in st.session_state: st.session_state.show_download_pptx = False

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def process_bg_image(uploaded_file, target_ratio=16/9):
    if not uploaded_file: return None
    try:
        bg_bytes = uploaded_file.getvalue()
        with Image.open(io.BytesIO(bg_bytes)) as img:
            img = ImageOps.exif_transpose(img)
            if img.mode != 'RGB': img = img.convert('RGB')
            img_ratio = img.width / img.height
            if img_ratio > target_ratio:
                new_w = int(img.height * target_ratio)
                offset = (img.width - new_w) // 2
                img = img.crop((offset, 0, offset + new_w, img.height))
            elif img_ratio < target_ratio:
                new_h = int(img.width / target_ratio)
                offset = (img.height - new_h) // 2
                img = img.crop((0, offset, img.width, offset + new_h))
            bg_stream = io.BytesIO()
            img.save(bg_stream, format='JPEG', quality=90)
            return bg_stream
    except Exception: return None

def optimize_and_extract_info(img_bytes, file_name):
    try:
        with Image.open(io.BytesIO(img_bytes)) as img:
            dt_str = "9999"
            exif = img.getexif()
            if exif: dt_str = exif.get(36867) or exif.get(306) or "9999"
            
            img_t = ImageOps.exif_transpose(img)
            if img_t.mode != 'RGB': img_t = img_t.convert('RGB')
            
            img_t.thumbnail((1600, 1600), Image.LANCZOS)
            w, h = img_t.size
            is_portrait = h >= w
            
            opt_stream = io.BytesIO()
            img_t.save(opt_stream, format='JPEG', quality=85)
            
            return {
                "bytes": opt_stream.getvalue(),
                "name": file_name,
                "timestamp": str(dt_str),
                "w": w, "h": h, "is_portrait": is_portrait
            }
    except Exception: return None

# --- HÀM BÓC TÁCH ẢNH TỪ PPTX VÀ PDF ---
def extract_images_from_pptx(file_bytes, base_name):
    extracted = []
    try:
        prs_ext = Presentation(io.BytesIO(file_bytes))
        for s_idx, slide in enumerate(prs_ext.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if hasattr(shape, "image"):
                    ext_name = f"{base_name}_slide{s_idx}_img{shape_idx}.png"
                    extracted.append((ext_name, shape.image.blob))
    except Exception as e:
        st.error(f"Lỗi đọc PPTX: {e}")
    return extracted

def extract_images_from_pdf(file_bytes, base_name):
    extracted = []
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        for p_idx, page in enumerate(reader.pages):
            for img_obj in page.images:
                ext_name = f"{base_name}_page{p_idx}_{img_obj.name}"
                extracted.append((ext_name, img_obj.data))
    except Exception as e:
        st.error(f"Lỗi đọc PDF (Hãy chắc chắn bạn đã cài thư viện pypdf): {e}")
    return extracted

def add_image_exact(slide, img_stream, left, top, width, height):
    img_stream.seek(0)
    pic = slide.shapes.add_picture(img_stream, int(left), int(top), width=int(width), height=int(height))
    pic.line.color.rgb = RGBColor(30, 30, 30)
    pic.line.width = Inches(0.03) 

def move_slide(prs, old_index, new_index):
    if old_index == new_index: return
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    slide_to_move = slides[old_index]
    xml_slides.remove(slide_to_move)
    xml_slides.insert(new_index, slide_to_move)

def draw_adaptive_grid(slide, layout_rows, start_x_base, start_y_base, usable_w, usable_h, GAP):
    num_rows = len(layout_rows)
    if num_rows == 0: return
    H_max = (usable_h - (num_rows - 1) * GAP) / num_rows
    H_final = H_max
    for row in layout_rows:
        if not row: continue
        sum_ratios = sum(img['w'] / img['h'] for img in row)
        H_row_max = (usable_w - (len(row) - 1) * GAP) / sum_ratios
        if H_row_max < H_final: H_final = H_row_max
    Total_H = num_rows * H_final + (num_rows - 1) * GAP
    current_y = start_y_base + (usable_h - Total_H) / 2
    for row in layout_rows:
        if not row: continue
        row_w = sum(H_final * (img['w'] / img['h']) for img in row) + (len(row) - 1) * GAP
        current_x = start_x_base + (usable_w - row_w) / 2
        for img in row:
            img_w = H_final * (img['w'] / img['h'])
            add_image_exact(slide, img['stream'], current_x, current_y, img_w, H_final)
            current_x += img_w + GAP
        current_y += H_final + GAP

def partition_images(imgs, max_size):
    if not imgs: return []
    n = len(imgs)
    num_slides = math.ceil(n / max_size)
    base = n // num_slides
    rem = n % num_slides
    res = []
    idx = 0
    for i in range(num_slides):
        s = base + 1 if i < rem else base
        res.append(imgs[idx:idx+s])
        idx += s
    return res

def emu_to_px(emu): return int((emu / 914400.0) * 300)

def add_pdf_slide(pdf_slides, w_emu, h_emu, bg_stream=None):
    pw = emu_to_px(w_emu)
    ph = emu_to_px(h_emu)
    img = Image.new('RGB', (pw, ph), (255, 255, 255))
    if bg_stream:
        bg_stream.seek(0)
        bg = Image.open(bg_stream).convert('RGB')
        bg = bg.resize((pw, ph), Image.LANCZOS)
        img.paste(bg, (0, 0))
    pdf_slides.append(img)
    return img

def draw_text_pdf(slide_img, text, x_emu, y_emu, w_emu, h_emu, align_pos, pt_size, color_hex, underline=False):
    if not text: return
    draw = ImageDraw.Draw(slide_img)
    px, py, pw = emu_to_px(x_emu), emu_to_px(y_emu), emu_to_px(w_emu)
    font_size = int(pt_size * 300 / 72) 
    
    font_path = "Roboto-Medium.ttf"
    if not os.path.exists(font_path):
        try: urllib.request.urlretrieve("https://github.com/googlefonts/roboto/raw/main/src/hinted/Roboto-Medium.ttf", font_path)
        except: pass
    try: font = ImageFont.truetype(font_path, font_size)
    except: font = ImageFont.load_default()
        
    r, g, b = hex_to_rgb(color_hex)
    words = text.split(); lines = []; current_line = ""
    for word in words:
        test_line = current_line + " " + word if current_line else word
        bbox = draw.textbbox((0,0), test_line, font=font)
        if bbox[2] - bbox[0] <= pw: current_line = test_line
        else:
            if current_line: lines.append(current_line)
            current_line = word
    if current_line: lines.append(current_line)
        
    current_y = py
    for line in lines:
        bbox = draw.textbbox((0,0), line, font=font)
        tw = bbox[2] - bbox[0]
        text_bottom = bbox[3] 
        th = text_bottom - bbox[1]
        
        if "Trái" in align_pos or align_pos == PP_ALIGN.LEFT: tx = px
        elif "Phải" in align_pos or align_pos == PP_ALIGN.RIGHT: tx = px + pw - tw
        else: tx = px + (pw - tw) / 2
            
        draw.text((tx, current_y), line, font=font, fill=(r,g,b))
        if underline:
            line_y = current_y + text_bottom + 15
            draw.line([tx, line_y, tx + tw, line_y], fill=(r,g,b), width=6)
        current_y += th + 40 

def add_image_pdf(slide_img, img_stream, left_emu, top_emu, width_emu, height_emu):
    img_stream.seek(0)
    img = Image.open(img_stream).convert('RGB')
    px, py = emu_to_px(left_emu), emu_to_px(top_emu)
    pw, ph = emu_to_px(width_emu), emu_to_px(height_emu)
    img = img.resize((pw, ph), Image.LANCZOS)
    slide_img.paste(img, (px, py))
    draw = ImageDraw.Draw(slide_img)
    draw.rectangle([px, py, px+pw, py+ph], outline=(30,30,30), width=6)

def draw_adaptive_grid_pdf(slide_img, layout_rows, start_x_base, start_y_base, usable_w, usable_h, GAP):
    num_rows = len(layout_rows)
    if num_rows == 0: return
    H_max = (usable_h - (num_rows - 1) * GAP) / num_rows
    H_final = H_max
    for row in layout_rows:
        if not row: continue
        sum_ratios = sum(img['w'] / img['h'] for img in row)
        H_row_max = (usable_w - (len(row) - 1) * GAP) / sum_ratios
        if H_row_max < H_final: H_final = H_row_max
    Total_H = num_rows * H_final + (num_rows - 1) * GAP
    current_y = start_y_base + (usable_h - Total_H) / 2
    for row in layout_rows:
        if not row: continue
        row_w = sum(H_final * (img['w'] / img['h']) for img in row) + (len(row) - 1) * GAP
        current_x = start_x_base + (usable_w - row_w) / 2
        for img in row:
            img_w = H_final * (img['w'] / img['h'])
            add_image_pdf(slide_img, img['stream'], current_x, current_y, img_w, H_final)
            current_x += img_w + GAP
        current_y += H_final + GAP

def render_position_grid(state_key, key_prefix):
    def set_pos(pos): st.session_state[state_key] = pos
    current_pos = st.session_state[state_key]
    grid_config = [
        [("Trên - Trái", "↖️"), ("Trên - Giữa", "⬆️"), ("Trên - Phải", "↗️")],
        [("Giữa - Trái", "⬅️"), ("Trung tâm", "⏺️"), ("Giữa - Phải", "➡️")],
        [("Dưới - Trái", "↙️"), ("Dưới - Giữa", "⬇️"), ("Dưới - Phải", "↘️")]
    ]
    st.markdown(f"📍 Đang định vị: **{current_pos}**")
    for r_idx, row in enumerate(grid_config):
        cols = st.columns(3, gap="small")
        for c_idx, (pos_name, icon) in enumerate(row):
            btn_label = f"✅ {icon}" if current_pos == pos_name else icon
            cols[c_idx].button(
                btn_label, key=f"{key_prefix}_{r_idx}_{c_idx}", 
                use_container_width=True, on_click=set_pos, args=(pos_name,)
            )

def get_alignment(pos_str):
    if "Trái" in pos_str: return PP_ALIGN.LEFT
    elif "Phải" in pos_str: return PP_ALIGN.RIGHT
    else: return PP_ALIGN.CENTER

def get_cover_y(pos_str, slide_h, is_sub=False):
    if "Trên" in pos_str: return Inches(0.5) if not is_sub else Inches(1.5)
    elif "Dưới" in pos_str: return slide_h - Inches(2.2) if not is_sub else slide_h - Inches(1.2)
    else: return slide_h/2 - Inches(1.0) if not is_sub else slide_h/2 + Inches(0.2)

def get_ty_y(pos_str, slide_h):
    if "Trên" in pos_str: return Inches(0.8)
    elif "Dưới" in pos_str: return slide_h - Inches(1.8)
    else: return slide_h/2 - Inches(0.5)

# ==========================================
# GIAO DIỆN HIỂN THỊ
# ==========================================
st.title("⚡ TRỢ LÝ TẠO REPORT BẰNG HÌNH ẢNH")

st.header("Bước 1: Tải lên File PowerPoint Mẫu (Không bắt buộc)")
template_file = st.file_uploader("Chọn file .pptx (Chỉ cần up 1 lần)", type=["pptx"])
if template_file: st.session_state.template_bytes = template_file.getvalue()
if st.session_state.template_bytes:
    st.success("✅ Đã lưu file PPTX mẫu vào bộ nhớ!")
    if st.button("🗑️ Xóa file mẫu (Để tạo file trắng mới)"):
        st.session_state.template_bytes = None
        st.rerun()

st.header("Bước 2: Phân loại Ảnh Theo Từng Đơn Vị / Hạng Mục")
st.info("Mỗi Đơn vị sẽ được tự động gom thành 1 phần trong báo cáo. Trang đầu tiên của Đơn vị đó sẽ có tên, các trang sau tự động ẩn tên để gọn gàng.")

if st.button("➕ THÊM ĐƠN VỊ / HẠNG MỤC MỚI", use_container_width=True):
    new_id = max([u["id"] for u in st.session_state.units], default=0) + 1
    st.session_state.units.append({"id": new_id, "name": f"Đơn vị {new_id}"})
    st.session_state.photo_cart[new_id] = {}
    st.rerun()

# --- Hiển thị danh sách các Đơn vị ---
total_images = 0
for idx, unit in enumerate(st.session_state.units):
    uid = unit["id"]
    if uid not in st.session_state.photo_cart:
        st.session_state.photo_cart[uid] = {}
        
    num_imgs = len(st.session_state.photo_cart[uid])
    total_images += num_imgs
    
    with st.expander(f"📁 {unit['name']} ({num_imgs} ảnh đã tải)", expanded=True):
        col1, col2 = st.columns([4, 1])
        with col1:
            new_name = st.text_input("Tên hiển thị trên Báo Cáo:", value=unit['name'], key=f"name_{uid}")
            st.session_state.units[idx]['name'] = new_name
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🗑️ Xóa", key=f"del_{uid}"):
                st.session_state.units.pop(idx)
                if uid in st.session_state.photo_cart: del st.session_state.photo_cart[uid]
                st.rerun()

        # NÂNG CẤP: Cho phép nhận cả file ảnh và file pdf/pptx vào cùng 1 chỗ
        uploaded_files = st.file_uploader(f"Thêm ảnh hoặc file PDF/PPTX cho {new_name}", accept_multiple_files=True, key=f"up_{uid}", type=['jpg', 'jpeg', 'png', 'heic', 'pdf', 'pptx'])
        
        if uploaded_files:
            count_img = 0
            count_doc = 0
            with st.spinner("Đang hút ảnh và lục soát tài liệu..."):
                for f in uploaded_files:
                    fname = f.name.lower()
                    # Nếu là file tài liệu (PDF/PPTX) -> Bóc tách ảnh
                    if fname.endswith(('.pdf', '.pptx')):
                        extracted_items = []
                        if fname.endswith(".pdf"):
                            extracted_items = extract_images_from_pdf(f.getvalue(), f.name)
                        elif fname.endswith(".pptx"):
                            extracted_items = extract_images_from_pptx(f.getvalue(), f.name)
                            
                        for ext_name, ext_bytes in extracted_items:
                            if ext_name not in st.session_state.photo_cart[uid]:
                                optimized_data = optimize_and_extract_info(ext_bytes, ext_name)
                                if optimized_data:
                                    st.session_state.photo_cart[uid][ext_name] = optimized_data
                                    count_doc += 1
                    # Nếu là file ảnh thường
                    else:
                        if f.name not in st.session_state.photo_cart[uid]:
                            optimized_data = optimize_and_extract_info(f.getvalue(), f.name)
                            if optimized_data:
                                st.session_state.photo_cart[uid][f.name] = optimized_data
                                count_img += 1
                gc.collect() 
            if count_img > 0 or count_doc > 0:
                st.success(f"🎉 Đã thêm {count_img} ảnh thường và bóc tách {count_doc} ảnh từ tài liệu vào {new_name}!")
                st.rerun()
