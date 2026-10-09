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

# Thư viện đọc PDF để bóc tách ảnh nét cao
try:
    import fitz  # Tên gọi gọn của PyMuPDF
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
if "units" not in st.session_state: 
    st.session_state.units = [{"id": 1, "name": "Đơn vị 1"}]
if "photo_cart" not in st.session_state: 
    st.session_state.photo_cart = {1: {}} 

if "template_bytes" not in st.session_state: 
    st.session_state.template_bytes = None 
if "cover_pos" not in st.session_state: 
    st.session_state.cover_pos = "Dưới - Giữa"
if "ty_pos" not in st.session_state: 
    st.session_state.ty_pos = "Trung tâm"

if "final_pdf" not in st.session_state: 
    st.session_state.final_pdf = None
if "final_pptx" not in st.session_state: 
    st.session_state.final_pptx = None
if "show_download_pdf" not in st.session_state: 
    st.session_state.show_download_pdf = False
if "show_download_pptx" not in st.session_state: 
    st.session_state.show_download_pptx = False

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def process_bg_image(uploaded_file, target_ratio=16/9):
    if not uploaded_file: 
        return None
    try:
        bg_bytes = uploaded_file.getvalue()
        with Image.open(io.BytesIO(bg_bytes)) as img:
            img = ImageOps.exif_transpose(img)
            if img.mode != 'RGB': 
                img = img.convert('RGB')
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
    except Exception: 
        return None

def optimize_and_extract_info(img_bytes, file_name):
    try:
        with Image.open(io.BytesIO(img_bytes)) as img:
            dt_str = "9999"
            exif = img.getexif()
            if exif: 
                dt_str = exif.get(36867) or exif.get(306) or "9999"
            
            img_t = ImageOps.exif_transpose(img)
            if img_t.mode != 'RGB': 
                img_t = img_t.convert('RGB')
            
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
    except Exception: 
        return None

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
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        for p_idx in range(len(doc)):
            page = doc[p_idx]
            image_list = page.get_images(full=True)
            for img_idx, img_info in enumerate(image_list):
                xref = img_info[0]
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                image_ext = base_image["ext"]
                
                # Bỏ qua ảnh icon bé tí (dưới 10KB)
                if len(image_bytes) > 10240: 
                    ext_name = f"{base_name}_page{p_idx}_img{img_idx}.{image_ext}"
                    extracted.append((ext_name, image_bytes))
    except Exception as e:
        st.error(f"Lỗi đọc PDF (Đảm bảo file requirements.txt có chữ PyMuPDF): {e}")
    return extracted

def add_image_exact(slide, img_stream, left, top, width, height):
    img_stream.seek(0)
    pic = slide.shapes.add_picture(img_stream, int(left), int(top), width=int(width), height=int(height))
    pic.line.color.rgb = RGBColor(30, 30, 30)
    pic.line.width = Inches(0.03) 

def move_slide(prs, old_index, new_index):
    if old_index == new_index: 
        return
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    slide_to_move = slides[old_index]
    xml_slides.remove(slide_to_move)
    xml_slides.insert(new_index, slide_to_move)

def draw_adaptive_grid(slide, layout_rows, start_x_base, start_y_base, usable_w, usable_h, GAP):
    num_rows = len(layout_rows)
    if num_rows == 0: 
        return
    H_max = (usable_h - (num_rows - 1) * GAP) / num_rows
    H_final = H_max
    for row in layout_rows:
        if not row: 
            continue
        sum_ratios = sum(img['w'] / img['h'] for img in row)
        H_row_max = (usable_w - (len(row) - 1) * GAP) / sum_ratios
        if H_row_max < H_final: 
            H_final = H_row_max
    Total_H = num_rows * H_final + (num_rows - 1) * GAP
    current_y = start_y_base + (usable_h - Total_H) / 2
    for row in layout_rows:
        if not row: 
            continue
        row_w = sum(H_final * (img['w'] / img['h']) for img in row) + (len(row) - 1) * GAP
        current_x = start_x_base + (usable_w - row_w) / 2
        for img in row:
            img_w = H_final * (img['w'] / img['h'])
            add_image_exact(slide, img['stream'], current_x, current_y, img_w, H_final)
            current_x += img_w + GAP
        current_y += H_final + GAP

def partition_images(imgs, max_size):
    if not imgs: 
        return []
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

def emu_to_px(emu): 
    return int((emu
