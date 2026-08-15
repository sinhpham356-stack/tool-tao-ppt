import io
import re
import math
from datetime import datetime
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
from pptx.dml.color import RGBColor
from PIL import Image, ImageOps

# ==========================================
# CẤU HÌNH GIAO DIỆN WEB
# ==========================================
st.set_page_config(page_title="Công Cụ Chèn Ảnh PowerPoint", page_icon="⚡", layout="centered")

# ==========================================
# HÀM LÕI XỬ LÝ PPTX TRÊN WEB
# ==========================================
def add_image_exact(slide, img_stream, left, top, width, height):
    with Image.open(img_stream) as img:
        img = ImageOps.exif_transpose(img) 
        if img.mode != 'RGB':
            img = img.convert('RGB')
            
        temp_stream = io.BytesIO()
        img.save(temp_stream, format='JPEG', quality=95)
        temp_stream.seek(0)
        
    pic = slide.shapes.add_picture(temp_stream, int(left), int(top), width=int(width), height=int(height))
    pic.line.color.rgb = RGBColor(30, 30, 30)
    pic.line.width = Inches(0.03) 

def move_slide(prs, old_index, new_index):
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
        if H_row_max < H_final:
            H_final = H_row_max

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

# ==========================================
# GIAO DIỆN HIỂN THỊ
# ==========================================
st.title("⚡ TRỢ LÝ TẠO REPORT BẰNG HÌNH ẢNH")
st.markdown("Chèn ảnh tự động, giữ đúng thời gian chụp, tràn viền chuẩn form.")

st.header("Bước 1: Tải lên File PowerPoint Mẫu")
template_file = st.file_uploader("Chọn file .pptx", type=["pptx"])

st.header("Bước 2: Chọn Ảnh (Từ Thư Viện Hoặc Chụp Camera)")

# Cho phép chọn file linh hoạt không bị giới hạn định dạng ngặt nghèo
uploaded_images = st.file_uploader(
    "📁 Chọn hoặc bôi đen nhiều ảnh từ Thư viện máy (Hỗ trợ JPG, PNG, WEBP...)", 
    type=["jpg", "jpeg", "png", "webp", "heic"], 
    accept_multiple_files=True
)

st.markdown("---")
st.markdown("📷 **Hoặc chụp ảnh trực tiếp bằng Camera (nếu cần):**")
camera_photo = st.camera_input("Bấm để chụp ảnh thực tế")

all_images_to_process = []

if uploaded_images:
    all_images_to_process.extend(uploaded_images)

if camera_photo:
    if "captured_photos" not in st.session_state:
        st.session_state.captured_photos = []
    if not st.session_state.captured_photos or st.session_state.captured_photos[-1].name != camera_photo.name:
        st.session_state.captured_photos.append(camera_photo)
    
    st.info(f"Đang có {len(st.session_state.captured_photos)} ảnh chụp từ camera.")
    if st.button("🗑️ Xóa danh sách ảnh chụp"):
        st.session_state.captured_photos = []
        st.rerun()
        
    all_images_to_process.extend(st.session_state.captured_photos)

st.header("Bước 3: Tùy Chỉnh Layout")
mode = st.radio("Chọn chế độ:", 
                ("Layout 1: Cơ bản (1-2 ảnh/trang)", 
                 "Layout 2: Tràn viền (Lưới tự động chia đều, nhiều ảnh/trang)"))

align_mode = "2"
if "Layout 1" in mode:
    align_mode = st.radio("Căn lề cho ảnh (Layout 1):", ("1 - Trái", "2 - Giữa", "3 - Phải"))
    align_mode = align_mode[0]

st.header("Bước 4: Vị trí chèn")
vitri_input = st.text_input("Chèn vào sau Slide số mấy? (Gõ 0 để chèn cuối cùng):", "0")

if st.button("🚀 XUẤT FILE POWERPOINT", use_container_width=True, type="primary"):
    if not template_file:
        st.error("⚠️ Vui lòng tải lên file PowerPoint mẫu ở Bước 1!")
    elif not all_images_to_process:
        st.error("⚠️ Vui lòng chọn hoặc chụp ít nhất 1 tấm ảnh ở Bước 2!")
    else:
        with st.spinner("Đang tự động dàn trang... Vui lòng đợi nhé..."):
            try:
                prs = Presentation(template_file)
                tong_slide = len(prs.slides)
                
                vi_tri_hien_tai = tong_slide 
                match = re.search(r'\d+', vitri_input)
                if match:
                    trang_chon = int(match.group())
                    if 0 < trang_chon <= tong_slide:
                        vi_tri_hien_tai = trang_chon
                
                try:
                    slide_layout = prs.slide_layouts[6]
                except:
                    slide_layout = prs.slide_layouts[0] 

                slide_w = prs.slide_width
                slide_h = prs.slide_height

                CACH_LE_TREN = Inches(0.9)  
                CACH_LE_DUOI = Inches(0.8)  
                CACH_LE_TRAI = Inches(0.2) 
                CACH_LE_PHAI = Inches(0.2) 
                GAP = Inches(0.12)          
                
                usable_w = slide_w - CACH_LE_TRAI - CACH_LE_PHAI
                usable_h = slide_h - CACH_LE_TREN - CACH_LE_DUOI

                image_data = []
                for img_file in all_images_to_process:
                    img_bytes = img_file.read()
                    img_stream = io.BytesIO(img_bytes)
                    
                    try:
                        with Image.open(img_stream) as im:
                            im = ImageOps.exif_transpose(im)
                            width, height = im.size
                            is_portrait = height >= width
                            
                            dt_str = "9999"
                            exif = im.getexif()
                            if exif:
                                dt_str = exif.get(36867) or exif.get(306) or "9999"

                        img_stream.seek(0)
                        
                        image_data.append({
                            'stream': img_stream, 
                            'is_portrait': is_portrait, 
                            'w': width, 
                            'h': height,
                            'name': getattr(img_file, 'name', 'photo.jpg'),
                            'timestamp': str(dt_str)
                        })
                    except Exception:
                        pass

                image_data.sort(key=lambda x: (x['timestamp'], x['name']))

                if "Layout 1" in mode:
                    i = 0
                    while i < len(image_data):
                        current_img = image_data[i]
                        slide = prs.slides.add_slide(slide_layout) 
                        move_slide(prs, len(prs.slides) - 1, vi_tri_hien_tai)
                        
                        if current_img['is_portrait'] and (i + 1 < len(image_data)) and image_data[i+1]['is_portrait']:
                            next_img = image_data[i+1]
                            r1 = current_img['w'] / current_img['h']
                            r2 = next_img['w'] / next_img['h']
                            
                            test_w = usable_h * r1 + usable_h * r2
                            if test_w <= usable_w - GAP:
                                final_h = usable_h
                            else:
                                final_h = (usable_w - GAP) / (r1 + r2)
                                
                            final_w1 = final_h * r1
                            final_w2 = final_h * r2
                            block_w = final_w1 + GAP + final_w2
                            
                            if align_mode == '1':
                                start_x = CACH_LE_TRAI
                            elif align_mode == '3':
                                start_x = slide_w - CACH_LE_PHAI - block_w
                            else:
                                start_x = CACH_LE_TRAI + (usable_w - block_w) / 2
                                
                            start_y = CACH_LE_TREN + (usable_h - final_h) / 2
                            
                            add_image_exact(slide, current_img['stream'], start_x, start_y, final_w1, final_h)
                            add_image_exact(slide, next_img['stream'], start_x + final_w1 + GAP, start_y, final_w2, final_h)
                            i += 2 
                        else:
                            r_img = current_img['w'] / current_img['h']
                            if usable_h * r_img <= usable_w:
                                f_h = usable_h
                                f_w = usable_h * r_img
                            else:
                                f_w = usable_w
                                f_h = usable_w / r_img
                                
                            if align_mode == '1':
                                s_x = CACH_LE_TRAI
                            elif align_mode == '3':
                                s_x = slide_w - CACH_LE_PHAI - f_w
                            else:
                                s_x = CACH_LE_TRAI + (usable_w - f_w) / 2
                                
                            s_y = CACH_LE_TREN + (usable_h - f_h) / 2
                            
                            add_image_exact(slide, current_img['stream'], s_x, s_y, f_w, f_h)
                            i += 1
                            
                        vi_tri_hien_tai += 1

                elif "Layout 2" in mode:
                    landscapes = [img for img in image_data if not img['is_portrait']]
                    portraits = [img for img in image_data if img['is_portrait']]

                    land_chunks = partition_images(landscapes, 6) 
                    port_chunks = partition_images(portraits, 4)  

                    smart_chunks = []
                    for c in land_chunks:
                        smart_chunks.append({'type': 'landscape', 'images': c})
                    for c in port_chunks:
                        smart_chunks.append({'type': 'portrait', 'images': c})

                    start_x_base = CACH_LE_TRAI
                    start_y_base = CACH_LE_TREN

                    for chunk_dict in smart_chunks:
                        chunk = chunk_dict['images']
                        c_type = chunk_dict['type']
                        n = len(chunk)

                        slide = prs.slides.add_slide(slide_layout)
                        move_slide(prs, len(prs.slides) - 1, vi_tri_hien_tai)

                        layout_rows = []
                        if c_type == 'portrait':
                            layout_rows = [chunk] 
                        else:
                            if n == 6:   layout_rows = [chunk[0:3], chunk[3:6]]
                            elif n == 5: layout_rows = [chunk[0:3], chunk[3:5]]
                            elif n == 4: layout_rows = [chunk[0:2], chunk[2:4]]
                            elif n == 3: layout_rows = [chunk] 
                            elif n == 2: layout_rows = [chunk]
                            elif n == 1: layout_rows = [chunk]

                        draw_adaptive_grid(slide, layout_rows, start_x_base, start_y_base, usable_w, usable_h, GAP)
                        vi_tri_hien_tai += 1

                output_stream = io.BytesIO()
                prs.save(output_stream)
                output_stream.seek(0)
                
                st.success("✅ Thành công! Hãy bấm nút bên dưới để tải File về máy.")
                st.download_button(
                    label="📥 BẤM VÀO ĐÂY ĐỂ TẢI FILE POWERPOINT XUỐNG",
                    data=output_stream,
                    file_name="Report_Auto_Exported.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra: {e}")
