import io
import re
import math
from datetime import datetime
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image, ImageOps

# --- SỬA LỖI ĐỌC ẢNH IPHONE (HEIC) ---
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except ImportError:
    pass

st.set_page_config(page_title="Công Cụ Chèn Ảnh PowerPoint", page_icon="⚡", layout="centered")

# ==========================================
# KHỞI TẠO BỘ NHỚ TẠM (GIỎ HÀNG LÁCH LUẬT iOS)
# ==========================================
if "photo_cart" not in st.session_state:
    st.session_state.photo_cart = {} # Chứa ảnh cộng dồn
if "template_bytes" not in st.session_state:
    st.session_state.template_bytes = None # Nhớ file mẫu

# ==========================================
# HÀM LÕI XỬ LÝ PPTX
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
st.error("🍎 **CÔNG NGHỆ CHỐNG VĂNG ẢNH TRÊN IPHONE:** Tải lên từng đợt nhỏ (2-5 ảnh), hệ thống tự động GOM VÀO GIỎ HÀNG!")

# --- BƯỚC 1 ---
st.header("Bước 1: Tải lên File PowerPoint Mẫu (Không bắt buộc)")
template_file = st.file_uploader("Chọn file .pptx (Chỉ cần up 1 lần)", type=["pptx"])
if template_file:
    st.session_state.template_bytes = template_file.getvalue()

if st.session_state.template_bytes:
    st.success("✅ Đã lưu file PPTX mẫu vào bộ nhớ!")
    if st.button("🗑️ Xóa file mẫu (Để tạo file trắng mới)"):
        st.session_state.template_bytes = None
        st.rerun()

# --- BƯỚC 2 ---
st.header("Bước 2: Ném ảnh vào Giỏ")

uploaded_images = st.file_uploader("📂 Nhấn để chọn ảnh từ máy", accept_multiple_files=True)
if uploaded_images:
    count = 0
    for f in uploaded_images:
        if f.name not in st.session_state.photo_cart:
            st.session_state.photo_cart[f.name] = {"bytes": f.getvalue(), "name": f.name}
            count += 1
    if count > 0:
        st.success(f"🎉 Vừa nhặt thêm {count} ảnh vào giỏ!")

camera_photo = st.camera_input("📷 Hoặc chụp ảnh thực tế tại hiện trường")
if camera_photo:
    cam_bytes = camera_photo.getvalue()
    
    # BỘ LỌC CHỐNG NHÂN BẢN ẢNH KHI TRANG WEB RELOAD
    is_duplicate = False
    for item in st.session_state.photo_cart.values():
        if item["bytes"] == cam_bytes:
            is_duplicate = True
            break
            
    if not is_duplicate:
        cam_name = f"cam_{datetime.now().strftime('%H%M%S')}.jpg"
        st.session_state.photo_cart[cam_name] = {"bytes": cam_bytes, "name": cam_name}
        st.success("📸 Đã ném ảnh vừa chụp vào giỏ!")

if st.session_state.photo_cart:
    st.info(f"🛒 **TRONG GIỎ ĐANG CÓ: {len(st.session_state.photo_cart)} ẢNH** ĐÃ SẴN SÀNG.")
    if st.button("🗑️ Làm trống Giỏ hàng để chọn lại từ đầu"):
        st.session_state.photo_cart = {}
        st.rerun()

# --- BƯỚC 3 ---
st.header("Bước 3: Tùy Chỉnh Layout")
mode = st.radio("Chọn chế độ:", 
                ("Layout 1: Cơ bản (1-2 ảnh/trang)", 
                 "Layout 2: Tràn viền (Lưới tự động chia đều, nhiều ảnh/trang)"))

align_mode = "2"
if "Layout 1" in mode:
    align_mode = st.radio("Căn lề cho ảnh (Layout 1):", ("1 - Trái", "2 - Giữa", "3 - Phải"))
    align_mode = align_mode[0]

# --- BƯỚC 4 ---
st.header("Bước 4: Vị trí chèn & Thiết kế")
vitri_input = st.text_input("Chèn vào sau Slide số mấy? (Gõ 0 để chèn cuối cùng):", "0")

st.markdown("---")
# TÍNH NĂNG MỚI: TRANG BÌA & ẢNH NỀN
use_blank = False
main_title = ""
sub_title = ""
bg_image_file = None

if not st.session_state.template_bytes:
    st.warning("⚠️ Bạn chưa tải file PowerPoint mẫu (ở Bước 1).")
    use_blank = st.checkbox("👉 Tích vào đây để BỎ QUA và TẠO FILE MỚI TINH")
    
    if use_blank:
        st.info("🎨 THIẾT KẾ TRANG BÌA & ẢNH NỀN (Tùy chọn):")
        main_title = st.text_input("Tiêu đề chính (Sẽ tự động tạo Slide Bìa):", placeholder="VD: BÁO CÁO NGHIỆM THU SỰ KIỆN")
        sub_title = st.text_input("Tiêu đề phụ:", placeholder="VD: Ngày 16/08/2026 - Địa điểm: SECC")
        bg_image_file = st.file_uploader("🖼️ Tải 1 ảnh làm HÌNH NỀN (Chỉ áp dụng cho Bìa đầu và Bìa cuối):", type=['jpg', 'jpeg', 'png', 'webp', 'heic'])

# --- XỬ LÝ XUẤT FILE ---
if st.button("🚀 XUẤT FILE POWERPOINT", use_container_width=True, type="primary"):
    if not st.session_state.template_bytes and not use_blank:
        st.error("⚠️ Vui lòng tải file PowerPoint mẫu ở Bước 1, HOẶC tích vào ô '👉 Tích vào đây để tạo file mới' nhé!")
    elif not st.session_state.photo_cart:
        st.error("⚠️ Giỏ ảnh đang trống trơn! Bác chọn thêm ảnh ở Bước 2 nhé!")
    else:
        with st.spinner("Đang tự động dàn trang... Vui lòng đợi nhé..."):
            try:
                # 1. Chuẩn bị Ảnh nền (nếu có)
                bg_stream = None
                if use_blank and bg_image_file:
                    try:
                        bg_bytes = bg_image_file.getvalue()
                        with Image.open(io.BytesIO(bg_bytes)) as img:
                            img = ImageOps.exif_transpose(img)
                            if img.mode != 'RGB':
                                img = img.convert('RGB')
                            bg_stream = io.BytesIO()
                            img.save(bg_stream, format='JPEG', quality=90)
                    except Exception:
                        pass

                # 2. Mở file PPTX
                if st.session_state.template_bytes:
                    prs = Presentation(io.BytesIO(st.session_state.template_bytes))
                else:
                    prs = Presentation()
                    # ÉP FILE TRẮNG VỀ CHUẨN WIDESCREEN 16:9
                    prs.slide_width = Inches(13.3333333)
                    prs.slide_height = Inches(7.5)

                tong_slide = len(prs.slides)
                vi_tri_hien_tai = tong_slide 
                
                # Nếu ghép vào file có sẵn, xác định vị trí chèn
                if tong_slide > 0:
                    match = re.search(r'\d+', vitri_input)
                    if match:
                        trang_chon = int(match.group())
                        if 0 < trang_chon <= tong_slide:
                            vi_tri_hien_tai = trang_chon
                else:
                    vi_tri_hien_tai = 0
                
                slide_w = prs.slide_width
                slide_h = prs.slide_height

                try:
                    slide_layout = prs.slide_layouts[6] # Blank
                except:
                    slide_layout = prs.slide_layouts[0] 

                # 3. TẠO TRANG BÌA (NẾU CÓ TIÊU ĐỀ)
                if use_blank and (main_title or sub_title):
                    title_slide = prs.slides.add_slide(slide_layout)
                    
                    # Áp ảnh nền trang bìa
                    if bg_stream:
                        bg_stream.seek(0)
                        title_slide.shapes.add_picture(bg_stream, 0, 0, width=slide_w, height=slide_h)
                    
                    # Chèn Tiêu đề chính (Dời xuống dưới)
                    if main_title:
                        txBox = title_slide.shapes.add_textbox(Inches(0.5), slide_h - Inches(2.2), slide_w - Inches(1), Inches(1))
                        tf = txBox.text_frame
                        tf.word_wrap = True
                        p = tf.paragraphs[0]
                        p.text = main_title.upper() # In hoa
                        p.alignment = PP_ALIGN.CENTER
                        p.font.size = Pt(44)
                        p.font.bold = True
                        p.font.color.rgb = RGBColor(255, 255, 255) if bg_stream else RGBColor(0, 51, 102)

                    # Chèn Tiêu đề phụ (Dời xuống dưới)
                    if sub_title:
                        txBox2 = title_slide.shapes.add_textbox(Inches(0.5), slide_h - Inches(1.2), slide_w - Inches(1), Inches(0.8))
                        tf2 = txBox2.text_frame
                        tf2.word_wrap = True
                        p2 = tf2.paragraphs[0]
                        p2.text = sub_title
                        p2.alignment = PP_ALIGN.CENTER
                        p2.font.size = Pt(24)
                        p2.font.color.rgb = RGBColor(240, 240, 240) if bg_stream else RGBColor(80, 80, 80)
                    
                    # Cập nhật lại vị trí chèn ảnh ngay sau trang bìa
                    vi_tri_hien_tai = len(prs.slides)

                CACH_LE_TREN = Inches(0.9)  
                CACH_LE_DUOI = Inches(0.8)  
                CACH_LE_TRAI = Inches(0.2) 
                CACH_LE_PHAI = Inches(0.2) 
                GAP = Inches(0.12)          
                
                usable_w = slide_w - CACH_LE_TRAI - CACH_LE_PHAI
                usable_h = slide_h - CACH_LE_TREN - CACH_LE_DUOI

                # 4. Xử lý ảnh trong giỏ
                image_data = []
                for key, img_info in st.session_state.photo_cart.items():
                    img_stream = io.BytesIO(img_info["bytes"])
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
                            'stream': img_stream, 'is_portrait': is_portrait, 
                            'w': width, 'h': height, 'name': img_info["name"], 'timestamp': str(dt_str)
                        })
                    except Exception:
                        pass

                image_data.sort(key=lambda x: (x['timestamp'], x['name']))

                # 5. DÀN TRANG NỘI DUNG CHÍNH (Bỏ Background nền, để trắng)
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
                            
                            if align_mode == '1': start_x = CACH_LE_TRAI
                            elif align_mode == '3': start_x = slide_w - CACH_LE_PHAI - block_w
                            else: start_x = CACH_LE_TRAI + (usable_w - block_w) / 2
                                
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
                                
                            if align_mode == '1': s_x = CACH_LE_TRAI
                            elif align_mode == '3': s_x = slide_w - CACH_LE_PHAI - f_w
                            else: s_x = CACH_LE_TRAI + (usable_w - f_w) / 2
                                
                            s_y = CACH_LE_TREN + (usable_h - f_h) / 2
                            add_image_exact(slide, current_img['stream'], s_x, s_y, f_w, f_h)
                            i += 1
                        vi_tri_hien_tai += 1

                elif "Layout 2" in mode:
                    landscapes = [img for img in image_data if not img['is_portrait']]
                    portraits = [img for img in image_data if img['is_portrait']]

                    smart_chunks = []
                    for c in partition_images(landscapes, 6): smart_chunks.append({'type': 'landscape', 'images': c})
                    for c in partition_images(portraits, 4): smart_chunks.append({'type': 'portrait', 'images': c})

                    for chunk_dict in smart_chunks:
                        chunk = chunk_dict['images']
                        n = len(chunk)

                        slide = prs.slides.add_slide(slide_layout)

                        move_slide(prs, len(prs.slides) - 1, vi_tri_hien_tai)

                        layout_rows = []
                        if chunk_dict['type'] == 'portrait':
                            layout_rows = [chunk] 
                        else:
                            if n == 6:   layout_rows = [chunk[0:3], chunk[3:6]]
                            elif n == 5: layout_rows = [chunk[0:3], chunk[3:5]]
                            elif n == 4: layout_rows = [chunk[0:2], chunk[2:4]]
                            else: layout_rows = [chunk]

                        draw_adaptive_grid(slide, layout_rows, CACH_LE_TRAI, CACH_LE_TREN, usable_w, usable_h, GAP)
                        vi_tri_hien_tai += 1

                # 6. TẠO TRANG KẾT THÚC (THANK YOU) NẾU LÀ TẠO FILE MỚI
                if use_blank:
                    ty_slide = prs.slides.add_slide(slide_layout)
                    
                    # Áp ảnh nền
                    if bg_stream:
                        bg_stream.seek(0)
                        ty_slide.shapes.add_picture(bg_stream, 0, 0, width=slide_w, height=slide_h)
                    
                    move_slide(prs, len(prs.slides) - 1, vi_tri_hien_tai)

                    # Chữ Thank You (Đã thu nhỏ & dời xuống dưới)
                    txBox_ty = ty_slide.shapes.add_textbox(Inches(0.5), slide_h - Inches(2.2), slide_w - Inches(1), Inches(1))
                    tf_ty = txBox_ty.text_frame
                    tf_ty.word_wrap = True
                    p_ty = tf_ty.paragraphs[0]
                    p_ty.text = "THANK YOU!"
                    p_ty.alignment = PP_ALIGN.CENTER
                    p_ty.font.size = Pt(44) # Thu bé lại
                    p_ty.font.bold = True
                    p_ty.font.color.rgb = RGBColor(255, 255, 255) if bg_stream else RGBColor(0, 51, 102)
                    
                    vi_tri_hien_tai += 1

                output_stream = io.BytesIO()
                prs.save(output_stream)
                output_stream.seek(0)
                
                st.success("✅ Thành công mỹ mãn! File đã sẵn sàng.")
                st.download_button(
                    label="📥 TẢI FILE POWERPOINT XUỐNG",
                    data=output_stream,
                    file_name="Report_Kiem_Tra_Chuan.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra: {e}")
