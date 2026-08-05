import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from io import BytesIO
import zipfile
import re
import os

# --- นำเข้าไลบรารีสำหรับสร้าง PowerPoint ---
try:
    from pptx import Presentation
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
except ImportError:
    st.error("⚠️ ไม่พบไลบรารี python-pptx")
    st.stop()

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================

def get_font(font_data, size):
    """โหลดฟอนต์จากข้อมูลที่อัปโหลด"""
    try:
        return ImageFont.truetype(BytesIO(font_data), size)
    except:
        return ImageFont.load_default()

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "certificate"

def render_certificate(template_img, texts, row_data=None, fonts_dict=None):
    """สร้างเกียรติบัตรแต่ละใบ"""
    img = template_img.copy()
    if img.mode != 'RGB':
        img = img.convert('RGB')
    
    draw = ImageDraw.Draw(img)
    
    for txt in texts:
        # ดึงข้อความ
        if txt['type'] == 'static':
            content = txt['text']
        else:
            if row_data and txt['column'] in row_data:
                val = row_data[txt['column']]
                content = str(val) if pd.notna(val) else ""
            else:
                content = "ตัวอย่าง"
        
        if not content:
            continue
        
        # โหลดฟอนต์
        font_data = None
        if fonts_dict and txt.get('font_name') in fonts_dict:
            font_data = fonts_dict[txt['font_name']]
        
        if font_data:
            font = get_font(font_data, txt['size'])
        else:
            font = ImageFont.load_default()
        
        # วัดขนาดข้อความ
        try:
            bbox = font.getbbox(content)
            text_width = bbox[2] - bbox[0]
        except:
            text_width = draw.textlength(content, font=font)
        
        # จัดกึ่งกลาง
        start_x = txt['x'] - (text_width / 2)
        
        # วาดข้อความ
        draw.text((start_x, txt['y']), content, fill=txt['color'], font=font, anchor="ls")
    
    return img

def create_pptx(template_img, texts, data_df, filename_col, fonts_dict=None):
    """สร้าง PowerPoint"""
    prs = Presentation()
    
    # กำหนดขนาดสไลด์
    img_width, img_height = template_img.size
    prs.slide_width = img_width * 9525
    prs.slide_height = img_height * 9525
    
    blank_slide_layout = prs.slide_layouts[6]
    
    for _, row in data_df.iterrows():
        # สร้างภาพเกียรติบัตร
        final_img = render_certificate(template_img, texts, row.to_dict(), fonts_dict)
        
        # บันทึกภาพ
        img_io = BytesIO()
        final_img.save(img_io, format="PNG")
        img_io.seek(0)
        
        # เพิ่มสไลด์
        slide = prs.slides.add_slide(blank_slide_layout)
        slide.shapes.add_picture(img_io, 0, 0, width=prs.slide_width, height=prs.slide_height)
    
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# ==========================================
# 🎨 UI - STREAMLIT APP
# ==========================================
st.set_page_config(page_title="Certificate Generator", layout="wide")

# Session State
if "click_x" not in st.session_state:
    st.session_state.click_x = 0
if "click_y" not in st.session_state:
    st.session_state.click_y = 0
if 'texts' not in st.session_state:
    st.session_state.texts = []
if 'template' not in st.session_state:
    st.session_state.template = None
if 'data' not in st.session_state:
    st.session_state.data = None
if 'fonts_dict' not in st.session_state:
    st.session_state.fonts_dict = {}
if 'font_names' not in st.session_state:
    st.session_state.font_names = []

st.title("📜 Certificate Generator")

# --- SIDEBAR ---
with st.sidebar:
    st.header("📁 อัปโหลดไฟล์")
    
    # 1. Template
    template_file = st.file_uploader("🖼️ พื้นหลังเกียรติบัตร", type=['jpg', 'jpeg', 'png'])
    if template_file:
        st.session_state.template = Image.open(template_file)
        st.success("✅ โหลดพื้นหลังสำเร็จ")
    
    # 2. Font Management
    st.markdown("---")
    st.header("🔤 จัดการฟอนต์")
    
    uploaded_font = st.file_uploader("อัปโหลดฟอนต์ (.ttf)", type=['ttf'])
    if uploaded_font:
        font_name = uploaded_font.name.split('.')[0]
        if font_name not in st.session_state.fonts_dict:
            st.session_state.fonts_dict[font_name] = uploaded_font.getvalue()
            st.session_state.font_names.append(font_name)
            st.success(f"✅ เพิ่มฟอนต์ '{font_name}' แล้ว")
    
    # แสดงฟอนต์ที่มี
    if st.session_state.font_names:
        st.write("**ฟอนต์ที่มี:**")
        for f in st.session_state.font_names:
            st.write(f"- {f}")
    
    # 3. Data
    st.markdown("---")
    st.header("📊 รายชื่อข้อมูล")
    data_file = st.file_uploader("ไฟล์ Excel/CSV", type=['xlsx', 'xls', 'csv'])
    if data_file:
        try:
            if data_file.name.endswith('.csv'):
                st.session_state.data = pd.read_csv(data_file)
            else:
                st.session_state.data = pd.read_excel(data_file)
            st.success(f"✅ โหลดข้อมูล {len(st.session_state.data)} รายการ")
        except Exception as e:
            st.error(f"❌ โหลดข้อมูลไม่สำเร็จ: {e}")

# --- MAIN ---
if st.session_state.template is None:
    st.info("👈 กรุณาอัปโหลดพื้นหลังเกียรติบัตรทางด้านซ้าย")
    st.stop()

st.header("📍 กำหนดข้อความ")

col1, col2 = st.columns([1.5, 1])

with col1:
    st.markdown("**🖱️ คลิกที่รูปเพื่อกำหนดตำแหน่ง (พิกัดจะอัปเดตอัตโนมัติ)**")
    
    # แสดงรูปและรับตำแหน่งคลิกแบบง่าย
    original_w, original_h = st.session_state.template.size
    display_w = 650
    ratio = original_w / display_w if original_w > display_w else 1.0
    display_h = int(original_h / ratio) if original_w > display_w else original_h
    display_img = st.session_state.template.resize((display_w, display_h))
    
    # ใช้ image + st.button แทนการคลิก
    col_img, col_coords = st.columns([3, 1])
    
    with col_img:
        st.image(display_img, use_column_width=True)
    
    with col_coords:
        st.markdown("### 📍 พิกัด")
        
        # ใช้滑块แทนการคลิก (ง่ายกว่าและใช้ได้ทุกที่)
        st.write("**ปรับตำแหน่ง X**")
        new_x = st.slider("", 0, original_w, st.session_state.click_x, key="slider_x")
        
        st.write("**ปรับตำแหน่ง Y**")
        new_y = st.slider("", 0, original_h, st.session_state.click_y, key="slider_y")
        
        # อัปเดตพิกัด
        st.session_state.click_x = new_x
        st.session_state.click_y = new_y
        
        st.metric("X", st.session_state.click_x)
        st.metric("Y", st.session_state.click_y)
        
        # ปุ่มรีเซ็ต
        if st.button("🔄 รีเซ็ตพิกัด (0,0)"):
            st.session_state.click_x = 0
            st.session_state.click_y = 0
            st.rerun()

with col2:
    st.markdown("**✏️ เพิ่มข้อความ**")
    
    with st.form("add_text_form", clear_on_submit=False):
        # ประเภทข้อความ
        t_type = st.radio("ชนิด", ["พิมพ์เอง", "ดึงจาก Excel"], horizontal=True)
        
        # ข้อความ
        if t_type == "พิมพ์เอง":
            t_text = st.text_input("พิมพ์ข้อความ", placeholder="เช่น ชื่อผู้รับเกียรติบัตร")
            t_col = ""
        else:
            if st.session_state.data is not None:
                t_col = st.selectbox("เลือกคอลัมน์", st.session_state.data.columns)
                t_text = ""
                # แสดงตัวอย่าง
                if st.session_state.data is not None and len(st.session_state.data) > 0:
                    st.caption(f"📝 ตัวอย่าง: {st.session_state.data[t_col].iloc[0]}")
            else:
                st.warning("⚠️ กรุณาอัปโหลดไฟล์ Excel ก่อน")
                t_col = ""
        
        # พิกัด (ใช้ค่าจาก slider)
        c1, c2 = st.columns(2)
        with c1:
            x_pos = st.number_input("X", value=st.session_state.click_x, step=1)
        with c2:
            y_pos = st.number_input("Y", value=st.session_state.click_y, step=1)
        
        # ขนาดและสี
        size = st.slider("ขนาดฟอนต์", 10, 500, 60)
        color = st.color_picker("สีข้อความ", "#000000")
        
        # เลือกฟอนต์
        if st.session_state.font_names:
            selected_font = st.selectbox("เลือกฟอนต์", st.session_state.font_names)
        else:
            st.warning("⚠️ กรุณาอัปโหลดฟอนต์ก่อน")
            selected_font = None
        
        if st.form_submit_button("➕ เพิ่มข้อความ"):
            if t_type == "พิมพ์เอง" and t_text and selected_font:
                st.session_state.texts.append({
                    'type': 'static',
                    'text': t_text,
                    'column': '',
                    'x': x_pos,
                    'y': y_pos,
                    'size': size,
                    'color': color,
                    'font_name': selected_font
                })
                st.success("✅ เพิ่มข้อความแล้ว!")
                st.rerun()
            elif t_type == "ดึงจาก Excel" and t_col and selected_font:
                st.session_state.texts.append({
                    'type': 'excel',
                    'text': '',
                    'column': t_col,
                    'x': x_pos,
                    'y': y_pos,
                    'size': size,
                    'color': color,
                    'font_name': selected_font
                })
                st.success("✅ เพิ่มข้อความแล้ว!")
                st.rerun()
            else:
                if not selected_font:
                    st.error("❌ กรุณาอัปโหลดและเลือกฟอนต์ก่อน")
                else:
                    st.error("❌ กรุณากรอกข้อมูลให้ครบ")

# --- แสดงข้อความที่เพิ่ม ---
st.markdown("---")
st.header("📋 รายการข้อความ")

if st.session_state.texts:
    for i, txt in enumerate(st.session_state.texts):
        cols = st.columns([4, 1])
        label = txt['text'] if txt['type'] == 'static' else f"📊 {txt['column']}"
        font_label = f"🔤 {txt.get('font_name', 'Default')}"
        cols[0].write(f"{i+1}. {label} | {font_label} | ขนาด: {txt['size']} | พิกัด: ({txt['x']}, {txt['y']})")
        if cols[1].button("🗑️", key=f"del_{i}"):
            st.session_state.texts.pop(i)
            st.rerun()
    
    # Preview
    st.markdown("---")
    st.header("👁️ ตัวอย่าง")
    
    preview_row = None
    if st.session_state.data is not None and len(st.session_state.data) > 0:
        row_idx = st.number_input("แถวที่", 0, len(st.session_state.data)-1, 0)
        preview_row = st.session_state.data.iloc[row_idx].to_dict()
    
    preview_img = render_certificate(
        st.session_state.template, 
        st.session_state.texts, 
        preview_row,
        st.session_state.fonts_dict
    )
    st.image(preview_img, width=650)

else:
    st.info("💡 ยังไม่มีข้อความ เพิ่มข้อความด้านบน")

# --- Export ---
if st.session_state.data is not None and st.session_state.texts:
    st.markdown("---")
    st.header("📦 สร้างและดาวน์โหลด")
    
    col_export1, col_export2 = st.columns(2)
    
    with col_export1:
        filename_col = st.selectbox("เลือกคอลัมน์ชื่อไฟล์", st.session_state.data.columns)
    
    with col_export2:
        export_format = st.radio("รูปแบบ", ["PNG", "PDF", "PowerPoint"], horizontal=True)
    
    if st.button("🚀 สร้างไฟล์ทั้งหมด", type="primary"):
        with st.spinner("กำลังสร้าง..."):
            if export_format == "PowerPoint":
                # สร้าง PowerPoint
                pptx_io = create_pptx(
                    st.session_state.template, 
                    st.session_state.texts, 
                    st.session_state.data, 
                    filename_col,
                    st.session_state.fonts_dict
                )
                st.success("✅ สร้าง PowerPoint สำเร็จ!")
                st.download_button(
                    "📥 ดาวน์โหลด PowerPoint",
                    pptx_io.getvalue(),
                    "certificates.pptx",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                # สร้าง ZIP
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for _, row in st.session_state.data.iterrows():
                        img = render_certificate(
                            st.session_state.template, 
                            st.session_state.texts, 
                            row.to_dict(),
                            st.session_state.fonts_dict
                        )
                        img_io = BytesIO()
                        
                        ext = "png" if export_format == "PNG" else "pdf"
                        img.save(img_io, format=export_format)
                        
                        filename = sanitize_filename(row[filename_col])
                        zf.writestr(f"{filename}.{ext}", img_io.getvalue())
                
                st.success("✅ สร้างไฟล์ทั้งหมดสำเร็จ!")
                st.download_button(
                    f"📥 ดาวน์โหลด ZIP ({export_format})",
                    zip_buffer.getvalue(),
                    "certificates.zip",
                    "application/zip"
                )

# --- คำแนะนำการใช้งาน ---
with st.sidebar:
    st.markdown("---")
    st.markdown("""
    ### 📖 วิธีใช้
    1. อัปโหลดพื้นหลังเกียรติบัตร
    2. อัปโหลดฟอนต์ที่ต้องการ
    3. อัปโหลดไฟล์ Excel (ถ้ามี)
    4. ปรับ slider เพื่อกำหนดตำแหน่ง
    5. เพิ่มข้อความ
    6. สร้างไฟล์และดาวน์โหลด
    """)
