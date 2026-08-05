import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from io import BytesIO
import zipfile
import re
import os
import json
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

# ==========================================
# 🛠️ CONSTANTS
# ==========================================
EMU_PER_PIXEL = 9525
FONT_CACHE = {}

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================

def get_system_font_path():
    """ค้นหาฟอนต์ที่ปรับขนาดได้ในระบบ"""
    paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\tahoma.ttf",
        "/System/Library/Fonts/Helvetica.ttf"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

def fix_thai_text_old(text):
    """จัดตำแหน่งสระและวรรณยุกต์สำหรับฟอนต์รุ่นเก่า PUA"""
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    
    tone_marks = ['\u0e48', '\u0e49', '\u0e4a', '\u0e4b', '\u0e4c']
    high_tone_marks = ['\uf713', '\uf714', '\uf715', '\uf716', '\uf717']
    
    upper_vowels = ['\u0e31', '\u0e34', '\u0e35', '\u0e36', '\u0e37', '\u0e4d']
    for i, tone in enumerate(tone_marks):
        for vowel in upper_vowels:
            text = text.replace(vowel + tone, vowel + high_tone_marks[i])
    
    tall_consonants = ['ป', 'ฝ', 'ฟ']
    left_tone_marks = ['\uf70a', '\uf70b', '\uf70c', '\uf70d', '\uf70e']
    for i, tone in enumerate(tone_marks):
        for cons in tall_consonants:
            text = text.replace(cons + tone, cons + left_tone_marks[i])
    
    text = text.replace('\u0e4d\u0e32', '\u0e33')
    
    YO_YING_NO_BASE = '\uF70F'
    THO_THAN_NO_BASE = '\uF700'
    text = text.replace('ญุ', YO_YING_NO_BASE + 'ุ')
    text = text.replace('ญู', YO_YING_NO_BASE + 'ู')
    text = text.replace('ญฺ', YO_YING_NO_BASE + 'ฺ')
    text = text.replace('ฐุ', THO_THAN_NO_BASE + 'ุ')
    text = text.replace('ฐู', THO_THAN_NO_BASE + 'ู')
    text = text.replace('ฐฺ', THO_THAN_NO_BASE + 'ฺ')
    
    for i, tone in enumerate(tone_marks):
        text = text.replace('ญ' + tone, YO_YING_NO_BASE + high_tone_marks[i])
    
    text = text.replace('เกี๊ย', 'เกี\uf714ย')
    text = text.replace('เกี๋ย', 'เกี\uf712ย')
    text = text.replace('เกื๊อ', 'เกื\uf714อ')
    text = text.replace('เกื๋อ', 'เกื\uf712อ')
    text = text.replace('ำ้', '\u0e33\u0e49')
    text = text.replace('ณ์', '\u0e13\u0e4c')
    text = text.replace('ร์', '\u0e23\u0e4c')
    
    return text

def fix_thai_text_new(text):
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    return text

def fix_thai_text(text, font_version="ใหม่"):
    if font_version == "เก่า":
        return fix_thai_text_old(text)
    return fix_thai_text_new(text)

def get_font(font_name, size):
    """โหลดฟอนต์พร้อม Cache"""
    cache_key = f"{font_name}_{size}"
    if cache_key in FONT_CACHE:
        return FONT_CACHE[cache_key]
    
    try:
        if font_name and font_name in st.session_state.fonts_dict:
            font_data = st.session_state.fonts_dict[font_name]['data']
            font = ImageFont.truetype(BytesIO(font_data), size)
            FONT_CACHE[cache_key] = font
            return font
    except Exception:
        pass
    
    sys_path = get_system_font_path()
    if sys_path:
        try:
            font = ImageFont.truetype(sys_path, size)
            FONT_CACHE[cache_key] = font
            return font
        except:
            pass
    
    font = ImageFont.load_default()
    FONT_CACHE[cache_key] = font
    return font

def get_text_bbox(text, font):
    """คำนวณ bbox ของข้อความ"""
    try:
        bbox = font.getbbox(text)
        return bbox
    except AttributeError:
        try:
            width, height = font.getsize(text)
            return (0, 0, width, height)
        except:
            draw = ImageDraw.Draw(Image.new('RGB', (1, 1)))
            width = draw.textlength(text, font=font)
            return (0, 0, int(width), int(width * 0.6))

def render_certificate(template_img, texts, row_data=None):
    """สร้างภาพเกียรติบัตรพร้อมข้อความ"""
    img = template_img.copy()
    if img.mode != 'RGB':
        img = img.convert('RGB')
    draw = ImageDraw.Draw(img)
    
    for txt in texts:
        if txt['type'] == 'static':
            content = txt['text']
        else:
            if row_data and txt['column'] in row_data:
                val = row_data[txt['column']]
                content = str(val) if pd.notna(val) else ""
            else:
                content = "ตัวอย่างข้อมูล"
        
        if not content: 
            continue
        
        font_version = txt.get('font_version', 'ใหม่')
        content = fix_thai_text(content, font_version)
        font = get_font(txt.get('font_name'), txt['size'])
        
        bbox = get_text_bbox(content, font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]
        
        start_x = txt['x'] - (text_width / 2)
        start_y = txt['y'] - (text_height / 2)
        
        draw.text((start_x, start_y), content, fill=txt['color'], font=font)
    
    return img

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return (0, 0, 0)

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "certificate"

def create_pptx_with_editable_text(template_img, texts, data_df):
    """สร้าง PowerPoint ที่ข้อความสามารถแก้ไขได้"""
    prs = Presentation()
    
    img_width, img_height = template_img.size
    prs.slide_width = img_width * EMU_PER_PIXEL
    prs.slide_height = img_height * EMU_PER_PIXEL
    
    blank_slide_layout = prs.slide_layouts[6]
    
    bg_io = BytesIO()
    template_img.save(bg_io, format="PNG")
    bg_bytes = bg_io.getvalue()
    
    for idx, row in data_df.iterrows():
        slide = prs.slides.add_slide(blank_slide_layout)
        
        bg_io = BytesIO(bg_bytes)
        slide.shapes.add_picture(bg_io, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        for txt in texts:
            if txt['type'] == 'static':
                content = txt['text']
            else:
                if txt['column'] in row:
                    val = row[txt['column']]
                    content = str(val) if pd.notna(val) else ""
                else:
                    continue
            
            if not content: 
                continue
            
            ppt_font_version = txt.get('ppt_font_version', 'ใหม่')
            ppt_content = fix_thai_text(content, ppt_font_version)
            
            font_size_pt = txt['size'] * 0.75
            
            font = get_font(txt.get('font_name'), txt['size'])
            bbox = get_text_bbox(content, font)
            text_width_px = bbox[2] - bbox[0]
            text_height_px = bbox[3] - bbox[1]
            
            margin_px = 20
            box_width_px = text_width_px + margin_px * 2
            box_height_px = text_height_px + margin_px * 2
            
            box_width_emu = int(box_width_px * EMU_PER_PIXEL)
            box_height_emu = int(box_height_px * EMU_PER_PIXEL)
            
            center_x_emu = txt['x'] * EMU_PER_PIXEL
            center_y_emu = txt['y'] * EMU_PER_PIXEL
            
            left_emu = center_x_emu - (box_width_emu / 2)
            top_emu = center_y_emu - (box_height_emu / 2)
            
            left_emu = max(0, left_emu)
            top_emu = max(0, top_emu)
            
            txBox = slide.shapes.add_textbox(
                int(left_emu), 
                int(top_emu), 
                int(box_width_emu), 
                int(box_height_emu)
            )
            
            tf = txBox.text_frame
            tf.text = ppt_content
            tf.word_wrap = False
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            
            run = p.runs[0]
            run.font.size = Pt(font_size_pt)
            if txt.get('font_name'):
                run.font.name = txt['font_name']
            
            rgb = hex_to_rgb(txt['color'])
            run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

# ==========================================
# 🎨 UI - STREAMLIT APP
# ==========================================
st.set_page_config(page_title="Auto Cert Pro", layout="wide")

# ตั้งค่า Session State
if "click_x" not in st.session_state: 
    st.session_state.click_x = 0
if "click_y" not in st.session_state: 
    st.session_state.click_y = 0
if 'texts' not in st.session_state: 
    st.session_state.texts = []
if 'fonts_dict' not in st.session_state: 
    st.session_state.fonts_dict = {}
if 'font_names' not in st.session_state: 
    st.session_state.font_names = []
if 'template' not in st.session_state:
    st.session_state.template = None
if 'data' not in st.session_state:
    st.session_state.data = None
if 'preview_row' not in st.session_state:
    st.session_state.preview_row = 0
if 'edit_index' not in st.session_state:
    st.session_state.edit_index = None
if 'zoom_level' not in st.session_state:
    st.session_state.zoom_level = 1.0

st.title("📜 Auto Certificate Generator")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1️⃣ อัปโหลดไฟล์")
    template_file = st.file_uploader("🖼️ พื้นหลังเกียรติบัตร (JPG/PNG)", type=['jpg', 'jpeg', 'png'])
    if template_file:
        st.session_state.template = Image.open(template_file)
        st.success("✅ โหลดพื้นหลังสำเร็จ")

    st.markdown("---")
    st.header("2️⃣ จัดการฟอนต์")
    
    st.subheader("ฟอนต์สำหรับ PNG/PDF")
    font_type_png = st.radio(
        "ประเภทฟอนต์ PNG/PDF",
        ["ฟอนต์รุ่นใหม่ (OpenType)", "ฟอนต์รุ่นเก่า (PUA)"],
        key="font_type_png"
    )
    
    st.subheader("ฟอนต์สำหรับ PowerPoint")
    font_type_ppt = st.radio(
        "ประเภทฟอนต์ PowerPoint",
        ["ฟอนต์รุ่นใหม่ (OpenType)", "ฟอนต์รุ่นเก่า (PUA)"],
        key="font_type_ppt"
    )
    
    uploaded_font = st.file_uploader("🔤 อัปโหลดฟอนต์ .ttf", type=['ttf'])
    
    if uploaded_font:
        f_name = uploaded_font.name.split('.')[0]
        if f_name not in st.session_state.fonts_dict:
            st.session_state.fonts_dict[f_name] = {
                'data': uploaded_font.getvalue()
            }
            st.session_state.font_names.append(f_name)
            st.success(f"✅ เพิ่มฟอนต์ '{f_name}' แล้ว")
    
    if st.session_state.font_names:
        st.markdown("---")
        st.write("**📋 ฟอนต์ที่มี:**")
        for f in st.session_state.font_names:
            st.write(f"- {f}")

    st.markdown("---")
    st.header("3️⃣ รายชื่อข้อมูล")
    data_file = st.file_uploader("📊 ไฟล์ Excel/CSV", type=['xlsx', 'xls', 'csv'])
    if data_file:
        try:
            if data_file.name.endswith('.csv'):
                st.session_state.data = pd.read_csv(data_file)
            else:
                st.session_state.data = pd.read_excel(data_file)
            st.success(f"✅ โหลดข้อมูล {len(st.session_state.data)} รายการ")
        except Exception as e:
            st.error(f"❌ โหลดข้อมูลไม่สำเร็จ: {e}")

# ตรวจสอบ template
if st.session_state.template is None:
    st.info("👈 กรุณาอัปโหลด 'พื้นหลังเกียรติบัตร' ทางด้านซ้ายเพื่อเริ่มต้น")
    st.stop()

# --- MAIN AREA ---
st.header("📍 กำหนดตำแหน่งและข้อความ")

col_img, col_form = st.columns([1.5, 1])

with col_img:
    # แสดงตัวเลือกแถวตัวอย่าง
    if st.session_state.data is not None and not st.session_state.data.empty:
        st.session_state.preview_row = st.number_input(
            "ดูตัวอย่างจากแถวที่:", 
            0, 
            len(st.session_state.data)-1, 
            st.session_state.preview_row,
            key="preview_row_input"
        )
        preview_row = st.session_state.data.iloc[st.session_state.preview_row].to_dict()
    else:
        preview_row = None
    
    # สร้างรูปพรีวิว
    current_preview = render_certificate(st.session_state.template, st.session_state.texts, preview_row)
    
    st.markdown("**🖱️ คลิกที่รูปเพื่อกำหนดตำแหน่ง (ข้อความจะอยู่กึ่งกลางจุดคลิก)**")
    
    # ✅ แทนที่ streamlit_image_coordinates ด้วยการคลิกบนรูป + JavaScript
    original_w, original_h = current_preview.size
    
    # ตั้งค่า Zoom
    zoom_levels = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0]
    st.session_state.zoom_level = st.select_slider(
        "🔍 ซูม",
        options=zoom_levels,
        value=st.session_state.zoom_level,
        format_func=lambda x: f"{int(x*100)}%"
    )
    
    display_w = int(700 * st.session_state.zoom_level)
    display_h = int(original_h * (display_w / original_w))
    
    # ใช้วิธีการคลิกแบบอื่น - ใช้ปุ่มเลื่อน
    st.markdown("**🎯 ปรับตำแหน่งด้วยปุ่มควบคุม (ละเอียด 10 พิกเซล)**")
    
    # สร้างปุ่มควบคุม 8 ทิศทาง
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    with col_btn1:
        if st.button("↖️", use_container_width=True):
            st.session_state.click_x = max(0, st.session_state.click_x - 10)
            st.session_state.click_y = max(0, st.session_state.click_y - 10)
            st.rerun()
    with col_btn2:
        if st.button("⬆️", use_container_width=True):
            st.session_state.click_y = max(0, st.session_state.click_y - 10)
            st.rerun()
    with col_btn3:
        if st.button("↗️", use_container_width=True):
            st.session_state.click_x = min(original_w, st.session_state.click_x + 10)
            st.session_state.click_y = max(0, st.session_state.click_y - 10)
            st.rerun()
    
    col_btn4, col_btn5, col_btn6 = st.columns(3)
    with col_btn4:
        if st.button("⬅️", use_container_width=True):
            st.session_state.click_x = max(0, st.session_state.click_x - 10)
            st.rerun()
    with col_btn5:
        if st.button("🎯 กลาง", use_container_width=True):
            st.session_state.click_x = original_w // 2
            st.session_state.click_y = original_h // 2
            st.rerun()
    with col_btn6:
        if st.button("➡️", use_container_width=True):
            st.session_state.click_x = min(original_w, st.session_state.click_x + 10)
            st.rerun()
    
    col_btn7, col_btn8, col_btn9 = st.columns(3)
    with col_btn7:
        if st.button("↙️", use_container_width=True):
            st.session_state.click_x = max(0, st.session_state.click_x - 10)
            st.session_state.click_y = min(original_h, st.session_state.click_y + 10)
            st.rerun()
    with col_btn8:
        if st.button("⬇️", use_container_width=True):
            st.session_state.click_y = min(original_h, st.session_state.click_y + 10)
            st.rerun()
    with col_btn9:
        if st.button("↘️", use_container_width=True):
            st.session_state.click_x = min(original_w, st.session_state.click_x + 10)
            st.session_state.click_y = min(original_h, st.session_state.click_y + 10)
            st.rerun()
    
    # แสดงตำแหน่งปัจจุบัน
    st.info(f"📍 พิกัดปัจจุบัน: X={st.session_state.click_x}, Y={st.session_state.click_y}")
    
    # แสดงรูปภาพพร้อมตำแหน่ง
    img_with_marker = current_preview.copy()
    draw = ImageDraw.Draw(img_with_marker)
    
    # วาดเครื่องหมายกากบาท
    marker_size = 15
    draw.line([(st.session_state.click_x - marker_size, st.session_state.click_y), 
               (st.session_state.click_x + marker_size, st.session_state.click_y)], 
              fill='red', width=3)
    draw.line([(st.session_state.click_x, st.session_state.click_y - marker_size), 
               (st.session_state.click_x, st.session_state.click_y + marker_size)], 
              fill='red', width=3)
    draw.ellipse([(st.session_state.click_x - 5, st.session_state.click_y - 5), 
                  (st.session_state.click_x + 5, st.session_state.click_y + 5)], 
                 fill='red')
    
    # แสดงรูป
    img_display = img_with_marker.resize((display_w, display_h))
    st.image(img_display, use_column_width=True)

with col_form:
    edit_mode = st.session_state.edit_index is not None
    
    if edit_mode:
        st.subheader("✏️ แก้ไขข้อความ")
        edit_txt = st.session_state.texts[st.session_state.edit_index]
        default_text = edit_txt.get('text', '')
        default_column = edit_txt.get('column', None)
        default_type = 'excel' if edit_txt['type'] == 'excel' else 'พิมพ์เอง'
    else:
        st.subheader("➕ เพิ่มข้อความ")
        default_text = ''
        default_column = None
        default_type = 'พิมพ์เอง'
    
    cols_options = st.session_state.data.columns.tolist() if st.session_state.data is not None else []
    
    t_type = st.radio(
        "ชนิดข้อความ",
        ["ดึงจากไฟล์รายชื่อ", "พิมพ์เอง"],
        horizontal=True,
        key="text_type",
        index=0 if default_type == 'excel' else 1
    )
    
    if t_type == "ดึงจากไฟล์รายชื่อ":
        if not cols_options:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์รายชื่อก่อน")
            selected_column = None
        else:
            default_index = 0
            if default_column and default_column in cols_options:
                default_index = cols_options.index(default_column)
            selected_column = st.selectbox(
                "เลือกหัวข้อ (คอลัมน์)",
                cols_options,
                index=default_index,
                key="excel_column_select"
            )
        t_val = ""
    else:
        if edit_mode and default_type == 'พิมพ์เอง':
            t_val = st.text_input("ข้อความที่ต้องการพิมพ์", value=default_text, key="text_input")
        else:
            t_val = st.text_input("ข้อความที่ต้องการพิมพ์", key="text_input")
        selected_column = None
    
    with st.form("add_text_form", clear_on_submit=not edit_mode):
        c1, c2 = st.columns(2)
        
        if edit_mode:
            x_pos = c1.number_input("ตำแหน่ง X", value=edit_txt['x'])
            y_pos = c2.number_input("ตำแหน่ง Y", value=edit_txt['y'])
            f_size = st.slider("ขนาดฟอนต์", 10, 500, value=edit_txt['size'])
            f_color = st.color_picker("เลือกสีข้อความ", value=edit_txt['color'])
            default_font_index = 0
            if edit_txt['font_name'] in st.session_state.font_names:
                default_font_index = st.session_state.font_names.index(edit_txt['font_name'])
            selected_font = st.selectbox("เลือกฟอนต์", st.session_state.font_names, index=default_font_index)
        else:
            x_pos = c1.number_input("ตำแหน่ง X", value=st.session_state.click_x)
            y_pos = c2.number_input("ตำแหน่ง Y", value=st.session_state.click_y)
            f_size = st.slider("ขนาดฟอนต์", 10, 500, value=60)
            f_color = st.color_picker("เลือกสีข้อความ", value="#000000")
            
            if not st.session_state.font_names:
                st.warning("⚠️ กรุณาอัปโหลดฟอนต์ก่อน")
                selected_font = None
            else:
                selected_font = st.selectbox("เลือกฟอนต์", st.session_state.font_names)
        
        if edit_mode:
            submit_label = "💾 อัปเดตข้อความ"
        else:
            submit_label = "➕ เพิ่มข้อความ"
        
        submit = st.form_submit_button(submit_label)
        if submit:
            if not selected_font:
                st.error("❌ กรุณาเลือกฟอนต์")
            elif t_type == "พิมพ์เอง" and not t_val and not edit_mode:
                st.error("❌ กรุณาพิมพ์ข้อความ")
            elif t_type == "ดึงจากไฟล์รายชื่อ" and not selected_column and not edit_mode:
                st.error("❌ กรุณาเลือกคอลัมน์")
            else:
                font_version_png = "เก่า" if "เก่า" in font_type_png else "ใหม่"
                font_version_ppt = "เก่า" if "เก่า" in font_type_ppt else "ใหม่"
                
                if edit_mode and st.session_state.edit_index is not None:
                    st.session_state.texts[st.session_state.edit_index] = {
                        'type': 'excel' if t_type == "ดึงจากไฟล์รายชื่อ" else 'static',
                        'text': t_val if t_type == "พิมพ์เอง" else edit_txt.get('text', ''),
                        'column': selected_column if t_type == "ดึงจากไฟล์รายชื่อ" else edit_txt.get('column', None),
                        'x': x_pos,
                        'y': y_pos,
                        'size': f_size,
                        'color': f_color,
                        'font_name': selected_font,
                        'font_version': font_version_png,
                        'ppt_font_version': font_version_ppt
                    }
                    st.session_state.edit_index = None
                    st.success("✅ อัปเดตข้อความสำเร็จ!")
                else:
                    st.session_state.texts.append({
                        'type': 'excel' if t_type == "ดึงจากไฟล์รายชื่อ" else 'static',
                        'text': t_val,
                        'column': selected_column,
                        'x': x_pos,
                        'y': y_pos,
                        'size': f_size,
                        'color': f_color,
                        'font_name': selected_font,
                        'font_version': font_version_png,
                        'ppt_font_version': font_version_ppt
                    })
                    st.success("✅ เพิ่มข้อความสำเร็จ!")
                st.rerun()

    if edit_mode:
        if st.button("❌ ยกเลิกการแก้ไข"):
            st.session_state.edit_index = None
            st.rerun()

    if st.session_state.texts:
        st.markdown("---")
        st.write("**📋 รายการข้อความ:**")
        for i, t in enumerate(st.session_state.texts):
            lbl = t['text'] if t['type'] == 'static' else f"📊 {t['column']}"
            col_act1, col_act2, col_act3 = st.columns([3, 1, 1])
            font_version_display = t.get('font_version', 'ใหม่')
            col_act1.write(f"{i+1}. {lbl} | ฟอนต์: {t['font_name']} (รุ่น{font_version_display}) | ขนาด: {t['size']}")
            
            if col_act2.button("✏️", key=f"edit_{i}"):
                st.session_state.edit_index = i
                st.rerun()
            
            if col_act3.button("🗑️", key=f"del_{i}"):
                st.session_state.texts.pop(i)
                if st.session_state.edit_index == i:
                    st.session_state.edit_index = None
                st.rerun()

# --- Export ---
if st.session_state.data is not None and not st.session_state.data.empty and st.session_state.texts:
    st.markdown("---")
    st.header("📦 สร้างและดาวน์โหลด")
    
    c1, c2 = st.columns(2)
    filename_col = c1.selectbox("เลือกคอลัมน์สำหรับชื่อไฟล์", st.session_state.data.columns)
    file_format = c2.radio("รูปแบบไฟล์", ["PNG", "PDF", "PowerPoint"], horizontal=True)
    
    if st.button("🚀 เริ่มสร้างทั้งหมด", type="primary"):
        with st.spinner("กำลังประมวลผล..."):
            if file_format == "PowerPoint":
                pptx_io = create_pptx_with_editable_text(
                    st.session_state.template,
                    st.session_state.texts,
                    st.session_state.data
                )
                st.success("✅ สร้างไฟล์ PowerPoint เรียบร้อย! (ข้อความแก้ไขได้)")
                st.download_button(
                    "📥 ดาวน์โหลด PowerPoint",
                    pptx_io.getvalue(),
                    "certificates.pptx",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for _, row in st.session_state.data.iterrows():
                        final_img = render_certificate(st.session_state.template, st.session_state.texts, row.to_dict())
                        img_io = BytesIO()
                        if file_format == "PNG":
                            final_img.save(img_io, format="PNG")
                            ext = "png"
                        else:
                            final_img.save(img_io, format="PDF", resolution=300.0)
                            ext = "pdf"
                        zf.writestr(f"{sanitize_filename(row[filename_col])}.{ext}", img_io.getvalue())
                st.success("✅ สร้างไฟล์ ZIP เรียบร้อย!")
                st.download_button(
                    f"📥 ดาวน์โหลด ZIP ({file_format})",
                    zip_buffer.getvalue(),
                    "certificates.zip",
                    "application/zip"
                )

# --- คำแนะนำ ---
with st.sidebar:
    st.markdown("---")
    st.markdown("""
    ### 📖 วิธีใช้
    1. อัปโหลดพื้นหลังเกียรติบัตร
    2. เลือกประเภทฟอนต์และอัปโหลด .ttf
    3. อัปโหลดไฟล์ Excel/CSV
    4. ใช้ปุ่มควบคุมเพื่อกำหนดพิกัด
    5. เพิ่ม/แก้ไขข้อความ
    6. สร้างและดาวน์โหลด
    """)
