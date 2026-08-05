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
    from pptx.util import Pt, Inches
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.dml.color import RGBColor
except ImportError:
    st.error("⚠️ ไม่พบไลบรารี python-pptx (สำหรับสร้าง PowerPoint)")
    st.info("กรุณาเปิด Terminal แล้วพิมพ์: pip install python-pptx")
    st.stop()

# ==========================================
# 🛠️ HELPER FUNCTIONS
# ==========================================

def get_system_font_path():
    """ค้นหาฟอนต์ที่ปรับขนาดได้ในระบบเพื่อใช้เป็นสำรอง"""
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

def fix_thai_text(text):
    """จัดตำแหน่งสระและวรรณยุกต์ภาษาไทยให้ถูกต้อง (เวอร์ชันปรับปรุง)"""
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    
    # จัดการสระและวรรณยุกต์
    tone_marks = ['\u0e48', '\u0e49', '\u0e4a', '\u0e4b', '\u0e4c']
    upper_vowels = ['\u0e31', '\u0e34', '\u0e35', '\u0e36', '\u0e37', '\u0e4d']
    high_tone_marks = ['\uf713', '\uf714', '\uf715', '\uf716', '\uf717']
    
    for i, tone in enumerate(tone_marks):
        for vowel in upper_vowels:
            text = text.replace(vowel + tone, vowel + high_tone_marks[i])
    
    tall_consonants = ['ป', 'ฝ', 'ฟ']
    left_tone_marks = ['\uf70a', '\uf70b', '\uf70c', '\uf70d', '\uf70e']
    
    for i, tone in enumerate(tone_marks):
        for cons in tall_consonants:
            text = text.replace(cons + tone, cons + left_tone_marks[i])
    
    text = text.replace('\u0e4d\u0e32', '\u0e33')
    return text

def get_font(font_name, size):
    """ดึงฟอนต์ตามชื่อและขนาดจากหน่วยความจำ"""
    if font_name in st.session_state.fonts_dict:
        font_data = st.session_state.fonts_dict[font_name]
        try:
            return ImageFont.truetype(BytesIO(font_data), size)
        except Exception as e:
            st.error(f"❌ ไม่สามารถโหลดฟอนต์ '{font_name}': {e}")
    
    # หากไม่พบฟอนต์ที่ระบุ พยายามใช้ฟอนต์ระบบที่ปรับขนาดได้
    sys_path = get_system_font_path()
    if sys_path:
        return ImageFont.truetype(sys_path, size)
    return ImageFont.load_default()

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "certificate"

def fix_thai_baseless_chars(text):
    """
    ฟังก์ชันสำหรับตัดฐาน ญ และ ฐ เมื่อมีสระล่าง (ุ, ู, ฺ)
    เพื่อเตรียมข้อความก่อนนำไปใช้กับ PIL.ImageDraw หรือสร้างเอกสาร
    """
    if not isinstance(text, str):
        return text

    # รหัส Unicode สำหรับตัวอักษรพิเศษ (PUA) ที่ไม่มีฐาน
    YO_YING_NO_BASE = '\uF70F'  # ญ (ตัดฐาน)
    THO_THAN_NO_BASE = '\uF700' # ฐ (ตัดฐาน)

    # กรณี ญ.หญิง
    text = text.replace('ญุ', YO_YING_NO_BASE + 'ุ')
    text = text.replace('ญู', YO_YING_NO_BASE + 'ู')
    text = text.replace('ญฺ', YO_YING_NO_BASE + 'ฺ') # สระพินทุ (จุดด้านล่าง)

    # แถม: กรณี ฐ.ฐาน (ใช้หลักการเดียวกัน)
    text = text.replace('ฐุ', THO_THAN_NO_BASE + 'ุ')
    text = text.replace('ฐู', THO_THAN_NO_BASE + 'ู')
    text = text.replace('ฐฺ', THO_THAN_NO_BASE + 'ฺ')

    return text

def prepare_thai_text(text):
    """เตรียมข้อความภาษาไทยให้แสดงผลถูกต้อง"""
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    
    # แก้สระภาษาไทยและวรรณยุกต์
    text = fix_thai_text(text)
    # ตัดฐาน ญ, ฐ เมื่อมีสระล่าง
    text = fix_thai_baseless_chars(text)
    return text

def render_certificate(template_img, texts, row_data=None):
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
        
        if not content: continue
        
        # เตรียมข้อความภาษาไทย
        content = prepare_thai_text(content)
            
        # ใช้ฟอนต์ที่บันทึกไว้สำหรับข้อความนี้โดยเฉพาะ
        font = get_font(txt.get('font_name'), txt['size'])

        # วัดขนาดและจัดกึ่งกลาง
        try:
            text_bbox = font.getbbox(content)
            text_width = text_bbox[2] - text_bbox[0]
        except:
            text_width = draw.textlength(content, font=font)

        # คำนวณจุดเริ่มต้น x โดยให้ข้อความกึ่งกลางที่ตำแหน่ง x ที่เลือก
        start_x = txt['x'] - (text_width / 2)
        
        # วาดข้อความ
        draw.text((start_x, txt['y']), content, fill=txt['color'], font=font, anchor="ls")
    return img

def hex_to_rgb(hex_color):
    """แปลงสี Hex เป็น RGB tuple"""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return (0, 0, 0)

def create_pptx_with_editable_text(template_img, texts, data_df):
    """สร้าง PowerPoint ที่ข้อความสามารถแก้ไขได้ และรองรับภาษาไทย"""
    prs = Presentation()
    
    # กำหนดขนาดสไลด์
    img_width, img_height = template_img.size
    prs.slide_width = img_width * 9525
    prs.slide_height = img_height * 9525
    
    blank_slide_layout = prs.slide_layouts[6]
    
    for idx, row in data_df.iterrows():
        # สร้างสไลด์ใหม่
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # 1. แปะภาพพื้นหลัง
        bg_img = template_img.copy()
        if bg_img.mode != 'RGB':
            bg_img = bg_img.convert('RGB')
        bg_io = BytesIO()
        bg_img.save(bg_io, format="PNG")
        bg_io.seek(0)
        slide.shapes.add_picture(bg_io, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        # 2. เพิ่มข้อความแต่ละรายการ
        for txt in texts:
            # ดึงข้อความ
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
            
            # เตรียมข้อความภาษาไทย (สำหรับแสดงใน PowerPoint)
            # สำหรับ PowerPoint ใช้ข้อความปกติ (ไม่ต้องใช้ PUA)
            ppt_content = content
            
            # คำนวณขนาดและตำแหน่ง
            # ใช้ฟอนต์เพื่อวัดขนาดข้อความ (ใช้ฟอนต์ที่อัปโหลด)
            font = get_font(txt.get('font_name'), txt['size'])
            
            try:
                # วัดขนาดข้อความด้วย PIL
                bbox = font.getbbox(content)
                text_width_px = bbox[2] - bbox[0]
                text_height_px = bbox[3] - bbox[1]
            except:
                text_width_px = len(content) * txt['size'] * 0.6
                text_height_px = txt['size'] * 1.2
            
            # แปลงพิกัดจาก pixel เป็น EMUs (1 pixel = 9525 EMUs)
            # จุดกึ่งกลางของข้อความ
            center_x_emu = txt['x'] * 9525
            center_y_emu = txt['y'] * 9525
            
            # เพิ่ม padding เพื่อให้ข้อความไม่ติดขอบ
            padding = 20
            box_width_emu = int((text_width_px + padding * 2) * 9525)
            box_height_emu = int((text_height_px + padding * 2) * 9525)
            
            # คำนวณตำแหน่งซ้ายบนของกล่องข้อความ (ให้กึ่งกลาง)
            left_emu = center_x_emu - (box_width_emu / 2)
            top_emu = center_y_emu - (box_height_emu / 2)
            
            # สร้างกล่องข้อความ
            txBox = slide.shapes.add_textbox(
                int(left_emu), 
                int(top_emu), 
                int(box_width_emu), 
                int(box_height_emu)
            )
            
            # ตั้งค่าข้อความ
            tf = txBox.text_frame
            tf.text = ppt_content
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            # จัดกึ่งกลางแนวนอน
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            
            # ตั้งค่ารูปแบบฟอนต์
            run = p.runs[0]
            font_size_pt = txt['size'] * 0.75  # แปลง pixel เป็น point
            run.font.size = Pt(font_size_pt)
            
            # ตั้งชื่อฟอนต์
            if txt.get('font_name'):
                run.font.name = txt['font_name']
            
            # ตั้งค่าสี
            rgb = hex_to_rgb(txt['color'])
            run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    pptx_io = BytesIO()
    prs.save(pptx_io)
    pptx_io.seek(0)
    return pptx_io

def get_text_size(text, font_name, size):
    """คำนวณขนาดข้อความเพื่อแสดงตัวอย่าง"""
    font = get_font(font_name, size)
    try:
        bbox = font.getbbox(text)
        return bbox[2] - bbox[0], bbox[3] - bbox[1]
    except:
        return len(text) * size * 0.6, size * 1.2

# ==========================================
# 🎨 UI - STREAMLIT APP
# ==========================================
st.set_page_config(page_title="Auto Cert Pro", layout="wide")

# ตั้งค่า Session State
if "click_x" not in st.session_state: st.session_state.click_x = 0
if "click_y" not in st.session_state: st.session_state.click_y = 0
if 'texts' not in st.session_state: st.session_state.texts = []
if 'fonts_dict' not in st.session_state: st.session_state.fonts_dict = {}
if 'font_names' not in st.session_state: st.session_state.font_names = []

st.title("📜 Auto Certificate Generator")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1️⃣ อัปโหลดไฟล์")
    
    # 1. Template
    template_file = st.file_uploader("🖼️ พื้นหลังเกียรติบัตร (JPG/PNG)", type=['jpg', 'jpeg', 'png'])
    if template_file:
        st.session_state.template = Image.open(template_file)

    # 2. Font Management
    st.markdown("---")
    st.header("2️⃣ จัดการฟอนต์")
    uploaded_font = st.file_uploader("🔤 อัปโหลดฟอนต์ใหม่ (.ttf)", type=['ttf'])
    if uploaded_font:
        f_name = uploaded_font.name.split('.')[0]
        if f_name not in st.session_state.fonts_dict:
            st.session_state.fonts_dict[f_name] = uploaded_font.getvalue()
            st.session_state.font_names.append(f_name)
            st.success(f"✅ เพิ่มฟอนต์ '{f_name}' แล้ว")
    
    # แสดงฟอนต์ที่มี
    if st.session_state.font_names:
        st.markdown("---")
        st.write("**📋 ฟอนต์ที่มี:**")
        for f in st.session_state.font_names:
            st.write(f"- {f}")

    # 3. Data
    st.markdown("---")
    st.header("3️⃣ รายชื่อข้อมูล")
    data_file = st.file_uploader("📊 ไฟล์ Excel/CSV", type=['xlsx', 'xls', 'csv'])
    if data_file:
        if data_file.name.endswith('.csv'):
            st.session_state.data = pd.read_csv(data_file)
        else:
            st.session_state.data = pd.read_excel(data_file)
        st.success(f"✅ โหลดข้อมูล {len(st.session_state.data)} รายการ")

if 'template' not in st.session_state:
    st.info("👈 กรุณาอัปโหลด 'พื้นหลังเกียรติบัตร' ที่เมนูด้านซ้ายเพื่อเริ่มต้น")
    st.stop()

# --- MAIN AREA ---
st.header("📍 กำหนดตำแหน่งและข้อความ")
st.markdown("💡 **คำแนะนำ:** คลิกที่ปุ่มเพื่อเลื่อนตำแหน่ง หรือปรับค่า X, Y โดยตรง จุดที่เลือกคือ **กึ่งกลาง** ของข้อความ")

col_img, col_form = st.columns([1.5, 1])

with col_img:
    st.markdown("**🖱️ คลิกที่ปุ่มเพื่อปรับตำแหน่ง (จุดกึ่งกลางข้อความ)**")
    
    original_w, original_h = st.session_state.template.size
    display_w = 700 
    ratio = original_w / display_w if original_w > display_w else 1.0
    display_h = int(original_h / ratio) if original_w > display_w else original_h
    display_img = st.session_state.template.resize((display_w, display_h)) if original_w > display_w else st.session_state.template
    
    # แสดงรูปภาพพร้อมเครื่องหมายกากบาท
    img_with_marker = display_img.copy()
    draw = ImageDraw.Draw(img_with_marker)
    marker_x = int(st.session_state.click_x / ratio)
    marker_y = int(st.session_state.click_y / ratio)
    
    # วาดเครื่องหมายกากบาทที่ตำแหน่งกึ่งกลาง
    marker_size = 25
    draw.line([(marker_x - marker_size, marker_y), (marker_x + marker_size, marker_y)], fill='red', width=3)
    draw.line([(marker_x, marker_y - marker_size), (marker_x, marker_y + marker_size)], fill='red', width=3)
    draw.ellipse([(marker_x - 6, marker_y - 6), (marker_x + 6, marker_y + 6)], fill='red', outline='white', width=2)
    
    # แสดงข้อความตัวอย่างที่ตำแหน่งกึ่งกลาง
    if st.session_state.texts and st.session_state.font_names:
        last_text = st.session_state.texts[-1]
        if last_text.get('font_name'):
            font = get_font(last_text['font_name'], last_text['size'])
            content = last_text['text'] if last_text['type'] == 'static' else "ตัวอย่าง"
            try:
                bbox = font.getbbox(content)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                
                # แสดงกรอบข้อความที่ตำแหน่ง
                start_x = marker_x - (text_width / 2 / ratio)
                start_y = marker_y - (text_height / 2 / ratio)
                draw.rectangle([start_x, start_y, start_x + text_width/ratio, start_y + text_height/ratio], 
                             outline='blue', width=2)
                draw.text((marker_x - (text_width / 2 / ratio), marker_y - (text_height / 2 / ratio)), 
                         content, fill='blue', font=font)
            except:
                pass
    
    st.image(img_with_marker, use_column_width=True)
    
    # ระบบควบคุมตำแหน่ง
    st.markdown("---")
    st.markdown("**🎯 ปรับตำแหน่งกึ่งกลางข้อความ**")
    
    # แถวบน
    col_btn1, col_btn2, col_btn3 = st.columns(3)
    with col_btn1:
        if st.button("⬆️ บน", use_container_width=True):
            st.session_state.click_y = max(0, st.session_state.click_y - 20)
            st.rerun()
    with col_btn2:
        if st.button("⬅️ ซ้าย", use_container_width=True):
            st.session_state.click_x = max(0, st.session_state.click_x - 20)
            st.rerun()
    with col_btn3:
        if st.button("➡️ ขวา", use_container_width=True):
            st.session_state.click_x = min(original_w, st.session_state.click_x + 20)
            st.rerun()
    
    # แถวล่าง
    col_btn4, col_btn5, col_btn6 = st.columns(3)
    with col_btn4:
        if st.button("⬇️ ล่าง", use_container_width=True):
            st.session_state.click_y = min(original_h, st.session_state.click_y + 20)
            st.rerun()
    with col_btn5:
        if st.button("🎯 กลางภาพ", use_container_width=True):
            st.session_state.click_x = original_w // 2
            st.session_state.click_y = original_h // 2
            st.rerun()
    with col_btn6:
        if st.button("🔄 รีเซ็ต", use_container_width=True):
            st.session_state.click_x = 0
            st.session_state.click_y = 0
            st.rerun()
    
    # แสดงตำแหน่งปัจจุบัน
    st.markdown(f"**📍 กึ่งกลางข้อความอยู่ที่: X = {st.session_state.click_x}, Y = {st.session_state.click_y}**")
    
    # แสดงขนาดภาพ
    st.caption(f"ขนาดภาพต้นฉบับ: {original_w} x {original_h} พิกเซล")

with col_form:
    with st.form("add_text_form", clear_on_submit=True):
        st.markdown("**✏️ เพิ่มข้อความ**")
        
        t_type = st.radio("ชนิดข้อความ", ["พิมพ์เอง", "ดึงจากไฟล์รายชื่อ"], horizontal=True)
        
        if t_type == "พิมพ์เอง":
            t_val = st.text_input("ข้อความ", placeholder="พิมพ์ข้อความที่นี่")
            t_col = ""
        else:
            if 'data' in st.session_state and st.session_state.data is not None:
                t_col = st.selectbox("เลือกหัวข้อ", st.session_state.data.columns)
                t_val = ""
                if len(st.session_state.data) > 0:
                    st.caption(f"📝 ตัวอย่าง: {st.session_state.data[t_col].iloc[0]}")
            else:
                st.warning("⚠️ กรุณาอัปโหลดไฟล์รายชื่อก่อน")
                t_col = ""
                t_val = ""
        
        c1, c2 = st.columns(2)
        x_pos = c1.number_input("ตำแหน่ง X (กึ่งกลาง)", value=st.session_state.click_x)
        y_pos = c2.number_input("ตำแหน่ง Y (กึ่งกลาง)", value=st.session_state.click_y)
        
        f_size = st.slider("ขนาดฟอนต์", 10, 500, value=60)
        f_color = st.color_picker("เลือกสีข้อความ", value="#000000")
        
        # เลือกฟอนต์
        if not st.session_state.font_names:
            st.warning("⚠️ กรุณาอัปโหลดฟอนต์ที่เมนูด้านซ้ายก่อน")
            selected_font = None
        else:
            selected_font = st.selectbox("เลือกฟอนต์", st.session_state.font_names)
        
        # แสดงตัวอย่างขนาดข้อความ
        if selected_font and (t_val or t_col):
            sample_text = t_val if t_type == "พิมพ์เอง" else "ตัวอย่างข้อความ"
            if sample_text:
                try:
                    font = get_font(selected_font, f_size)
                    bbox = font.getbbox(sample_text)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                    st.caption(f"📏 ขนาดข้อความ: {text_w} x {text_h} พิกเซล (กว้าง x สูง)")
                except:
                    pass
        
        if st.form_submit_button("➕ เพิ่มข้อความลงเกียรติบัตร"):
            if selected_font:
                if t_type == "พิมพ์เอง" and not t_val:
                    st.error("❌ กรุณาพิมพ์ข้อความ")
                elif t_type == "ดึงจากไฟล์รายชื่อ" and not t_col:
                    st.error("❌ กรุณาเลือกหัวข้อ")
                else:
                    st.session_state.texts.append({
                        'type': 'static' if t_type == "พิมพ์เอง" else 'excel',
                        'text': t_val, 
                        'column': t_col,
                        'x': x_pos, 
                        'y': y_pos,
                        'size': f_size, 
                        'color': f_color,
                        'font_name': selected_font
                    })
                    st.success("✅ เพิ่มข้อความสำเร็จ!")
                    st.rerun()
            else:
                st.error("❌ กรุณาอัปโหลดและเลือกฟอนต์ก่อนเพิ่มข้อความ")

st.markdown("---")
st.header("👁️ ตรวจสอบและพรีวิว")

if st.session_state.texts:
    for i, t in enumerate(st.session_state.texts):
        lbl = t['text'] if t['type'] == 'static' else f"📊 {t['column']}"
        cols = st.columns([4, 1])
        cols[0].write(f"**{i+1}. {lbl}** | ฟอนต์: {t['font_name']} | ขนาด: {t['size']} | กึ่งกลาง: ({t['x']}, {t['y']})")
        if cols[1].button("🗑️ ลบ", key=f"del_{i}"):
            st.session_state.texts.pop(i)
            st.rerun()

    # แสดงตัวอย่าง
    if 'data' in st.session_state and st.session_state.data is not None and len(st.session_state.data) > 0:
        row_idx = st.number_input("ดูตัวอย่างจากรายชื่อแถวที่:", 0, len(st.session_state.data)-1, 0)
        preview_row = st.session_state.data.iloc[row_idx].to_dict()
        preview_img = render_certificate(st.session_state.template, st.session_state.texts, preview_row)
        st.image(preview_img, width=700)
    else:
        preview_img = render_certificate(st.session_state.template, st.session_state.texts, None)
        st.image(preview_img, width=700)
else:
    st.info("💡 ยังไม่มีข้อความถูกเพิ่ม")

# --- Export ---
if 'data' in st.session_state and st.session_state.data is not None and st.session_state.texts:
    st.markdown("---")
    st.header("📦 สร้างและดาวน์โหลด")
    
    st.info("💡 **PowerPoint:** ข้อความจะอยู่ในกล่องข้อความที่แก้ไขได้ พร้อมฟอนต์และสีตามที่กำหนด")
    
    c1, c2 = st.columns(2)
    filename_col = c1.selectbox("เลือกคอลัมน์ที่จะใช้เป็นชื่อไฟล์ (ไม่ใช้กับ PowerPoint)", st.session_state.data.columns)
    
    file_format = c2.radio("เลือกรูปแบบไฟล์ส่งออก", ["PNG", "PDF", "PowerPoint"], horizontal=True)
    
    if st.button("🚀 เริ่มสร้างเกียรติบัตรทั้งหมด", type="primary"):
        with st.spinner("กำลังประมวลผล..."):
            
            # ---------------------------------------------------------
            # กรณีที่ผู้ใช้เลือกส่งออกแบบ PowerPoint (ข้อความแก้ไขได้)
            # ---------------------------------------------------------
            if file_format == "PowerPoint":
                pptx_io = create_pptx_with_editable_text(
                    st.session_state.template, 
                    st.session_state.texts, 
                    st.session_state.data
                )
                
                st.success("✅ สร้างไฟล์ PowerPoint เรียบร้อย! (ข้อความแก้ไขได้)")
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ PowerPoint (.pptx)", 
                    data=pptx_io.getvalue(), 
                    file_name="certificates.pptx", 
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            # ---------------------------------------------------------
            # กรณีที่ผู้ใช้เลือกส่งออกแบบ PNG หรือ PDF
            # ---------------------------------------------------------
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as zf:
                    for _, row in st.session_state.data.iterrows():
                        final_img = render_certificate(st.session_state.template, st.session_state.texts, row.to_dict())
                        img_io = BytesIO()
                        
                        if file_format == "PNG":
                            final_img.save(img_io, format="PNG")
                            ext = "png"
                        else:  # PDF
                            final_img.save(img_io, format="PDF")
                            ext = "pdf"
                            
                        zf.writestr(f"{sanitize_filename(row[filename_col])}.{ext}", img_io.getvalue())
                        
                st.success("✅ สร้างไฟล์ ZIP ทั้งหมดเรียบร้อย!")
                st.download_button(
                    f"📥 ดาวน์โหลดไฟล์ทั้งหมด ({file_format} ใน ZIP)", 
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
    2. อัปโหลดฟอนต์ .ttf
    3. อัปโหลดไฟล์ Excel/CSV (ถ้ามี)
    4. ปรับตำแหน่งกึ่งกลางด้วยปุ่ม หรือใส่ค่า X, Y
    5. พิมพ์ข้อความและเพิ่มลงในเกียรติบัตร
    6. กดสร้างและดาวน์โหลด
    
    💡 **PowerPoint:** ข้อความสามารถแก้ไขได้
    """)
