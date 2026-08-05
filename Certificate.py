import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from io import BytesIO
import zipfile
import re
import os
import time

# --- นำเข้าไลบรารีสำหรับสร้าง PowerPoint ---
try:
    from pptx import Presentation
    from pptx.util import Pt
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

def fix_thai_text_old(text):
    """
    จัดตำแหน่งสระและวรรณยุกต์สำหรับ 'ฟอนต์รุ่นเก่า' (ใช้รหัส PUA)
    """
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
        
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
    
    # ญ และ ฐ เมื่อมีสระล่าง
    YO_YING_NO_BASE = '\uF70F'
    THO_THAN_NO_BASE = '\uF700'
    
    text = text.replace('ญุ', YO_YING_NO_BASE + 'ุ')
    text = text.replace('ญู', YO_YING_NO_BASE + 'ู')
    text = text.replace('ญฺ', YO_YING_NO_BASE + 'ฺ')
    text = text.replace('ฐุ', THO_THAN_NO_BASE + 'ุ')
    text = text.replace('ฐู', THO_THAN_NO_BASE + 'ู')
    text = text.replace('ฐฺ', THO_THAN_NO_BASE + 'ฺ')
    
    return text

def fix_thai_text_new(text):
    """สำหรับฟอนต์รุ่นใหม่ (OpenType)"""
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
    return text

def fix_thai_text(text, font_version="ใหม่"):
    """เลือกใช้ฟังก์ชันปรับแต่งภาษาไทยตามเวอร์ชันฟอนต์"""
    if font_version == "เก่า":
        return fix_thai_text_old(text)
    else:
        return fix_thai_text_new(text)

def get_font(font_name, size):
    """ดึงฟอนต์ตามชื่อและขนาดจากหน่วยความจำ"""
    try:
        if font_name and font_name in st.session_state.fonts_dict:
            font_data = st.session_state.fonts_dict[font_name]['data']
            return ImageFont.truetype(BytesIO(font_data), size)
    except Exception as e:
        pass
    
    # ใช้ฟอนต์ระบบสำรอง
    sys_path = get_system_font_path()
    if sys_path:
        try:
            return ImageFont.truetype(sys_path, size)
        except:
            return ImageFont.load_default()
    return ImageFont.load_default()

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '_', str(name)).strip() or "certificate"

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
        
        # ดึงประเภทฟอนต์
        f_info = st.session_state.fonts_dict.get(txt.get('font_name'), {'version': 'ใหม่'})
        content = fix_thai_text(content, f_info['version'])
            
        font = get_font(txt.get('font_name'), txt['size'])

        # วาดข้อความกึ่งกลาง
        draw.text((txt['x'], txt['y']), content, fill=txt['color'], font=font, anchor="mm")
    return img

def hex_to_rgb(hex_color):
    """แปลงสี Hex เป็น RGB tuple"""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return (0, 0, 0)

def create_pptx_with_editable_text(template_img, texts, data_df):
    """สร้าง PowerPoint ที่ข้อความสามารถแก้ไขได้"""
    prs = Presentation()
    
    img_width, img_height = template_img.size
    prs.slide_width = img_width * 9525
    prs.slide_height = img_height * 9525
    
    blank_slide_layout = prs.slide_layouts[6]
    
    for idx, row in data_df.iterrows():
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # แปะภาพพื้นหลัง
        bg_io = BytesIO()
        template_img.save(bg_io, format="PNG")
        bg_io.seek(0)
        slide.shapes.add_picture(bg_io, 0, 0, width=prs.slide_width, height=prs.slide_height)
        
        # เพิ่มข้อความ
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
            
            ppt_content = content
            font_size_pt = txt['size'] * 0.75
            
            center_x_emu = txt['x'] * 9525
            center_y_emu = txt['y'] * 9525
            
            box_width_emu = prs.slide_width
            box_height_emu = int(txt['size'] * 2.5 * 9525)
            
            left_emu = center_x_emu - (box_width_emu / 2)
            top_emu = center_y_emu - (box_height_emu / 2)
            
            txBox = slide.shapes.add_textbox(
                int(left_emu), 
                int(top_emu), 
                int(box_width_emu), 
                int(box_height_emu)
            )
            
            tf = txBox.text_frame
            tf.text = ppt_content
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            
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
    
    font_type = st.radio("ประเภทฟอนต์", ["ฟอนต์รุ่นใหม่ (OpenType)", "ฟอนต์รุ่นเก่า (PUA)"])
    uploaded_font = st.file_uploader("🔤 อัปโหลดฟอนต์ .ttf", type=['ttf'])
    
    if uploaded_font:
        f_name = uploaded_font.name.split('.')[0]
        v_type = "ใหม่" if "ใหม่" in font_type else "เก่า"
        if f_name not in st.session_state.fonts_dict:
            st.session_state.fonts_dict[f_name] = {
                'data': uploaded_font.getvalue(),
                'version': v_type
            }
            st.session_state.font_names.append(f_name)
            st.success(f"✅ เพิ่มฟอนต์ '{f_name}' (รุ่น {v_type})")
    
    if st.session_state.font_names:
        st.markdown("---")
        st.write("**📋 ฟอนต์ที่มี:**")
        for f in st.session_state.font_names:
            v = st.session_state.fonts_dict[f]['version']
            st.write(f"- {f} (รุ่น {v})")

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
    preview_row = None
    if st.session_state.data is not None and not st.session_state.data.empty:
        row_idx = st.number_input("ดูตัวอย่างจากแถวที่:", 0, len(st.session_state.data)-1, 0)
        preview_row = st.session_state.data.iloc[row_idx].to_dict()
    
    # สร้างรูปพรีวิว
    current_preview = render_certificate(st.session_state.template, st.session_state.texts, preview_row)
    
    st.markdown("**🖱️ คลิกที่รูปเพื่อกำหนดตำแหน่ง**")
    original_w, original_h = current_preview.size
    display_w = 700 
    ratio = original_w / display_w if original_w > display_w else 1.0
    display_img = current_preview.resize((display_w, int(original_h / ratio))) if original_w > display_w else current_preview
    
    # ใช้ streamlit_image_coordinates
    try:
        coords = streamlit_image_coordinates(display_img, key="coords")
        if coords:
            st.session_state.click_x = int(coords['x'] * ratio)
            st.session_state.click_y = int(coords['y'] * ratio)
    except Exception as e:
        st.warning("⚠️ ระบบคลิกยังไม่ทำงาน กรุณาใช้ slider แทน")
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.click_x = st.slider("X", 0, original_w, st.session_state.click_x)
        with col2:
            st.session_state.click_y = st.slider("Y", 0, original_h, st.session_state.click_y)
    
    st.info(f"📍 พิกัดปัจจุบัน: X={st.session_state.click_x}, Y={st.session_state.click_y}")

with col_form:
    st.subheader("➕ เพิ่มข้อความ")
    
    cols_options = st.session_state.data.columns.tolist() if st.session_state.data is not None else []
    
    with st.form("add_text_form", clear_on_submit=True):
        t_type = st.radio("ชนิดข้อความ", ["ดึงจากไฟล์รายชื่อ", "พิมพ์เอง"], horizontal=True)
        
        if t_type == "ดึงจากไฟล์รายชื่อ":
            if not cols_options:
                st.warning("⚠️ กรุณาอัปโหลดไฟล์รายชื่อก่อน")
                t_col = None
            else:
                t_col = st.selectbox("เลือกหัวข้อ (คอลัมน์)", cols_options)
            t_val = ""
        else:
            t_val = st.text_input("ข้อความที่ต้องการพิมพ์")
            t_col = None
            
        c1, c2 = st.columns(2)
        x_pos = c1.number_input("ตำแหน่ง X", value=st.session_state.click_x)
        y_pos = c2.number_input("ตำแหน่ง Y", value=st.session_state.click_y)
        
        f_size = st.slider("ขนาดฟอนต์", 10, 500, value=60)
        f_color = st.color_picker("เลือกสีข้อความ", value="#000000")
        
        if not st.session_state.font_names:
            st.warning("⚠️ กรุณาอัปโหลดฟอนต์ก่อน")
            selected_font = None
        else:
            selected_font = st.selectbox("เลือกฟอนต์", st.session_state.font_names)
        
        submit = st.form_submit_button("➕ เพิ่มข้อความ")
        if submit:
            if not selected_font:
                st.error("❌ กรุณาเลือกฟอนต์")
            elif t_type == "พิมพ์เอง" and not t_val:
                st.error("❌ กรุณาพิมพ์ข้อความ")
            elif t_type == "ดึงจากไฟล์รายชื่อ" and not t_col:
                st.error("❌ กรุณาเลือกคอลัมน์")
            else:
                st.session_state.texts.append({
                    'type': 'excel' if t_type == "ดึงจากไฟล์รายชื่อ" else 'static',
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

    # แสดงรายการข้อความ
    if st.session_state.texts:
        st.markdown("---")
        st.write("**📋 รายการข้อความ:**")
        for i, t in enumerate(st.session_state.texts):
            lbl = t['text'] if t['type'] == 'static' else f"📊 {t['column']}"
            col_del1, col_del2 = st.columns([4, 1])
            font_info = st.session_state.fonts_dict.get(t['font_name'], {})
            font_version = font_info.get('version', 'ใหม่')
            col_del1.write(f"{i+1}. {lbl} | ฟอนต์: {t['font_name']} (รุ่น{font_version}) | ขนาด: {t['size']} | พิกัด: ({t['x']}, {t['y']})")
            if col_del2.button("🗑️", key=f"del_{i}"):
                st.session_state.texts.pop(i)
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
                st.success("✅ สร้างไฟล์ PowerPoint เรียบร้อย!")
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
                            final_img.save(img_io, format="PDF", resolution=100.0)
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
    4. คลิกบนรูปเพื่อกำหนดพิกัด
    5. เพิ่มข้อความ
    6. สร้างและดาวน์โหลด
    """)
