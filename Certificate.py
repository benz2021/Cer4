import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from io import BytesIO
import zipfile
import re
import os

# --- นำเข้าไลบรารีสำหรับคลิกหาพิกัด ---
try:
    from streamlit_image_coordinates import streamlit_image_coordinates
except ImportError:
    st.error("⚠️ ไม่พบไลบรารี streamlit-image-coordinates")
    st.info("กรุณาเปิด Terminal แล้วพิมพ์: pip install streamlit-image-coordinates")
    st.stop()

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

def fix_thai_text_old(text):
    """
    จัดตำแหน่งสระและวรรณยุกต์สำหรับ 'ฟอนต์รุ่นเก่า' (ใช้รหัส PUA)
    รวมถึงการตัดฐาน ญ และ ฐ เมื่อมีสระล่าง
    """
    if not isinstance(text, str):
        return str(text) if pd.notna(text) else ""
        
    tone_marks = ['\u0e48', '\u0e49', '\u0e4a', '\u0e4b', '\u0e4c']
    
    # 1. วรรณยุกต์ที่ตามหลังสระบน
    upper_vowels = ['\u0e31', '\u0e34', '\u0e35', '\u0e36', '\u0e37', '\u0e4d']
    high_tone_marks = ['\uf713', '\uf714', '\uf715', '\uf716', '\uf717']
    
    for i, tone in enumerate(tone_marks):
        for vowel in upper_vowels:
            text = text.replace(vowel + tone, vowel + high_tone_marks[i])
    
    # 2. วรรณยุกต์ที่ตามหลัง ป ฝ ฟ
    tall_consonants = ['ป', 'ฝ', 'ฟ']
    left_tone_marks = ['\uf70a', '\uf70b', '\uf70c', '\uf70d', '\uf70e']
    
    for i, tone in enumerate(tone_marks):
        for cons in tall_consonants:
            text = text.replace(cons + tone, cons + left_tone_marks[i])
            
    # 3. สระอำ
    text = text.replace('\u0e4d\u0e32', '\u0e33')
    
    # 4. ญ และ ฐ เมื่อมีสระล่าง (ตัดฐาน)
    YO_YING_NO_BASE = '\uF70F'  # ญ (ตัดฐาน)
    THO_THAN_NO_BASE = '\uF700' # ฐ (ตัดฐาน)
    
    text = text.replace('ญุ', YO_YING_NO_BASE + 'ุ')
    text = text.replace('ญู', YO_YING_NO_BASE + 'ู')
    text = text.replace('ญฺ', YO_YING_NO_BASE + 'ฺ')
    text = text.replace('ฐุ', THO_THAN_NO_BASE + 'ุ')
    text = text.replace('ฐู', THO_THAN_NO_BASE + 'ู')
    text = text.replace('ฐฺ', THO_THAN_NO_BASE + 'ฺ')
    
    return text

def fix_thai_text_new(text):
    """
    สำหรับฟอนต์รุ่นใหม่ (OpenType) ไม่ต้องปรับแต่งอะไร
    เพียงแค่คืนค่าข้อความเดิม
    """
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

def render_certificate(template_img, texts, row_data=None, font_version="ใหม่"):
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
        
        # ปรับแต่งข้อความตามเวอร์ชันฟอนต์
        content = fix_thai_text(content, font_version)
            
        # ใช้ฟอนต์ที่บันทึกไว้สำหรับข้อความนี้โดยเฉพาะ
        font = get_font(txt.get('font_name'), txt['size'])

        # วัดขนาดและจัดกึ่งกลาง
        try:
            text_bbox = font.getbbox(content)
            text_width = text_bbox[2] - text_bbox[0]
        except:
            text_width = draw.textlength(content, font=font)

        # คำนวณจุดเริ่มต้น x
        start_x = txt['x'] - (text_width / 2)
        draw.text((start_x, txt['y']), content, fill=txt['color'], font=font, anchor="ls")
    return img

def hex_to_rgb(hex_color):
    """แปลงสี Hex เป็น RGB tuple"""
    hex_color = hex_color.lstrip('#')
    if len(hex_color) == 6:
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    return (0, 0, 0)

def create_pptx_with_editable_text(template_img, texts, data_df, font_version="ใหม่"):
    """สร้าง PowerPoint ที่ข้อความสามารถแก้ไขได้"""
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
            
            # สำหรับ PowerPoint ใช้ข้อความปกติ (ไม่ปรับแต่ง PUA)
            # เพราะ PowerPoint จะจัดการ rendering เอง
            ppt_content = content
            
            # คำนวณขนาดและตำแหน่ง
            font = get_font(txt.get('font_name'), txt['size'])
            
            try:
                bbox = font.getbbox(content)
                text_width_px = bbox[2] - bbox[0]
                text_height_px = bbox[3] - bbox[1]
            except:
                text_width_px = len(content) * txt['size'] * 0.6
                text_height_px = txt['size'] * 1.2
            
            # แปลงพิกัดจาก pixel เป็น EMUs
            center_x_emu = txt['x'] * 9525
            center_y_emu = txt['y'] * 9525
            
            padding = 20
            box_width_emu = int((text_width_px + padding * 2) * 9525)
            box_height_emu = int((text_height_px + padding * 2) * 9525)
            
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
            font_size_pt = txt['size'] * 0.75
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
    
    # เลือกเวอร์ชันฟอนต์
    font_version = st.radio(
        "ประเภทฟอนต์",
        ["ใหม่", "เก่า"],
        index=0,
        format_func=lambda x: "ฟอนต์รุ่นใหม่ (OpenType)" if x == "ใหม่" else "ฟอนต์รุ่นเก่า (PUA)"
    )
    st.session_state.font_version = font_version
    
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

col_img, col_form = st.columns([1.5, 1])

with col_img:
    st.markdown("**🖱️ คลิกลงบนรูปภาพเพื่อระบุตำแหน่ง X และ Y**")
    original_w, original_h = st.session_state.template.size
    display_w = 700 
    ratio = original_w / display_w if original_w > display_w else 1.0
    display_img = st.session_state.template.resize((display_w, int(original_h / ratio))) if original_w > display_w else st.session_state.template

    coords = streamlit_image_coordinates(display_img, key="target_clicker")
    if coords:
        st.session_state.click_x = int(coords['x'] * ratio)
        st.session_state.click_y = int(coords['y'] * ratio)

with col_form:
    st.info(f"💡 ใช้ฟอนต์รุ่น: **{font_version}**")
    
    with st.form("add_text_form", clear_on_submit=True):
        t_type = st.radio("ชนิดข้อความ", ["พิมพ์เอง", "ดึงจากไฟล์รายชื่อ"], horizontal=True)
        t_val = st.text_input("ข้อความ (ถ้าพิมพ์เอง)")
        t_col = st.selectbox("เลือกหัวข้อ (ถ้าดึงจากไฟล์)", st.session_state.data.columns if 'data' in st.session_state else ["กรุณาอัปโหลดไฟล์รายชื่อ"])
        
        c1, c2 = st.columns(2)
        x_pos = c1.number_input("ตำแหน่ง X", value=st.session_state.click_x)
        y_pos = c2.number_input("ตำแหน่ง Y", value=st.session_state.click_y)
        
        f_size = st.slider("ขนาดฟอนต์", 10, 500, value=60)
        f_color = st.color_picker("เลือกสีข้อความ", value="#000000")
        
        # เลือกฟอนต์ (บังคับเลือกจากที่อัปโหลด)
        if not st.session_state.font_names:
            st.warning("⚠️ กรุณาอัปโหลดฟอนต์ที่เมนูด้านซ้ายก่อน")
            selected_font = None
        else:
            selected_font = st.selectbox("เลือกฟอนต์สำหรับข้อความนี้", st.session_state.font_names)
        
        if st.form_submit_button("➕ เพิ่มข้อความลงเกียรติบัตร"):
            if selected_font:
                st.session_state.texts.append({
                    'type': 'static' if "พิมพ์เอง" in t_type else 'excel',
                    'text': t_val, 'column': t_col,
                    'x': x_pos, 'y': y_pos,
                    'size': f_size, 'color': f_color,
                    'font_name': selected_font
                })
                st.success("เพิ่มข้อความสำเร็จ!")
                st.rerun()
            else:
                st.error("กรุณาอัปโหลดและเลือกฟอนต์ก่อนเพิ่มข้อความ")

st.markdown("---")
st.header("👁️ ตรวจสอบและพรีวิว")

if st.session_state.texts:
    for i, t in enumerate(st.session_state.texts):
        lbl = t['text'] if t['type'] == 'static' else f"คอลัมน์: {t['column']}"
        cols = st.columns([4, 1])
        cols[0].write(f"**{i+1}. {lbl}** | ฟอนต์: {t['font_name']} | ขนาด: {t['size']} | พิกัด: ({t['x']}, {t['y']})")
        if cols[1].button("🗑️ ลบ", key=f"del_{i}"):
            st.session_state.texts.pop(i)
            st.rerun()

    preview_row = st.session_state.data.iloc[st.number_input("ดูตัวอย่างจากรายชื่อแถวที่:", 0, max(0, len(st.session_state.data)-1), 0)].to_dict() if 'data' in st.session_state else None
    preview_img = render_certificate(
        st.session_state.template, 
        st.session_state.texts, 
        preview_row,
        st.session_state.font_version
    )
    st.image(preview_img, width=700)
else:
    st.info("ยังไม่มีข้อความถูกเพิ่ม")

# --- Export ---
if 'data' in st.session_state and st.session_state.texts:
    st.markdown("---")
    st.header("📦 สร้างและดาวน์โหลด")
    
    st.info(f"💡 ใช้ฟอนต์รุ่น: **{font_version}** | PowerPoint: ข้อความแก้ไขได้")
    
    c1, c2 = st.columns(2)
    filename_col = c1.selectbox("เลือกคอลัมน์ที่จะใช้เป็นชื่อไฟล์ (ไม่ใช้กับ PowerPoint)", st.session_state.data.columns)
    
    file_format = c2.radio("เลือกรูปแบบไฟล์ส่งออก", ["PNG", "PDF", "PowerPoint"], horizontal=True)
    
    if st.button("🚀 เริ่มสร้างเกียรติบัตรทั้งหมด", type="primary"):
        with st.spinner("กำลังประมวลผล..."):
            
            # ---------------------------------------------------------
            # กรณีที่ผู้ใช้เลือกส่งออกแบบ PowerPoint
            # ---------------------------------------------------------
            if file_format == "PowerPoint":
                pptx_io = create_pptx_with_editable_text(
                    st.session_state.template,
                    st.session_state.texts,
                    st.session_state.data,
                    st.session_state.font_version
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
                        final_img = render_certificate(
                            st.session_state.template, 
                            st.session_state.texts, 
                            row.to_dict(),
                            st.session_state.font_version
                        )
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
    2. เลือกประเภทฟอนต์ (ใหม่/เก่า)
    3. อัปโหลดฟอนต์ .ttf
    4. อัปโหลดไฟล์ Excel/CSV
    5. คลิกบนรูปเพื่อกำหนดพิกัด
    6. พิมพ์ข้อความและเพิ่ม
    7. สร้างและดาวน์โหลด
    
    💡 **ฟอนต์ใหม่:** ไม่ต้องปรับแต่ง
    💡 **ฟอนต์เก่า:** ปรับสระลอย PUA
    """)
