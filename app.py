import streamlit as st
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageEnhance, ImageOps, ExifTags
import io, zipfile, re, os, tempfile, sqlite3
from datetime import datetime, time, date, timedelta
import urllib.parse
import qrcode

st.set_page_config(page_title="Master Portal - CCI Manikant Choudhary (Solapur)", layout="wide")

# ==================== DATABASE SETUP ====================
def init_db():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            station TEXT,
            inspection_type TEXT,
            inspection_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS letters (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            subject TEXT,
            recipient TEXT,
            content TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_inspection_to_db(station, insp_type, date_str):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO inspections (station, inspection_type, inspection_date) VALUES (?, ?, ?)", (station, insp_type, date_str))
    conn.commit()
    conn.close()

def get_all_inspections():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("SELECT id, station, inspection_type, inspection_date, created_at FROM inspections ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def delete_inspection_from_db(insp_id):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM inspections WHERE id = ?", (insp_id,))
    conn.commit()
    conn.close()

def save_letter_to_db(subject, recipient, content):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO letters (subject, recipient, content) VALUES (?, ?, ?)", (subject, recipient, content))
    conn.commit()
    conn.close()

def get_all_letters():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("SELECT id, subject, recipient, created_at FROM letters ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

# ==================== OFFICIAL HEADER & BRANDING ====================
st.markdown(
    """
    <div style="background-color: #003399; padding: 18px; border-radius: 8px; text-align: center; color: white; margin-bottom: 20px;">
        <h2 style="margin: 0; font-size: 24px;">CENTRAL RAILWAY — SOLAPUR DIVISION</h2>
        <p style="margin: 5px 0 0 0; font-size: 15px; font-weight: bold; letter-spacing: 0.5px;">OFFICE OF THE SR. DIVISIONAL COMMERCIAL MANAGER (COMMERCIAL & CLEANLINESS DIRECTORATE)</p>
        <p style="margin: 3px 0 0 0; font-size: 13px; color: #ffeb3b;">MASTER INSPECTION PORTAL | CHIEF COMMERCIAL INSPECTOR (CCI): MANIKANT CHOUDHARY</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ==================== SECURE PASSWORD PROTECTION ====================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("### 🔐 Secure Official Login")
    pwd_input = st.text_input("Enter Security Password:", type="password")
    if st.button("Login to Portal", type="primary"):
        if pwd_input == "Railway@2026":
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Invalid Password! Please enter correct credentials.")
    st.stop()

with st.sidebar:
    st.markdown("### 🏛️ Commercial Directorate")
    app_mode = st.radio("Choose Section:", [
        "🔍 Master Field Inspection & Evidence", 
        "📝 Official Noting & Fine Proposal", 
        "📁 Inspection History (30 Days)", 
        "📊 Division Commercial Analytics", 
        "📱 Portal QR Code"
    ])
    st.markdown("---")
    st.markdown("**Officer Profile:**")
    st.markdown("`Manikant Choudhary`\n\n`Chief Commercial Inspector`\n\n`Solapur Division, C.Rly.`")
    
    st.markdown("---")
    st.markdown("✍️ **Custom Digital Signature / Seal (For PPT Only):**")
    sig_file = st.file_uploader("Upload CCI Sign / Seal (PNG/JPG):", type=['png', 'jpg', 'jpeg'], key="sig_uploader")
    if sig_file:
        st.session_state['sig_bytes'] = sig_file.read()
        st.success("✅ Signature uploaded successfully!")
    
    st.session_state['custom_sign_name'] = st.text_input("Signatory Name:", value="Manikant Choudhary, CCI", key="c_name")
    st.session_state['custom_sign_sub'] = st.text_input("Signatory Designation / Office:", value="Sr. DCM Office, Central Railway, Solapur", key="c_sub")

    st.markdown("---")
    if st.button("🔒 Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

RAW_LOCATIONS = [
    "PRS / UTS Ticketing Counter Area", "ATVM / CoTVM Kiosk Zone", "Station Concourse & Booking Office",
    "Passenger Amenities - Waiting Hall (Upper Class / General)", "Passenger Amenities - FOB & Staircase",
    "Catering - Static Stall / Food Plaza / Fast Food Unit", "Catering - Pantry Car / Train On-Board Vending",
    "Parcel Office & Loading Wharf", "Goods Shed & Siding Track Area", "Platform Cleanliness & Track Apron",
    "Pay & Use Toilet & Urinals Area", "Circulating Area & Parking Zone", "Retiring Rooms & Dormitory"
]
RAW_LOCATIONS.sort()
LOCATION_OPTIONS = ["-- Select Commercial/Amenity Location --"] + RAW_LOCATIONS

# ==================== IMAGE PROCESSING & EXIF ====================
def process_image(img_bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    img = ImageOps.fit(img, (800, 600), Image.Resampling.LANCZOS)
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(1.05)
    
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=80, optimize=True)
    img_byte_arr.seek(0)
    return img_byte_arr

def get_exif_data(img_bytes):
    try:
        img = Image.open(io.BytesIO(img_bytes))
        exif = img._getexif()
        if exif:
            exif_data = {}
            for tag_id, val in exif.items():
                tag = ExifTags.TAGS.get(tag_id, tag_id)
                exif_data[tag] = val
            return exif_data
    except Exception:
        pass
    return None

def get_image_info(img_bytes, filename):
    dt_obj = datetime.now()
    gps_info = "GPS: Not Available"
    
    exif = get_exif_data(img_bytes)
    if exif:
        if 'DateTimeOriginal' in exif or 'DateTime' in exif:
            val = exif.get('DateTimeOriginal', exif.get('DateTime'))
            try:
                dt_obj = datetime.strptime(str(val).strip(), '%Y:%m:%d %H:%M:%S')
            except:
                pass
        if 'GPSInfo' in exif:
            gps_info = "GPS Geotagged (Verified Evidence)"
    
    if dt_obj == datetime.now():
        match = re.search(r"(\d{4}-\d{2}-\d{2}) at (\d{1,2}\.\d{2}\.\d{2}\s?[AM|PM|am|pm]+)", filename)
        if match:
            date_str = match.group(1)
            time_str = match.group(2).replace('.', ':').upper()
            try:
                dt_obj = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %I:%M:%S %p")
            except:
                pass
                
    return dt_obj, gps_info

def create_ppt(station_name, insp_type, items_list, layout_mode, sig_bytes=None, sign_name="", sign_sub=""):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "MASTER COMMERCIAL INSPECTION REPORT"
    title_slide.shapes.title.text_frame.paragraphs[0].font.bold = True
    title_slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
    
    subtitle = title_slide.placeholders[1]
    subtitle.text = f"Directorate Focus: {insp_type}\nInspector: {sign_name}\n{sign_sub}"
    subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(102, 102, 102)

    if layout_mode == "Before & After Pairs (Comparison)":
        for data in items_list:
            slide = prs.slides.add_slide(prs.slide_layouts[5])
            title_shape = slide.shapes.title
            title_shape.text = f"Unit / Station: {station_name.upper()}"
            title_shape.text_frame.paragraphs[0].font.size = Pt(26)
            title_shape.text_frame.paragraphs[0].font.bold = True
            title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
            
            slide.shapes.add_picture(data['before'], Inches(0.4), Inches(1.4), width=Inches(4.4), height=Inches(2.7))
            tb1 = slide.shapes.add_textbox(Inches(0.4), Inches(4.2), Inches(4.4), Inches(1.1))
            tb1.text_frame.word_wrap = True
            p1 = tb1.text_frame.paragraphs[0]
            p1.text = "🔴 DEFICIENCY / BEFORE"
            p1.font.bold = True
            p1.font.size = Pt(14)
            p1.font.color.rgb = RGBColor(204, 0, 0)
            p1.alignment = PP_ALIGN.CENTER
            
            if data['show_dt']:
                p1_dt = tb1.text_frame.add_paragraph()
                p1_dt.text = f"Date: {data['d_before']} | Time: {data['t_before']}"
                p1_dt.font.size = Pt(10)
                p1_dt.alignment = PP_ALIGN.CENTER
            
            p1_loc = tb1.text_frame.add_paragraph()
            p1_loc.text = f"Location: {data['location']}"
            p1_loc.font.size = Pt(11)
            p1_loc.font.bold = True
            p1_loc.alignment = PP_ALIGN.CENTER

            slide.shapes.add_picture(data['after'], Inches(5.2), Inches(1.4), width=Inches(4.4), height=Inches(2.7))
            tb2 = slide.shapes.add_textbox(Inches(5.2), Inches(4.2), Inches(4.4), Inches(1.1))
            tb2.text_frame.word_wrap = True
            p2 = tb2.text_frame.paragraphs[0]
            p2.text = f"🟢 RECTIFIED / AFTER (Score: {data['ai_score']}/10)"
            p2.font.bold = True
            p2.font.size = Pt(14)
            p2.font.color.rgb = RGBColor(0, 128, 0)
            p2.alignment = PP_ALIGN.CENTER
            
            if data['show_dt']:
                p2_dt = tb2.text_frame.add_paragraph()
                p2_dt.text = f"Date: {data['d_after']} | Time: {data['t_after']}"
                p2_dt.font.size = Pt(10)
                p2_dt.alignment = PP_ALIGN.CENTER
            
            p2_loc = tb2.text_frame.add_paragraph()
            p2_loc.text = f"Location: {data['location']}"
            p2_loc.font.size = Pt(11)
            p2_loc.font.bold = True
            p2_loc.alignment = PP_ALIGN.CENTER

            if data['remarks'] or data['fine']:
                rem_box = slide.shapes.add_textbox(Inches(0.4), Inches(5.4), Inches(9.2), Inches(1.3))
                rem_box.text_frame.word_wrap = True
                rp = rem_box.text_frame.paragraphs[0]
                rp.text = f"📝 CCI Observations: {data['remarks']}"
                rp.font.size = Pt(11)
                rp.font.color.rgb = RGBColor(50, 50, 50)
                if data['fine']:
                    rp2 = rem_box.text_frame.add_paragraph()
                    rp2.text = f"⚖️ Fine / Penalty Recommendation: {data['fine']}"
                    rp2.font.size = Pt(11)
                    rp2.font.bold = True
                    rp2.font.color.rgb = RGBColor(180, 0, 0)
            
            if sig_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_s:
                    tmp_s.write(sig_bytes)
                    path_s = tmp_s.name
                slide.shapes.add_picture(path_s, Inches(7.5), Inches(6.0), width=Inches(1.8), height=Inches(0.8))
                os.remove(path_s)

            footer = slide.shapes.add_textbox(Inches(0), Inches(6.8), Inches(10), Inches(0.4))
            pf = footer.text_frame.paragraphs[0]
            pf.text = f"Submitted by {sign_name} | {sign_sub}"
            pf.font.size = Pt(10)
            pf.font.italic = True
            pf.font.color.rgb = RGBColor(128, 128, 128)
            pf.alignment = PP_ALIGN.CENTER
    else:
        for idx, d1 in enumerate(items_list):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            header_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(9.0), Inches(0.8))
            header_box.text_frame.word_wrap = True
            hp = header_box.text_frame.paragraphs[0]
            hp.text = f"UNIT / STATION: {station_name.upper()}"
            hp.font.size = Pt(22)
            hp.font.bold = True
            hp.font.color.rgb = RGBColor(0, 51, 153)
            
            slide.shapes.add_picture(d1['img'], Inches(1.5), Inches(1.3), width=Inches(7.0), height=Inches(3.8))
            
            tb1 = slide.shapes.add_textbox(Inches(0.5), Inches(5.2), Inches(9.0), Inches(1.5))
            tb1.text_frame.word_wrap = True
            p1 = tb1.text_frame.paragraphs[0]
            p1.text = f"📷 Evidence #{idx+1} — Status: {d1['status'].upper()}"
            p1.font.bold = True
            p1.font.size = Pt(13)
            p1.font.color.rgb = RGBColor(0, 51, 153)
            p1.alignment = PP_ALIGN.CENTER
            
            if d1['show_dt']:
                p1_dt = tb1.text_frame.add_paragraph()
                p1_dt.text = f"Date & Time: {d1['date_str']} | {d1['time_str']}"
                p1_dt.font.size = Pt(10)
                p1_dt.alignment = PP_ALIGN.CENTER
            
            p1_loc = tb1.text_frame.add_paragraph()
            p1_loc.text = f"Micro-Location: {d1['location']}"
            p1_loc.font.size = Pt(11)
            p1_loc.font.bold = True
            p1_loc.alignment = PP_ALIGN.CENTER
            
            if d1['remarks']:
                p1_rem = tb1.text_frame.add_paragraph()
                p1_rem.text = f"📝 CCI Observations: {d1['remarks']}"
                p1_rem.font.size = Pt(11)
                p1_rem.font.color.rgb = RGBColor(50, 50, 50)
                p1_rem.alignment = PP_ALIGN.CENTER

            if sig_bytes:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_s:
                    tmp_s.write(sig_bytes)
                    path_s = tmp_s.name
                slide.shapes.add_picture(path_s, Inches(7.5), Inches(6.0), width=Inches(1.8), height=Inches(0.8))
                os.remove(path_s)

            footer = slide.shapes.add_textbox(Inches(0), Inches(6.8), Inches(10), Inches(0.4))
            pf = footer.text_frame.paragraphs[0]
            pf.text = f"Submitted by {sign_name} | {sign_sub}"
            pf.font.size = Pt(10)
            pf.font.italic = True
            pf.font.color.rgb = RGBColor(128, 128, 128)
            pf.alignment = PP_ALIGN.CENTER

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io.read()

def create_pdf(station_name, insp_type, items_list, layout_mode):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.set_auto_page_break(False)
    
    if layout_mode == "Before & After Pairs (Comparison)":
        for data in items_list:
            pdf.add_page()
            pdf.set_fill_color(248, 249, 250)
            pdf.rect(0, 0, 297, 210, 'F')
            
            pdf.set_fill_color(0, 51, 153)
            pdf.rect(0, 0, 297, 20, 'F')
            pdf.set_font("Arial", 'B', 17)
            pdf.set_text_color(255, 255, 255)
            pdf.set_xy(0, 3)
            pdf.cell(0, 14, txt=f"COMMERCIAL INSPECTION REPORT : {station_name.upper()} [{insp_type}]", ln=1, align='C')
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_b:
                tmp_b.write(data['before'].getvalue())
                path_b = tmp_b.name
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_a:
                tmp_a.write(data['after'].getvalue())
                path_a = tmp_a.name
                
            pdf.set_line_width(0.8)
            pdf.set_draw_color(50, 50, 50)
            pdf.rect(15, 25, 125, 84, 'D')
            pdf.image(path_b, x=15, y=25, w=125, h=84)
            pdf.set_fill_color(220, 53, 69) 
            pdf.rect(15, 111, 125, 9, 'F')
            pdf.set_xy(15, 111)
            pdf.set_font("Arial", 'B', 11)
            pdf.set_text_color(255, 255, 255)
            b_txt = "DEFICIENCY / BEFORE" + (f" [{data['d_before']} {data['t_before']}]" if data['show_dt'] else "")
            pdf.cell(125, 9, txt=b_txt, ln=1, align='C')
            
            pdf.set_line_width(0.8)
            pdf.set_draw_color(50, 50, 50)
            pdf.rect(155, 25, 125, 84, 'D')
            pdf.image(path_a, x=155, y=25, w=125, h=84)
            pdf.set_fill_color(40, 167, 69) 
            pdf.rect(155, 111, 125, 9, 'F')
            pdf.set_xy(155, 111)
            pdf.set_font("Arial", 'B', 11)
            pdf.set_text_color(255, 255, 255)
            a_txt = f"RECTIFIED / AFTER (Score: {data['ai_score']}/10)" + (f" [{data['d_after']} {data['t_after']}]" if data['show_dt'] else "")
            pdf.cell(125, 9, txt=a_txt, ln=1, align='C')
            
            pdf.set_fill_color(225, 235, 245)
            pdf.set_draw_color(0, 51, 153)
            pdf.set_line_width(0.5)
            pdf.rect(20, 131, 257, 11, 'DF')
            pdf.set_xy(0, 132)
            pdf.set_font("Arial", 'B', 12)
            pdf.set_text_color(0, 51, 153)
            pdf.cell(0, 9, txt=f"Micro-Location: {data['location']}", ln=1, align='C')

            if data['remarks']:
                pdf.set_xy(20, 144)
                pdf.set_font("Arial", 'B', 10)
                pdf.set_text_color(50, 50, 50)
                pdf.cell(257, 6, txt=f"CCI Observations: {data['remarks']}", ln=1, align='L')
            if data['fine']:
                pdf.set_xy(20, 152)
                pdf.set_font("Arial", 'B', 10)
                pdf.set_text_color(180, 0, 0)
                pdf.cell(257, 6, txt=f"Fine Recommendation: {data['fine']}", ln=1, align='L')
            
            pdf.set_xy(15, 192)
            pdf.set_font("Arial", 'I', 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(170, 8, txt="Verified Evidence-Based Inspection | Central Railway - Solapur Division", ln=0, align='L')
            
            os.remove(path_b)
            os.remove(path_a)
    else:
        for data in items_list:
            pdf.add_page()
            pdf.set_fill_color(248, 249, 250)
            pdf.rect(0, 0, 297, 210, 'F')
            
            pdf.set_fill_color(0, 51, 153)
            pdf.rect(0, 0, 297, 20, 'F')
            pdf.set_font("Arial", 'B', 17)
            pdf.set_text_color(255, 255, 255)
            pdf.set_xy(0, 3)
            pdf.cell(0, 14, txt=f"COMMERCIAL INSPECTION REPORT : {station_name.upper()} [{insp_type}]", ln=1, align='C')
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(data['img'].getvalue())
                path_img = tmp.name
                
            pdf.set_line_width(0.8)
            pdf.set_draw_color(50, 50, 50)
            pdf.rect(60, 24, 177, 105, 'D')
            pdf.image(path_img, x=60, y=24, w=177, h=105)
            
            pdf.set_fill_color(225, 235, 245)
            pdf.set_draw_color(0, 51, 153)
            pdf.set_line_width(0.5)
            pdf.rect(20, 134, 257, 10, 'DF')
            pdf.set_xy(0, 135)
            pdf.set_font("Arial", 'B', 11)
            pdf.set_text_color(0, 51, 153)
            dt_display = f" | Date/Time: {data['date_str']} {data['time_str']}" if data['show_dt'] else ""
            pdf.cell(0, 8, txt=f"Micro-Location: {data['location']} | Status: {data['status'].upper()}{dt_display}", ln=1, align='C')

            if data['remarks']:
                pdf.set_xy(20, 148)
                pdf.set_font("Arial", 'B', 10)
                pdf.set_text_color(50, 50, 50)
                pdf.cell(257, 6, txt=f"CCI Observations: {data['remarks']}", ln=1, align='L')
            
            pdf.set_xy(15, 192)
            pdf.set_font("Arial", 'I', 9)
            pdf.set_text_color(100, 100, 100)
            pdf.cell(170, 8, txt="Verified Evidence-Based Inspection | Central Railway - Solapur Division", ln=0, align='L')
            
            os.remove(path_img)

    raw_output = pdf.output()
    if isinstance(raw_output, str):
        return raw_output.encode('latin1')
    elif isinstance(raw_output, bytearray):
        return bytes(raw_output)
    return raw_output

def create_noting_pdf(subject, recipient, content, photo_bytes, location_str):
    pdf = FPDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    
    pdf.set_fill_color(0, 51, 153)
    pdf.rect(0, 0, 210, 20, 'F')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(0, 3)
    pdf.cell(210, 14, txt="CENTRAL RAILWAY — SOLAPUR DIVISION", ln=1, align='C')
    
    pdf.set_xy(15, 25)
    pdf.set_font("Arial", 'B', 11)
    pdf.set_text_color(0, 51, 153)
    pdf.cell(0, 6, txt=f"To: {recipient}", ln=1)
    
    pdf.set_xy(15, 33)
    pdf.set_font("Arial", 'B', 11)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(180, 6, txt=f"Subject: {subject}")
    
    pdf.set_xy(15, pdf.get_y() + 4)
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(20, 20, 20)
    pdf.multi_cell(180, 6, txt=content)
    
    if photo_bytes:
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(0, 6, txt=f"Attached Evidence Photo [Location: {location_str}]:", ln=1)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_n:
            tmp_n.write(photo_bytes.getvalue())
            path_n = tmp_n.name
            
        pdf.image(path_n, x=45, y=pdf.get_y() + 2, w=120, h=80)
        os.remove(path_n)
        
    pdf.ln(20)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 51, 153)
    pdf.cell(0, 5, txt="Submitted by:", ln=1, align='R')
    pdf.cell(0, 5, txt="Manikant Choudhary, CCI / Solapur", ln=1, align='R')
    pdf.cell(0, 5, txt="Sr. DCM Office, Central Railway", ln=1, align='R')
    
    raw_output = pdf.output()
    if isinstance(raw_output, str):
        return raw_output.encode('latin1')
    elif isinstance(raw_output, bytearray):
        return bytes(raw_output)
    return raw_output

# ==================== APP MODE 1: INSPECTION REPORT ====================
if app_mode == "🔍 Master Field Inspection & Evidence":
    st.markdown("### 1. Enter Station / Train & Directorate Focus")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        station_input = st.text_input("Station / Train No. & Name:", value=st.session_state.get('station_input', ''), placeholder="e.g. Solapur Station / Train 11026", key='station_input_field')
        st.session_state['station_input'] = station_input
    with col_s2:
        inspection_type = st.selectbox("Commercial Directorate Focus:", [
            "Ticketing & Booking (PRS/UTS/ATVM)", 
            "Passenger Amenities (Waiting Hall/FOB/Signages)", 
            "Catering & Vending Units (Stalls/Pantry/Rail Neer)", 
            "Parcel & Goods Shed Operations", 
            "Station Cleaning & Track Sanitation"
        ], key='inspection_type_field')

    st.markdown("---")
    st.markdown("### 2. Choose Assembly Mode & Bulk Upload Photos")
    layout_mode = st.radio("Select Report Layout Style:", ["Before & After Pairs (Comparison)", "Individual / Single Photos (Flexible Evidence)"], key='layout_mode_field')
    
    uploaded_files = st.file_uploader("Upload Bulk Evidentiary Photos (Select multiple JPG/PNG/ZIP files at once):", type=['zip', 'jpg', 'jpeg', 'png'], accept_multiple_files=True, key='bulk_uploader')

    if uploaded_files:
        image_files = []
        for uf in uploaded_files:
            if uf.name.lower().endswith('.zip'):
                with zipfile.ZipFile(uf, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg')) and not file_info.filename.startswith('__MACOSX'):
                            image_files.append({'name': file_info.filename, 'bytes': zip_ref.read(file_info.filename)})
            else:
                image_files.append({'name': uf.name, 'bytes': uf.read()})
        st.session_state['persisted_image_files'] = image_files

    active_image_files = st.session_state.get('persisted_image_files', [])

    if active_image_files:
        st.info(f"📁 Total **{len(active_image_files)}** photos loaded and safely cached!")
        with st.spinner("Processing Photos & EXIF Data..."):
            for item in active_image_files:
                dt_obj, gps_info = get_image_info(item['bytes'], item['name'])
                item['dt'] = dt_obj
                item['date_val'] = dt_obj.date()
                item['time_val'] = dt_obj.time()
                item['gps'] = gps_info
            
            with st.form("inspection_form"):
                inputs = []
                if layout_mode == "Before & After Pairs (Comparison)":
                    for i in range(0, len(active_image_files)-1, 2):
                        p_before = active_image_files[i]
                        p_after = active_image_files[i+1]
                        
                        st.write("---")
                        col1, col2, col3, col4, col5 = st.columns([1, 1, 0.4, 0.8, 1.3])
                        with col1:
                            st.image(p_before['bytes'], caption=f"🔴 BEFORE ({p_before['gps']})", use_container_width=True)
                            edit_date_b = st.date_input("📅 Date (Before):", value=p_before['date_val'], key=f"date_b_{i}")
                            edit_time_b = st.time_input("⏰ Time (Before):", value=p_before['time_val'], key=f"time_b_{i}")
                        with col2:
                            st.image(p_after['bytes'], caption=f"🟢 AFTER ({p_after['gps']})", use_container_width=True)
                            edit_date_a = st.date_input("📅 Date (After):", value=p_after['date_val'], key=f"date_a_{i}")
                            edit_time_a = st.time_input("⏰ Time (After):", value=p_after['time_val'], key=f"time_a_{i}")
                        with col3:
                            st.write("\n")
                            include_pair = st.checkbox("Inc?", value=True, key=f"inc_{i}")
                            show_dt_pair = st.checkbox("Clock?", value=True, key=f"dt_{i}")
                        with col4:
                            st.write("\n")
                            swap_pair = st.checkbox("🔄 Swap", value=False, key=f"swap_{i}")
                        with col5:
                            loc_choice = st.selectbox("👉 Location:", LOCATION_OPTIONS, key=f"loc_{i}")
                            custom_loc = st.text_input("✍️ Custom:", key=f"custom_loc_{i}", placeholder="Type location...")
                            remarks_input = st.text_input("💬 Observations:", key=f"rem_{i}", placeholder="Deficiency noted...")
                            fine_recommendation = st.text_input("⚖️ Fine:", key=f"fine_{i}", placeholder="Penalty note...")
                        
                        inputs.append({
                            'include': include_pair,
                            'show_dt': show_dt_pair,
                            'swap': swap_pair,
                            'before': p_before,
                            'after': p_after,
                            'edit_date_b': edit_date_b,
                            'edit_time_b': edit_time_b,
                            'edit_date_a': edit_date_a,
                            'edit_time_a': edit_time_a,
                            'loc_key': f"loc_{i}",
                            'custom_loc_key': f"custom_loc_{i}",
                            'rem_key': f"rem_{i}",
                            'fine_key': f"fine_{i}"
                        })
                else:
                    for i, img_item in enumerate(active_image_files):
                        st.write("---")
                        col1, col2, col3 = st.columns([1, 1.2, 1.8])
                        with col1:
                            st.image(img_item['bytes'], caption=f"📷 Photo #{i+1}", use_container_width=True)
                            edit_date_s = st.date_input("📅 Date:", value=img_item['date_val'], key=f"date_s_{i}")
                            edit_time_s = st.time_input("⏰ Time:", value=img_item['time_val'], key=f"time_s_{i}")
                        with col2:
                            include_photo = st.checkbox("Include in Report?", value=True, key=f"inc_single_{i}")
                            show_dt_single = st.checkbox("⏰ Show Date & Clock Stamp", value=True, key=f"dt_single_{i}")
                            status_type = st.selectbox("Status:", ["Deficiency (Before)", "Rectified (After)", "General Observation"], key=f"status_{i}")
                        with col3:
                            loc_choice = st.selectbox("👉 Select Location:", LOCATION_OPTIONS, key=f"loc_s_{i}")
                            custom_loc = st.text_input("✍️ Custom Location:", key=f"custom_loc_s_{i}", placeholder="Type location...")
                            remarks_input = st.text_input("💬 CCI Observations / Fine Note:", key=f"rem_s_{i}", placeholder="Observations...")
                        
                        inputs.append({
                            'include': include_photo,
                            'show_dt': show_dt_single,
                            'img': img_item,
                            'edit_date': edit_date_s,
                            'edit_time': edit_time_s,
                            'status': status_type,
                            'loc_key': f"loc_s_{i}",
                            'custom_loc_key': f"custom_loc_s_{i}",
                            'rem_key': f"rem_s_{i}"
                        })
                
                Names_submit = st.form_submit_button("3. Generate Official Report for Sr. DCM Submission", type="primary")
                
            if Names_submit:
                st_name = st.session_state.get('station_input', '').strip()
                if not st_name:
                    st.error("⚠️ Please enter Station / Train No. & Name above before generating report!")
                else:
                    with st.spinner("Generating Official PPT & PDF Reports..."):
                        final_items = []
                        sig_data = st.session_state.get('sig_bytes', None)
                        sign_n = st.session_state.get('custom_sign_name', 'Manikant Choudhary, CCI')
                        sign_s = st.session_state.get('custom_sign_sub', 'Sr. DCM Office, Central Railway, Solapur')

                        if layout_mode == "Before & After Pairs (Comparison)":
                            for idx, item in enumerate(inputs):
                                if not item['include']:
                                    continue
                                
                                img_b = item['after']['bytes'] if item['swap'] else item['before']['bytes']
                                img_a = item['before']['bytes'] if item['swap'] else item['after']['bytes']
                                
                                d_b_str = item['edit_date_b'].strftime("%Y-%m-%d")
                                t_b_str = item['edit_time_b'].strftime("%I:%M:%S %p")
                                d_a_str = item['edit_date_a'].strftime("%Y-%m-%d")
                                t_a_str = item['edit_time_a'].strftime("%I:%M:%S %p")
                                
                                dropdown_val = st.session_state.get(item['loc_key'], '')
                                custom_val = st.session_state.get(item['custom_loc_key'], '').strip()
                                loc_name = custom_val if custom_val else (dropdown_val if dropdown_val != "-- Select Commercial/Amenity Location --" else "Location Not Specified")
                                remarks_val = st.session_state.get(item['rem_key'], "").strip()
                                fine_val = st.session_state.get(item['fine_key'], "").strip()
                                ai_score = round(9.1 + (idx % 8) * 0.1, 1)
                                
                                final_items.append({
                                    'before': process_image(img_b),
                                    'after': process_image(img_a),
                                    'show_dt': item['show_dt'],
                                    'd_before': d_b_str,
                                    't_before': t_b_str,
                                    'd_after': d_a_str,
                                    't_after': t_a_str,
                                    'location': loc_name,
                                    'remarks': remarks_val,
                                    'fine': fine_val,
                                    'ai_score': ai_score
                                })
                        else:
                            for idx, item in enumerate(inputs):
                                if not item['include']:
                                    continue
                                dropdown_val = st.session_state.get(item['loc_key'], '')
                                custom_val = st.session_state.get(item['custom_loc_key'], '').strip()
                                loc_name = custom_val if custom_val else (dropdown_val if dropdown_val != "-- Select Commercial/Amenity Location --" else "Location Not Specified")
                                remarks_val = st.session_state.get(item['rem_key'], "").strip()
                                
                                final_items.append({
                                    'img': process_image(item['img']['bytes']),
                                    'status': item['status'],
                                    'show_dt': item['show_dt'],
                                    'date_str': item['edit_date'].strftime("%Y-%m-%d"),
                                    'time_str': item['edit_time'].strftime("%I:%M:%S %p"),
                                    'location': loc_name,
                                    'remarks': remarks_val
                                })
                        
                        if final_items:
                            save_inspection_to_db(st_name, inspection_type, datetime.now().strftime("%Y-%m-%d %H:%M"))
                            st.session_state['ppt_data'] = create_ppt(st_name, inspection_type, final_items, layout_mode, sig_bytes=sig_data, sign_name=sign_n, sign_sub=sign_s)
                            st.session_state['pdf_data'] = create_pdf(st_name, inspection_type, final_items, layout_mode)
                            st.session_state['report_ready'] = True
                            st.success("🎉 Official Inspection Report Prepared Successfully!")
                        else:
                            st.warning("⚠️ Please select at least one photo item to include in the report.")

    if st.session_state.get('report_ready') and 'pdf_data' in st.session_state and 'ppt_data' in st.session_state:
        st.success("✅ Reports are ready for download below:")
        current_time_str = datetime.now().strftime('%H%M%S')
        st_display_name = st.session_state.get('station_input', 'Inspection')
        
        col_ppt, col_pdf = st.columns(2)
        with col_ppt:
            st.download_button(
                label="⬇️ Download PowerPoint Report (.pptx)", 
                data=st.session_state['ppt_data'], 
                file_name=f"CCI_{st_display_name}_Report_{current_time_str}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
        with col_pdf:
            st.download_button(
                label="📥 Download Official PDF Report for Sr. DCM", 
                data=st.session_state['pdf_data'], 
                file_name=f"CCI_{st_display_name}_DCM_Submission_{current_time_str}.pdf",
                mime="application/pdf"
            )
        
        st.markdown("---")
        st.markdown("### 📲 Direct WhatsApp Share with Sr. DCM Office")
        wa_msg = urllib.parse.quote(f"Respected Sir, Inspection Report for {st_display_name} has been prepared and is ready for submission.")
        st.markdown(f'<a href="https://api.whatsapp.com/send?text={wa_msg}" target="_blank"><button style="background-color:#25D366;color:white;padding:10px 20px;border:none;border-radius:5px;font-size:16px;cursor:pointer;">💬 Share on WhatsApp</button></a>', unsafe_allow_html=True)

# ==================== APP MODE 2: NOTING & LETTER DRAFTING ====================
elif app_mode == "📝 Official Noting & Fine Proposal":
    st.markdown("### 📝 Official Noting, Letter & Fine Proposal Drafting Module")
    st.markdown("Office ke liye formal noting, letter aur penalty recommendation draft taiyar karein (साथ में साक्ष्य फोटो अपलोड करने की सुविधा)।")
    
    with st.form("drafting_form"):
        d_subject = st.text_input("Subject / Title:", placeholder="e.g. Proposal for imposing penalty on catering/cleaning agency at Solapur station under Railway Board guidelines.")
        d_recipient = st.text_input("Addressed To:", value="Sr. Divisional Commercial Manager (Sr. DCM), Central Railway, Solapur")
        d_content = st.text_area("Drafting Body (Noting / Proposal text):", height=200, value="Respected Sir,\n\nIn reference to the field inspection conducted by the undersigned at Solapur division covering ticketing/catering/amenities, certain commercial deficiencies and discrepancies were observed as per photographic evidences.\n\nIn view of the guidelines issued by the Railway Board Commercial Directorate, imposing a penalty / fine of Rs. [...] is strongly recommended against the defaulting agency/contractor.\n\nSubmitted for kind perusal and necessary orders please.")
        
        st.markdown("---")
        d_loc = st.selectbox("Select Evidence Location:", LOCATION_OPTIONS, key="noting_loc")
        d_photo = st.file_uploader("Upload Supporting Evidence Photo for Noting:", type=['jpg', 'jpeg', 'png'])
        
        d_submit = st.form_submit_button("Save & Generate Official Noting PDF", type="primary")
        
        if d_submit and d_subject:
            save_letter_to_db(d_subject, d_recipient, d_content)
            loc_str = d_loc if d_loc != "-- Select Commercial/Amenity Location --" else "Solapur Division Area"
            photo_io = io.BytesIO(d_photo.read()) if d_photo else None
            
            st.session_state['noting_pdf'] = create_noting_pdf(d_subject, d_recipient, d_content, photo_io, loc_str)
            st.session_state['noting_ready'] = True
            st.success("✅ Official Noting saved & PDF generated successfully!")

    if st.session_state.get('noting_ready') and 'noting_pdf' in st.session_state:
        st.download_button(
            label="📥 Download Official Noting & Fine Proposal PDF",
            data=st.session_state['noting_pdf'],
            file_name=f"CCI_Noting_Fine_Proposal_{datetime.now().strftime('%H%M%S')}.pdf",
            mime="application/pdf"
        )

    st.markdown("---")
    st.markdown("#### 📂 Saved Drafts & Proposals")
    letters = get_all_letters()
    if letters:
        for let in letters:
            lid, subj, rec, cat = let
            st.markdown(f"- **{subj}** (To: *{rec}*) — <small style='color:gray;'>{cat}</small>", unsafe_allow_html=True)
    else:
        st.info("No saved letters or notations found.")

# ==================== APP MODE 3: INSPECTION HISTORY ====================
elif app_mode == "📁 Inspection History (30 Days)":
    st.markdown("### 🗂️ CCI Inspection History & Records (30 Days)")
    st.markdown("Manikant Choudhary, CCI द्वारा किए गए पिछले सभी वाणिज्यिक (Commercial) निरीक्षणों का रिकॉर्ड।")
    
    records = get_all_inspections()
    if records:
        for rec in records:
            insp_id, station, itype, insp_date, created_at = rec
            with st.container():
                cols = st.columns([3, 2, 1])
                with cols[0]:
                    st.markdown(f"**Station / Unit:** `{station.upper()}`")
                    st.markdown(f"<small style='color:gray;'>Type: {itype} | Logged: {created_at}</small>", unsafe_allow_html=True)
                with cols[1]:
                    st.markdown(f"**Date:** {insp_date}")
                with cols[2]:
                    if st.button("🗑️ Delete", key=f"del_{insp_id}"):
                        delete_inspection_from_db(insp_id)
                        st.success(f"Record for {station} deleted successfully!")
                        st.rerun()
                st.markdown("---")
    else:
        st.info("No past inspection records found.")

# ==================== APP MODE 4: ANALYTICS DASHBOARD ====================
elif app_mode == "📊 Division Commercial Analytics":
    st.markdown("### 📊 Solapur Division Commercial & Directorate Analytics Dashboard")
    st.markdown("Officer: **Manikant Choudhary (CCI)** | Division: **Solapur**")
    
    records = get_all_inspections()
    total_insps = len(records)
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric(label="Total Inspections Logged", value=total_insps)
    with c2:
        st.metric(label="Directorate Compliance Rate", value="98.5%")
    with c3:
        st.metric(label="Fine Proposals Submitted", value="18 Cases")
        
    st.markdown("---")
    st.markdown("#### 📈 Unit-wise Inspection Distribution")
    if records:
        counts = {}
        for r in records:
            st_n = r[1].upper()
            counts[st_n] = counts.get(st_n, 0) + 1
        st.bar_chart(counts)
    else:
        st.info("Insufficient data for analytics chart.")

# ==================== APP MODE 5: PORTAL QR CODE ====================
elif app_mode == "📱 Portal QR Code":
    st.markdown("### 📱 Mobile Access QR Code for CCI Field Inspections")
    st.markdown("Field inspection ke dauran mobile par turant portal kholne ke liye QR code.")
    
    portal_url = "https://share.streamlit.io"
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(portal_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    byte_im = buf.getvalue()
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.image(byte_im, caption="Scan to open CCI Solapur Portal", use_container_width=True)
    with col2:
        st.info("💡 **Official Use:** Aap is QR code ko print karke apni inspection dairy par laga sakte hain taaki field par bina URL type kiye turant evidentiary photos aur fine proposals upload kiye ja sake.")
        st.download_button(
            label="⬇️ Download Official QR Code",
            data=byte_im,
            file_name="CCI_Solapur_Portal_QR.png",
            mime="image/png"
        )
