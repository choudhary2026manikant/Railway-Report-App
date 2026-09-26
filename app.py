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
        "📈 Monthly Dossier & Performance",
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

FINE_PRESETS = [
    "-- Select Railway Board Fine / Penalty Rule --",
    "Catering Hygiene Violation (RB Circular No. 12/2022) - ₹10,000/-",
    "Unauthorised Vending / Hawking (Sec 144/147) - ₹5,000/-",
    "Platform Cleanliness Default (Swachh Rail Policy) - ₹25,000/-",
    "Ticketless Travel / Irregular Ticketing Counter - ₹2,000/-",
    "Parcel Overloading / Wharfage Violation - ₹10,000/-"
]

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
    gps_info = "GPS Geotagged (Verified Evidence)"
    
    exif = get_exif_data(img_bytes)
    if exif:
        if 'DateTimeOriginal' in exif or 'DateTime' in exif:
            val = exif.get('DateTimeOriginal', exif.get('DateTime'))
            try:
                dt_obj = datetime.strptime(str(val).strip(), '%Y:%m:%d %H:%M:%S')
            except:
                pass
        if 'GPSInfo' not in exif:
            gps_info = "GPS: Manual / Standard Capture"
    
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
                    rp2.text = f"⚖️ Fine / Penalty Rule: {data['fine']}"
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
                pdf.cell(257, 6, txt=f"Fine Rule / Penalty: {data['fine']}", ln=1, align='L')
            
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
        return raw
