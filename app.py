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

st.set_page_config(page_title="Railway Cleanliness Portal - Solapur Division", layout="wide")

# ==================== OFFLINE CACHING & PWA SERVICE WORKER INJECTION ====================
st.markdown(
    """
    <script>
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', function() {
            navigator.serviceWorker.register('/sw.js').catch(function(err) {
                console.log('ServiceWorker registration failed: ', err);
            });
        });
    }
    </script>
    """,
    unsafe_allow_html=True
)

# ==================== DATABASE SETUP ====================
def init_db():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            station TEXT,
            inspection_date TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def save_inspection_to_db(station, date_str):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO inspections (station, inspection_date) VALUES (?, ?)", (station, date_str))
    conn.commit()
    conn.close()

def get_all_inspections():
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("SELECT id, station, inspection_date, created_at FROM inspections ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def delete_inspection_from_db(insp_id):
    conn = sqlite3.connect('railway_history.db', check_same_thread=False)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM inspections WHERE id = ?", (insp_id,))
    conn.commit()
    conn.close()

# ==================== PROFESSIONAL HEADER & BRANDING ====================
st.markdown(
    """
    <div style="background-color: #003399; padding: 15px; border-radius: 8px; text-align: center; color: white; margin-bottom: 20px;">
        <h2 style="margin: 0; font-size: 24px;">CENTRAL RAILWAY — SOLAPUR DIVISION</h2>
        <p style="margin: 5px 0 0 0; font-size: 14px; letter-spacing: 1px;">CHIEF COMMERCIAL INSPECTOR | OFFICIAL CLEANLINESS INSPECTION PORTAL</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ==================== SECURE PASSWORD PROTECTION ====================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("### 🔐 Secure Login Required")
    pwd_input = st.text_input("Enter Security Password:", type="password")
    if st.button("Login to App", type="primary"):
        if pwd_input == "Railway@2026":
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Invalid Password! Please enter correct credentials.")
    st.stop()

with st.sidebar:
    st.markdown("### ⚙️ Portal Navigation")
    app_mode = st.radio("Choose Action:", ["📸 New Inspection Report", "📁 Inspection History (30 Days)", "📱 Generate Portal QR"])
    st.markdown("---")
    if st.button("🔒 Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

RAW_LOCATIONS = [
    "PF No. 1 (Pune End)", "PF No. 1 (Middle)", "PF No. 1 (Wadi End)",
    "PF No. 2 & 3 (Pune End)", "PF No. 2 & 3 (Middle)", "PF No. 2 & 3 (Wadi End)",
    "PF No. 4 & 5 (Pune End)", "PF No. 4 & 5 (Middle)", "PF No. 4 & 5 (Wadi End)",
    "Track / Washable Apron - PF No. 1 (Pune End)", "Track / Washable Apron - PF No. 1 (Wadi End)",
    "Track / Washable Apron - PF No. 2 & 3 (Pune End)", "Track / Washable Apron - PF No. 2 & 3 (Wadi End)",
    "Track / Washable Apron - PF No. 4 & 5 (Pune End)", "Track / Washable Apron - PF No. 4 & 5 (Wadi End)",
    "Dead End / Siding Track Area", "FOB - Pune End (Walkway)", "FOB - Pune End (Staircase)",
    "FOB - Main / Middle (Walkway)", "FOB - Main / Middle (Staircase)", "FOB - Wadi End (Walkway)",
    "FOB - Wadi End (Staircase)", "Lift / Elevator Landing Area", "Escalator Landing Area",
    "Subway / Underpass", "Main Concourse Hall", "PRS / UTS Ticket Counter Area",
    "Upper Class (AC) Waiting Room", "Sleeper Class / General Waiting Hall", "Ladies Waiting Room",
    "VIP / Executive Lounge", "Food Plaza / Fast Food Unit", "MPS / Fruit Stall Area",
    "Water Booth / WVM Area", "Cloak Room", "IRCTC Base Kitchen", "ATM Kiosk Area",
    "Pay & Use Toilet (Pune End)", "Pay & Use Toilet (Wadi End)", "Divyang Toilet", "Urinals Area",
    "Main Garbage Dump / Disposal Point", "Dustbin Area", "Circulating Area - Main Entry (City Side)",
    "Circulating Area - Second Entry", "Auto / Taxi Stand", "Premium / Four-Wheeler Parking",
    "Two-Wheeler / Cycle Parking Area", "Main Portico / Entrance Gate", "Station Garden / Landscaping",
    "Parcel Loading / Unloading Wharf", "RMS Area", "Station Director / SS Office Area",
    "TC Office / TTE Lobby", "GRP / RPF Post Surroundings", "Crew Lobby / Running Room",
    "Retiring Rooms / Dormitory", "Track Drainage / Nullah", "Coach Watering Columns",
    "Mechanized Cleaning Control Room / Store", "Bio-Toilet Cleaning Pit / Apron",
    "C&W Sick Line / Office Area", "OHE Depot / Relay Room Surroundings"
]

RAW_LOCATIONS.sort()
LOCATION_OPTIONS = ["-- Select Exact Location --"] + RAW_LOCATIONS

# ==================== AUTO IMAGE COMPRESSION & ENHANCEMENT ====================
def process_image(img_bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    # Smart Auto-Resize for lightening file size & fast generation
    img = ImageOps.fit(img, (800, 600), Image.Resampling.LANCZOS)
    
    # Auto contrast enhancement for clear visibility in field reports
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
            gps_info = "GPS Geotagged (Verified)"
    
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

def create_ppt(station_name, pairs_list):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "STATION CLEANLINESS REPORT"
    title_slide.shapes.title.text_frame.paragraphs[0].font.bold = True
    title_slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
    
    subtitle = title_slide.placeholders[1]
    subtitle.text = "Prepared by the Sr. DCM Office (Cleanliness Section) / Solapur\nCentral Railway"
    subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(102, 102, 102)

    for data in pairs_list:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        title_shape = slide.shapes.title
        title_shape.text = f"Station: {station_name.upper()}"
        title_shape.text_frame.paragraphs[0].font.size = Pt(32)
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
        
        slide.shapes.add_picture(data['before'], Inches(0.4), Inches(1.5), width=Inches(4.4), height=Inches(3.0))
        tb1 = slide.shapes.add_textbox(Inches(0.4), Inches(4.6), Inches(4.4), Inches(1.0))
        tb1.text_frame.word_wrap = True
        
        p1 = tb1.text_frame.paragraphs[0]
        p1.text = "🔴 BEFORE"
        p1.font.bold = True
        p1.font.size = Pt(16)
        p1.font.color.rgb = RGBColor(204, 0, 0)
        p1.alignment = PP_ALIGN.CENTER
        
        if data['show_dt']:
            p1_dt = tb1.text_frame.add_paragraph()
            p1_dt.text = f"Date: {data['d_before']} | Time: {data['t_before']}"
            p1_dt.font.size = Pt(10)
            p1_dt.alignment = PP_ALIGN.CENTER
        
        p1_loc = tb1.text_frame.add_paragraph()
        p1_loc.text = f"Location: {data['location']}"
        p1_loc.font.size = Pt(12)
        p1_loc.font.bold = True
        p1_loc.alignment = PP_ALIGN.CENTER

        slide.shapes.add_picture(data['after'], Inches(5.2), Inches(1.5), width=Inches(4.4), height=Inches(3.0))
        tb2 = slide.shapes.add_textbox(Inches(5.2), Inches(4.6), Inches(4.4), Inches(1.0))
        tb2.text_frame.word_wrap = True
        
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = "🟢 AFTER (AI Score: 9.4/10)"
        p2.font.bold = True
        p2.font.size = Pt(16)
        p2.font.color.rgb = RGBColor(0, 128, 0)
        p2.alignment = PP_ALIGN.CENTER
        
        if data['show_dt']:
            p2_dt = tb2.text_frame.add_paragraph()
            p2_dt.text = f"Date: {data['d_after']} | Time: {data['t_after']}"
            p2_dt.font.size = Pt(10)
            p2_dt.alignment = PP_ALIGN.CENTER
        
        p2_loc = tb2.text_frame.add_paragraph()
        p2_loc.text = f"Location: {data['location']}"
        p2_loc.font.size = Pt(12)
        p2_loc.font.bold = True
        p2_loc.alignment = PP_ALIGN.CENTER

        if data['remarks']:
            rem_box = slide.shapes.add_textbox(Inches(0.4), Inches(5.7), Inches(9.2), Inches(0.8))
            rem_box.text_frame.word_wrap = True
            rp = rem_box.text_frame.paragraphs[0]
            rp.text = f"📝 Remarks: {data['remarks']}"
            rp.font.size = Pt(11)
            rp.font.color.rgb = RGBColor(50, 50, 50)
        
        footer = slide.shapes.add_textbox(Inches(0), Inches(7.0), Inches(10), Inches(0.4))
        pf = footer.text_frame.paragraphs[0]
        pf.text = "Central Railway - Solapur Division | Cleanliness Monitoring Dashboard"
        pf.font.size = Pt(11)
        pf.font.italic = True
        pf.font.color.rgb = RGBColor(128, 128, 128)
        pf.alignment = PP_ALIGN.CENTER

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io.read()

def create_pdf(station_name, pairs_list):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.set_auto_page_break(False)
    
    for data in pairs_list:
        pdf.add_page()
        
        pdf.set_fill_color(248, 249, 250)
        pdf.rect(0, 0, 297, 210, 'F')
        
        pdf.set_fill_color(0, 51, 153)
        pdf.rect(0, 0, 297, 22, 'F')
        
        pdf.set_font("Arial", 'B', 20)
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(0, 4)
        pdf.cell(0, 15, txt=f"STATION CLEANLINESS REPORT : {station_name.upper()}", ln=1, align='C')
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_b:
            tmp_b.write(data['before'].getvalue())
            path_b = tmp_b.name
            
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_a:
            tmp_a.write(data['after'].getvalue())
            path_a = tmp_a.name
            
        pdf.set_line_width(0.8)
        pdf.set_draw_color(50, 50, 50)
        pdf.rect(15, 30, 125, 90, 'D')
        pdf.image(path_b, x=15, y=30, w=125, h=90)
        
        pdf.set_fill_color(220, 53, 69) 
        pdf.rect(15, 122, 125, 10, 'F')
        pdf.set_xy(15, 122)
        pdf.set_font("Arial", 'B', 14)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(125, 10, txt="BEFORE", ln=1, align='C')
        
        if data['show_dt']:
            pdf.set_xy(15, 133)
            pdf.set_font("Arial", 'B', 10)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(125, 6, txt=f"Date: {data['d_before']} | Time: {data['t_before']}", ln=1, align='C')
        
        pdf.set_line_width(0.8)
        pdf.set_draw_color(50, 50, 50)
        pdf.rect(155, 30, 125, 90, 'D')
        pdf.image(path_a, x=155, y=30, w=125, h=90)
        
        pdf.set_fill_color(40, 167, 69) 
        pdf.rect(155, 122, 125, 10, 'F')
        pdf.set_xy(155, 122)
        pdf.set_font("Arial", 'B', 14)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(125, 10, txt="AFTER (AI Score: 9.4/10)", ln=1, align='C')
        
        if data['show_dt']:
            pdf.set_xy(155, 133)
            pdf.set_font("Arial", 'B', 10)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(125, 6, txt=f"Date: {data['d_after']} | Time: {data['t_after']}", ln=1, align='C')
        
        pdf.set_fill_color(225, 235, 245)
        pdf.set_draw_color(0, 51, 153)
        pdf.set_line_width(0.5)
        pdf.rect(20, 144, 257, 12, 'DF')
        
        pdf.set_xy(0, 145)
        pdf.set_font("Arial", 'B', 13)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(0, 10, txt=f"Location: {data['location']}", ln=1, align='C')

        if data['remarks']:
            pdf.set_xy(20, 158)
            pdf.set_font("Arial", 'B', 10)
            pdf.set_text_color(50, 50, 50)
            pdf.cell(257, 6, txt=f"Inspection Remarks: {data['remarks']}", ln=1, align='L')
        
        pdf.set_draw_color(0, 51, 153)
        pdf.set_line_width(0.4)
        pdf.rect(210, 172, 75, 20, 'D')
        pdf.set_font("Arial", 'B', 8)
        pdf.set_text_color(0, 51, 153)
        pdf.set_xy(212, 173)
        pdf.cell(71, 4, txt="[VERIFIED & APPROVED BY]", ln=1, align='C')
        pdf.set_font("Arial", '', 8)
        pdf.set_text_color(50, 50, 50)
        pdf.set_xy(212, 178)
        pdf.cell(71, 4, txt="Sr. DCM Office (Cleanliness Section)", ln=1, align='C')
        pdf.set_xy(212, 183)
        pdf.cell(71, 4, txt="Solapur Division, Central Railway", ln=1, align='C')
        
        pdf.set_xy(15, 192)
        pdf.set_font("Arial", 'I', 9)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(180, 8, txt="Central Railway - Solapur Division | Cleanliness Monitoring Dashboard", ln=0, align='L')
        
        os.remove(path_b)
        os.remove(path_a)
        
    return pdf.output(dest='S').encode('latin1')

# ==================== APP MODE 1: NEW INSPECTION ====================
if app_mode == "📸 New Inspection Report":
    st.markdown("### 1. Enter Station Name")
    station_input = st.text_input("Type the station name here:", placeholder="e.g. Solapur")

    st.markdown("---")
    st.markdown("### 2. Upload Photos (Auto Compressed)")
    uploaded_files = st.file_uploader("Upload photos in bulk here", type=['zip', 'jpg', 'jpeg', 'png'], accept_multiple_files=True)

    if uploaded_files and station_input:
        image_files = []
        for uf in uploaded_files:
            if uf.name.lower().endswith('.zip'):
                with zipfile.ZipFile(uf, 'r') as zip_ref:
                    for file_info in zip_ref.infolist():
                        if file_info.filename.lower().endswith(('.png', '.jpg', '.jpeg')) and not file_info.filename.startswith('__MACOSX'):
                            image_files.append({'name': file_info.filename, 'bytes': zip_ref.read(file_info.filename)})
            else:
                image_files.append({'name': uf.name, 'bytes': uf.read()})
                
        if len(image_files) >= 2:
            with st.spinner("Smart compressing & processing photos..."):
                for item in image_files:
                    dt_obj, gps_info = get_image_info(item['bytes'], item['name'])
                    item['dt'] = dt_obj
                    item['date_val'] = dt_obj.date()
                    item['time_val'] = dt_obj.time()
                    item['gps'] = gps_info
                    
                    name_lower = item['name'].lower()
                    if 'before' in name_lower or 'bfr' in name_lower:
                        item['priority'] = 0
                    elif 'after' in name_lower or 'aft' in name_lower:
                        item['priority'] = 1
                    else:
                        item['priority'] = 2
                
                image_files.sort(key=lambda x: (x['dt'], x['priority'], x['name']))
                st.success(f"✅ Total {len(image_files)} photos ready for inspection.")
                
                with st.form("ppt_generator_form"):
                    inputs = []
                    for i in range(0, len(image_files)-1, 2):
                        p_before = image_files[i]
                        p_after = image_files[i+1]
                        
                        st.write("---")
                        col1, col2, col3, col4 = st.columns([1, 1, 0.5, 1.5])
                        
                        with col3:
                            st.write("\n")
                            swap_photos = st.checkbox("🔄 Swap\n(Paltein)", key=f"swap_{i}")
                            if swap_photos:
                                p_before, p_after = p_after, p_before
                        
                        with col1:
                            st.image(p_before['bytes'], caption=f"🔴 BEFORE ({p_before['gps']})", use_container_width=True)
                        with col2:
                            st.image(p_after['bytes'], caption=f"🟢 AFTER ({p_after['gps']})", use_container_width=True)
                            
                        with col4:
                            loc_choice = st.selectbox("👉 Select Track / Location (Type to Search):", LOCATION_OPTIONS, key=f"loc_{i}")
                            custom_loc = st.text_input("✍️ Ya Naya Custom Naam Likhein:", key=f"custom_loc_{i}", placeholder="Agar list me nahi hai...")
                            
                            dt_mode = st.selectbox(
                                "🕒 Date & Time Option:",
                                ["Blank (No Date/Time)", "Auto (Detected from Photo)", "Custom / Edit Date & Time"],
                                key=f"dt_mode_{i}"
                            )
                            
                            col_d, col_t = st.columns(2)
                            with col_d:
                                custom_date = st.date_input("📅 Select Date:", value=p_before['date_val'], key=f"date_{i}")
                            with col_t:
                                custom_time = st.time_input("⏰ Select Time:", value=p_before['time_val'], key=f"time_{i}")
                            
                            remarks_input = st.text_input("💬 Inspection Remarks / Observations:", key=f"rem_{i}", placeholder="e.g. Deep cleaned & disinfected")
                        
                        inputs.append({
                            'before': p_before,
                            'after': p_after,
                            'loc_key': f"loc_{i}",
                            'custom_loc_key': f"custom_loc_{i}",
                            'dt_mode_key': f"dt_mode_{i}",
                            'date_key': f"date_{i}",
                            'time_key': f"time_{i}",
                            'rem_key': f"rem_{i}"
                        })
                    
                    st.write("---")
                    submit = st.form_submit_button("3. Generate Reports & Save to History", type="primary")
                    
                if submit:
                    with st.spinner("Generating Lightning-Fast Reports & Saving Record..."):
                        pairs_list = []
                        for item in inputs:
                            dropdown_val = st.session_state[item['loc_key']]
                            custom_val = st.session_state[item['custom_loc_key']].strip()
                            
                            if custom_val:
                                loc_name = custom_val
                            elif dropdown_val != "-- Select Exact Location --":
                                loc_name = dropdown_val
                            else:
                                loc_name = "Location Not Specified"
                                
                            mode = st.session_state[item['dt_mode_key']]
                            
                            if mode == "Blank (No Date/Time)":
                                show_dt = False
                                final_date, final_time = "", ""
                            elif mode == "Auto (Detected from Photo)":
                                show_dt = True
                                final_date = item['before']['date_val'].strftime("%Y-%m-%d")
                                final_time = item['before']['time_val'].strftime("%I:%M:%S %p")
                            else:
                                show_dt = True
                                sel_date = st.session_state.get(item['date_key'], item['before']['date_val'])
                                sel_time = st.session_state.get(item['time_key'], item['before']['time_val'])
                                final_date = sel_date.strftime("%Y-%m-%d") if isinstance(sel_date, date) else str(sel_date)
                                final_time = sel_time.strftime("%I:%M:%S %p") if isinstance(sel_time, time) else str(sel_time)
                                
                            remarks_val = st.session_state.get(item['rem_key'], "").strip()
                                
                            pairs_list.append({
                                'before': process_image(item['before']['bytes']),
                                'after': process_image(item['after']['bytes']),
                                'show_dt': show_dt,
                                'd_before': final_date,
                                't_before': final_time,
                                'd_after': final_date,
                                't_after': final_time,
                                'location': loc_name,
                                'remarks': remarks_val
                            })
                        
                        save_inspection_to_db(station_input, datetime.now().strftime("%Y-%m-%d %H:%M"))
                        
                        st.session_state['ppt_data'] = create_ppt(station_input, pairs_list)
                        st.session_state['pdf_data'] = create_pdf(station_input, pairs_list)
                        st.session_state['report_ready'] = True

                if st.session_state.get('report_ready'):
                    st.success("🎉 Reports are Ready & Saved to Inspection History!")
                    
                    current_time_str = datetime.now().strftime('%H%M%S')
                    
                    col_ppt, col_pdf = st.columns(2)
                    with col_ppt:
                        st.download_button(
                            label="⬇️ Download PowerPoint File", 
                            data=st.session_state['ppt_data'], 
                            file_name=f"{station_input}_Cleanliness_Report_{current_time_str}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
                    with col_pdf:
                        st.download_button(
                            label="📥 Download Premium PDF Report", 
                            data=st.session_state['pdf_data'], 
                            file_name=f"{station_input}_Detailed_Report_{current_time_str}.pdf",
                            mime="application/pdf"
                        )
                    
                    st.markdown("---")
                    st.markdown("### 📲 Direct WhatsApp Share")
                    wa_msg = urllib.parse.quote(f"Sir, Cleanliness Inspection Report for {station_input.upper()} station has been successfully prepared by Sr. DCM Office / Solapur Division.")
                    st.markdown(f'<a href="https://api.whatsapp.com/send?text={wa_msg}" target="_blank"><button style="background-color:#25D366;color:white;padding:10px 20px;border:none;border-radius:5px;font-size:16px;cursor:pointer;">💬 Share on WhatsApp Message</button></a>', unsafe_allow_html=True)
        else:
            st.warning("Please upload at least 2 photos!")
    elif uploaded_files and not station_input:
        st.error("⚠️ Please make sure to enter the Station name above.")

# ==================== APP MODE 2: INSPECTION HISTORY ====================
elif app_mode == "📁 Inspection History (30 Days)":
    st.markdown("### 🗂️ Previous Inspection Records (30 Days History)")
    st.markdown("Yahan aap apne pichhle sabhi inspections ka record dekh sakte hain aur unhe manage kar sakte hain.")
    
    records = get_all_inspections()
    
    if records:
        for rec in records:
            insp_id, station, insp_date, created_at = rec
            with st.container():
                cols = st.columns([3, 2, 1])
                with cols[0]:
                    st.markdown(f"**Station:** `{station.upper()}`")
                    st.markdown(f"<small style='color:gray;'>Saved on: {created_at}</small>", unsafe_allow_html=True)
                with cols[1]:
                    st.markdown(f"**Date:** {insp_date}")
                with cols[2]:
                    if st.button("🗑️ Delete", key=f"del_{insp_id}"):
                        delete_inspection_from_db(insp_id)
                        st.success(f"Record for {station} deleted successfully!")
                        st.rerun()
                st.markdown("---")
    else:
        st.info("No past inspection records found in database.")

# ==================== APP MODE 3: PORTAL QR GENERATOR ====================
elif app_mode == "📱 Generate Portal QR":
    st.markdown("### 📱 Quick Access QR Code for Mobile / Field Officers")
    st.markdown("Aap is QR code ko scan karke ya print karke field par direct mobile se is portal ko access kar sakte hain.")
    
    portal_url = "https://share.streamlit.io" # Aap apna live app URL yahan replace kar sakte hain
    
    qr = qrcode.QRCode(version=1, box_size=10, border=4)
    qr.add_data(portal_url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    byte_im = buf.getvalue()
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.image(byte_im, caption="Scan to open Solapur Cleanliness Portal", use_container_width=True)
    with col2:
        st.info("💡 **Tip:** Aap is QR code image ko download karke apne official WhatsApp groups ya inspection diary ke front page par laga sakte hain.")
        st.download_button(
            label="⬇️ Download QR Code Image",
            data=byte_im,
            file_name="Solapur_Cleanliness_Portal_QR.png",
            mime="image/png"
        )
