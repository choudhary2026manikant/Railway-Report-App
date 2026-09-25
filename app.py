import streamlit as st
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image, ImageEnhance, ImageOps, ExifTags
import io, zipfile, re, os, tempfile
from datetime import datetime

st.set_page_config(page_title="Railway Cleanliness Report", layout="wide")

RAW_LOCATIONS = [
    "PF No. 1 (Pune End)",
    "PF No. 1 (Middle)",
    "PF No. 1 (Wadi End)",
    "PF No. 2 & 3 (Pune End)",
    "PF No. 2 & 3 (Middle)",
    "PF No. 2 & 3 (Wadi End)",
    "PF No. 4 & 5 (Pune End)",
    "PF No. 4 & 5 (Middle)",
    "PF No. 4 & 5 (Wadi End)",
    "Track / Washable Apron - PF No. 1 (Pune End)",
    "Track / Washable Apron - PF No. 1 (Wadi End)",
    "Track / Washable Apron - PF No. 2 & 3 (Pune End)",
    "Track / Washable Apron - PF No. 2 & 3 (Wadi End)",
    "Track / Washable Apron - PF No. 4 & 5 (Pune End)",
    "Track / Washable Apron - PF No. 4 & 5 (Wadi End)",
    "Dead End / Siding Track Area",
    "FOB - Pune End (Walkway)",
    "FOB - Pune End (Staircase)",
    "FOB - Main / Middle (Walkway)",
    "FOB - Main / Middle (Staircase)",
    "FOB - Wadi End (Walkway)",
    "FOB - Wadi End (Staircase)",
    "Lift / Elevator Landing Area",
    "Escalator Landing Area",
    "Subway / Underpass",
    "Main Concourse Hall",
    "PRS / UTS Ticket Counter Area",
    "Upper Class (AC) Waiting Room",
    "Sleeper Class / General Waiting Hall",
    "Ladies Waiting Room",
    "VIP / Executive Lounge",
    "Food Plaza / Fast Food Unit",
    "MPS / Fruit Stall Area",
    "Water Booth / WVM Area",
    "Cloak Room",
    "IRCTC Base Kitchen",
    "ATM Kiosk Area",
    "Pay & Use Toilet (Pune End)",
    "Pay & Use Toilet (Wadi End)",
    "Divyang Toilet",
    "Urinals Area",
    "Main Garbage Dump / Disposal Point",
    "Dustbin Area",
    "Circulating Area - Main Entry (City Side)",
    "Circulating Area - Second Entry",
    "Auto / Taxi Stand",
    "Premium / Four-Wheeler Parking",
    "Two-Wheeler / Cycle Parking Area",
    "Main Portico / Entrance Gate",
    "Station Garden / Landscaping",
    "Parcel Loading / Unloading Wharf",
    "RMS Area",
    "Station Director / SS Office Area",
    "TC Office / TTE Lobby",
    "GRP / RPF Post Surroundings",
    "Crew Lobby / Running Room",
    "Retiring Rooms / Dormitory",
    "Track Drainage / Nullah",
    "Coach Watering Columns",
    "Mechanized Cleaning Control Room / Store",
    "Bio-Toilet Cleaning Pit / Apron",
    "C&W Sick Line / Office Area",
    "OHE Depot / Relay Room Surroundings"
]

# List ko A to Z sort kiya gaya
RAW_LOCATIONS.sort()

# Sort karne ke baad sabse upar Default option jod diya gaya
LOCATION_OPTIONS = ["-- Select Exact Location --"] + RAW_LOCATIONS

def process_image(img_bytes):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    img = ImageEnhance.Color(img).enhance(1.15)
    img = ImageEnhance.Sharpness(img).enhance(1.2)
    img = ImageOps.fit(img, (800, 600), Image.Resampling.LANCZOS)
    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format='JPEG', quality=95)
    img_byte_arr.seek(0)
    return img_byte_arr

def get_exif_datetime(img_bytes):
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if hasattr(img, '_getexif') and img._getexif() is not None:
            exif_data = img._getexif()
            for tag_id, val in exif_data.items():
                tag = ExifTags.TAGS.get(tag_id, tag_id)
                if tag == 'DateTimeOriginal' or tag == 'DateTime':
                    dt = datetime.strptime(str(val).strip(), '%Y:%m:%d %H:%M:%S')
                    return dt
    except Exception as e:
        pass
    return None

def extract_whatsapp_datetime(filename):
    match = re.search(r"(\d{4}-\d{2}-\d{2}) at (\d{1,2}\.\d{2}\.\d{2}\s?[AM|PM|am|pm]+)", filename)
    if match:
        date_str = match.group(1)
        time_str = match.group(2).replace('.', ':').upper()
        try:
            dt_obj = datetime.strptime(f"{date_str} {time_str}", "%Y-%m-%d %I:%M:%S %p")
            return dt_obj, date_str, time_str
        except:
            return None, date_str, time_str
    return None, None, None

def get_image_info(img_bytes, filename):
    date_str = "Date Not Found"
    time_str = "Time Not Found"
    dt_obj = datetime.max 
    
    exif_dt = get_exif_datetime(img_bytes)
    if exif_dt:
        dt_obj = exif_dt
        date_str = exif_dt.strftime("%Y-%m-%d")
        time_str = exif_dt.strftime("%I:%M:%S %p")
        return dt_obj, date_str, time_str
    
    wa_dt_obj, wa_date, wa_time = extract_whatsapp_datetime(filename)
    if wa_dt_obj:
        dt_obj = wa_dt_obj
        date_str = wa_date
        time_str = wa_time
    
    return dt_obj, date_str, time_str

def create_ppt(station_name, pairs_list):
    prs = Presentation()
    
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "STATION CLEANLINESS REPORT"
    title_slide.shapes.title.text_frame.paragraphs[0].font.bold = True
    title_slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
    
    subtitle = title_slide.placeholders[1]
    subtitle.text = "Solapur Division, Central Railway\nChief Commercial Inspector"
    subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(102, 102, 102)

    for data in pairs_list:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        
        title_shape = slide.shapes.title
        title_shape.text = f"Station: {station_name.upper()}"
        title_shape.text_frame.paragraphs[0].font.size = Pt(36)
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 51, 153)
        
        slide.shapes.add_picture(data['before'], Inches(0.4), Inches(1.8), width=Inches(4.4), height=Inches(3.3))
        tb1 = slide.shapes.add_textbox(Inches(0.4), Inches(5.2), Inches(4.4), Inches(1.2))
        tb1.text_frame.word_wrap = True
        
        p1 = tb1.text_frame.paragraphs[0]
        p1.text = "🔴 BEFORE"
        p1.font.bold = True
        p1.font.size = Pt(18)
        p1.font.color.rgb = RGBColor(204, 0, 0)
        p1.alignment = PP_ALIGN.CENTER
        
        p1_dt = tb1.text_frame.add_paragraph()
        p1_dt.text = f"Date: {data['d_before']} | Time: {data['t_before']}"
        p1_dt.font.size = Pt(12)
        p1_dt.alignment = PP_ALIGN.CENTER
        
        p1_loc = tb1.text_frame.add_paragraph()
        p1_loc.text = f"Location: {data['location']}"
        p1_loc.font.size = Pt(13)
        p1_loc.font.bold = True
        p1_loc.alignment = PP_ALIGN.CENTER

        slide.shapes.add_picture(data['after'], Inches(5.2), Inches(1.8), width=Inches(4.4), height=Inches(3.3))
        tb2 = slide.shapes.add_textbox(Inches(5.2), Inches(5.2), Inches(4.4), Inches(1.2))
        tb2.text_frame.word_wrap = True
        
        p2 = tb2.text_frame.paragraphs[0]
        p2.text = "🟢 AFTER"
        p2.font.bold = True
        p2.font.size = Pt(18)
        p2.font.color.rgb = RGBColor(0, 128, 0)
        p2.alignment = PP_ALIGN.CENTER
        
        p2_dt = tb2.text_frame.add_paragraph()
        p2_dt.text = f"Date: {data['d_after']} | Time: {data['t_after']}"
        p2_dt.font.size = Pt(12)
        p2_dt.alignment = PP_ALIGN.CENTER
        
        p2_loc = tb2.text_frame.add_paragraph()
        p2_loc.text = f"Location: {data['location']}"
        p2_loc.font.size = Pt(13)
        p2_loc.font.bold = True
        p2_loc.alignment = PP_ALIGN.CENTER
        
        footer = slide.shapes.add_textbox(Inches(0), Inches(7.0), Inches(10), Inches(0.4))
        pf = footer.text_frame.paragraphs[0]
        pf.text = "Central Railway - Solapur Division | Cleanliness Monitoring Dashboard"
        pf.font.size = Pt(12)
        pf.font.italic = True
        pf.font.color.rgb = RGBColor(128, 128, 128)
        pf.alignment = PP_ALIGN.CENTER

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# EKDAM PREMIUM AUR PROFESSIONAL PDF GENERATOR
def create_pdf(station_name, pairs_list):
    pdf = FPDF('L', 'mm', 'A4')
    
    for data in pairs_list:
        pdf.add_page()
        
        # 1. Premium Background Colour (Halka Professional Greyish-Blue)
        pdf.set_fill_color(248, 249, 250)
        pdf.rect(0, 0, 297, 210, 'F')
        
        # 2. Top Header Banner (Dark Blue)
        pdf.set_fill_color(0, 51, 153)
        pdf.rect(0, 0, 297, 25, 'F')
        
        # Banner Text (White)
        pdf.set_font("Arial", 'B', 22)
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(0, 5)
        pdf.cell(0, 15, txt=f"STATION CLEANLINESS REPORT : {station_name.upper()}", ln=1, align='C')
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_b:
            tmp_b.write(data['before'].getvalue())
            path_b = tmp_b.name
            
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_a:
            tmp_a.write(data['after'].getvalue())
            path_a = tmp_a.name
            
        # --- LEFT SIDE (BEFORE) ---
        # Photo Border
        pdf.set_line_width(0.8)
        pdf.set_draw_color(50, 50, 50)
        pdf.rect(15, 35, 125, 95, 'D')
        pdf.image(path_b, x=15, y=35, w=125, h=95)
        
        # 'BEFORE' Label (Red Background, White Text)
        pdf.set_fill_color(220, 53, 69) 
        pdf.rect(15, 135, 125, 12, 'F')
        pdf.set_xy(15, 135)
        pdf.set_font("Arial", 'B', 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(125, 12, txt="BEFORE", ln=1, align='C')
        
        # Date aur Time (Niche likha hua)
        pdf.set_xy(15, 149)
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(50, 50, 50) # Dark Grey
        pdf.cell(125, 8, txt=f"Date: {data['d_before']} | Time: {data['t_before']}", ln=1, align='C')
        
        # --- RIGHT SIDE (AFTER) ---
        # Photo Border
        pdf.set_line_width(0.8)
        pdf.set_draw_color(50, 50, 50)
        pdf.rect(155, 35, 125, 95, 'D')
        pdf.image(path_a, x=155, y=35, w=125, h=95)
        
        # 'AFTER' Label (Green Background, White Text)
        pdf.set_fill_color(40, 167, 69) 
        pdf.rect(155, 135, 125, 12, 'F')
        pdf.set_xy(155, 135)
        pdf.set_font("Arial", 'B', 16)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(125, 12, txt="AFTER", ln=1, align='C')
        
        # Date aur Time (Niche likha hua)
        pdf.set_xy(155, 149)
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(50, 50, 50) # Dark Grey
        pdf.cell(125, 8, txt=f"Date: {data['d_after']} | Time: {data['t_after']}", ln=1, align='C')
        
        # --- LOCATION BOX (SABSE NICHE BICH ME) ---
        pdf.set_fill_color(225, 235, 245) # Halka Neela (Light Blue) Box
        pdf.set_draw_color(0, 51, 153) # Dark blue border
        pdf.set_line_width(0.5)
        pdf.rect(20, 165, 257, 15, 'DF')
        
        pdf.set_xy(0, 167)
        pdf.set_font("Arial", 'B', 15)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(0, 10, txt=f"Location: {data['location']}", ln=1, align='C')
        
        # --- FOOTER (WATERMARK STYLE) ---
        pdf.set_xy(0, 190)
        pdf.set_font("Arial", 'I', 11)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(0, 10, txt="Central Railway - Solapur Division | Cleanliness Monitoring Dashboard", ln=1, align='C')
        
        os.remove(path_b)
        os.remove(path_a)
        
    return pdf.output(dest='S').encode('latin1')

st.title("🚆 Central Railway - Master Cleanliness Report")
st.markdown("**Searchable Dropdown, Custom Location & Premium PDF Design Added**")
st.markdown("---")

st.markdown("### 1. Enter Station Name")
station_input = st.text_input("Type the station name here:", placeholder="e.g. Solapur")

st.markdown("---")
st.markdown("### 2. Upload Photos or ZIP File")
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
        with st.spinner("Processing photos..."):
            for item in image_files:
                dt_obj, d_str, t_str = get_image_info(item['bytes'], item['name'])
                item['dt'] = dt_obj
                item['date_str'] = d_str
                item['time_str'] = t_str
                
                name_lower = item['name'].lower()
                if 'before' in name_lower or 'bfr' in name_lower:
                    item['priority'] = 0
                elif 'after' in name_lower or 'aft' in name_lower:
                    item['priority'] = 1
                else:
                    item['priority'] = 2
            
            image_files.sort(key=lambda x: (x['dt'], x['priority'], x['name']))
            st.success(f"✅ Total {len(image_files)} photos found.")
            
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
                        st.image(p_before['bytes'], caption=f"🔴 BEFORE ({p_before['time_str']})", use_container_width=True)
                    with col2:
                        st.image(p_after['bytes'], caption=f"🟢 AFTER ({p_after['time_str']})", use_container_width=True)
                        
                    with col4:
                        loc_choice = st.selectbox("👉 Select Track / Location (Type to Search):", LOCATION_OPTIONS, key=f"loc_{i}")
                        custom_loc = st.text_input("✍️ Ya Naya Custom Naam Likhein:", key=f"custom_loc_{i}", placeholder="Agar list me nahi hai...")
                    
                    inputs.append({
                        'before': p_before,
                        'after': p_after,
                        'loc_key': f"loc_{i}",
                        'custom_loc_key': f"custom_loc_{i}"
                    })
                
                st.write("---")
                submit = st.form_submit_button("3. Generate Premium Reports", type="primary")
                
            if submit:
                with st.spinner("Generating Professional PPT and PDF..."):
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
                            
                        pairs_list.append({
                            'before': process_image(item['before']['bytes']),
                            'after': process_image(item['after']['bytes']),
                            'd_before': item['before']['date_str'],
                            't_before': item['before']['time_str'],
                            'd_after': item['after']['date_str'],
                            't_after': item['after']['time_str'],
                            'location': loc_name
                        })
                    
                    st.session_state['ppt_data'] = create_ppt(station_input, pairs_list)
                    st.session_state['pdf_data'] = create_pdf(station_input, pairs_list)
                    st.session_state['report_ready'] = True

            if st.session_state.get('report_ready'):
                st.success("🎉 Reports are Ready! Download them below:")
                
                current_time_str = datetime.now().strftime('%H%M%S')
                
                col_ppt, col_pdf = st.columns(2)
                with col_ppt:
                    st.download_button(
                        label="⬇️ Download PowerPoint File", 
                        data=st.session_state['ppt_data'], 
                        file_name=f"{station_input}_Cleanliness_Report_{current_time_str}.pptx"
                    )
                with col_pdf:
                    st.download_button(
                        label="📥 Download Premium PDF Report", 
                        data=st.session_state['pdf_data'], 
                        file_name=f"{station_input}_Detailed_Report_{current_time_str}.pdf",
                        mime="application/pdf"
                    )
    else:
        st.warning("Please upload at least 2 photos!")
elif uploaded_files and not station_input:
    st.error("⚠️ Please make sure to enter the Station name above.")