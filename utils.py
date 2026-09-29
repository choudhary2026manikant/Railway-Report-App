import io, os, tempfile, re
from PIL import Image, ImageEnhance, ImageOps, ExifTags
from fpdf import FPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from datetime import datetime

def clean_text_for_pdf(text):
    if not text:
        return ""
    return ''.join(c for c in text if ord(c) < 256)

def process_image(img_bytes):
    try:
        img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        img = ImageOps.fit(img, (800, 600), Image.Resampling.LANCZOS)
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.05)
        
        img_byte_arr = io.BytesIO()
        img.save(img_byte_arr, format='JPEG', quality=85, optimize=True)
        img_byte_arr.seek(0)
        return img_byte_arr
    except Exception:
        blank_img = Image.new('RGB', (800, 600), color=(200, 200, 200))
        img_byte_arr = io.BytesIO()
        blank_img.save(img_byte_arr, format='JPEG')
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
            
    return dt_obj, gps_info

def create_pdf(station_name, insp_type, items_list, layout_mode):
    pdf = FPDF('L', 'mm', 'A4')
    pdf.set_auto_page_break(False)
    
    for data in items_list:
        pdf.add_page()
        pdf.set_fill_color(248, 249, 250)
        pdf.rect(0, 0, 297, 210, 'F')
        
        pdf.set_fill_color(0, 51, 153)
        pdf.rect(0, 0, 297, 20, 'F')
        pdf.set_font("Arial", 'B', 17)
        pdf.set_text_color(255, 255, 255)
        pdf.set_xy(0, 3)
        pdf.cell(0, 14, txt=clean_text_for_pdf(f"COMMERCIAL INSPECTION REPORT : {station_name.upper()} [{insp_type}]"), ln=1, align='C')
        
        pdf.set_xy(15, 30)
        pdf.set_font("Arial", 'B', 12)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(0, 10, txt=clean_text_for_pdf(f"Inspection Details & Evidence Records"), ln=1)

    raw_output = pdf.output()
    if isinstance(raw_output, str):
        return raw_output.encode('latin1')
    elif isinstance(raw_output, bytearray):
        return bytes(raw_output)
    return raw_output

def create_noting_pdf(subject, recipient, content, photos_bytes_list, location_str, fine_rule_str=""):
    pdf = FPDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    
    pdf.set_fill_color(0, 51, 153)
    pdf.rect(0, 0, 210, 20, 'F')
    pdf.set_font("Arial", 'B', 15)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(0, 3)
    pdf.cell(210, 14, txt="CENTRAL RAILWAY - SOLAPUR DIVISION", ln=1, align='C')
    
    pdf.set_xy(15, 25)
    pdf.set_font("Arial", 'B', 11)
    pdf.set_text_color(0, 51, 153)
    pdf.cell(0, 6, txt=clean_text_for_pdf(f"To: {recipient}"), ln=1)
    
    pdf.set_xy(15, 33)
    pdf.set_font("Arial", 'B', 11)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(180, 6, txt=clean_text_for_pdf(f"Subject: {subject}"))
    
    current_y = pdf.get_y() + 4
    if location_str and location_str != "-- Select Commercial/Amenity Location --":
        pdf.set_xy(15, current_y)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(180, 6, txt=clean_text_for_pdf(f"Inspection Location / Area: {location_str}"), ln=1)
        current_y = pdf.get_y() + 2
        
    if fine_rule_str and fine_rule_str != "-- Select Railway Board Fine / Penalty Rule --":
        pdf.set_xy(15, current_y)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(180, 0, 0)
        pdf.multi_cell(180, 6, txt=clean_text_for_pdf(f"Recommended Fine / Penalty Reference: {fine_rule_str}"))
        current_y = pdf.get_y() + 4

    pdf.set_xy(15, current_y)
    pdf.set_font("Arial", '', 10)
    pdf.set_text_color(20, 20, 20)
    pdf.multi_cell(180, 6, txt=clean_text_for_pdf(content))
    
    if photos_bytes_list:
        pdf.ln(5)
        pdf.set_font("Arial", 'B', 10)
        pdf.set_text_color(0, 51, 153)
        pdf.cell(0, 6, txt=clean_text_for_pdf(f"Attached Evidence Photos [Location: {location_str}]:"), ln=1)
        
        for p_bytes in photos_bytes_list:
            pdf.add_page()
            processed_io = process_image(p_bytes)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp_n:
                tmp_n.write(processed_io.getvalue())
                path_n = tmp_n.name
            pdf.image(path_n, x=25, y=30, w=160, h=110)
            os.remove(path_n)
            pdf.ln(120)
            
    pdf.ln(15)
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
