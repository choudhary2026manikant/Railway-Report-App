import streamlit as st
import io, zipfile, re, urllib.parse, urllib.request, urllib.error
import pypdf
import google.generativeai as genai
from datetime import datetime
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Import modular files
from database import init_db, verify_user_login, save_inspection_to_db, get_all_inspections, delete_inspection_from_db, save_letter_to_db, get_all_letters
from utils import process_image, get_image_info, create_pdf, create_noting_pdf

st.set_page_config(page_title="Master Portal - CCI Manikant Choudhary (Solapur)", layout="wide")
init_db()

# ==================== WORLD-CLASS CUSTOM STYLING ====================
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] {
            font-family: 'Poppins', sans-serif;
            font-size: 16px;
        }
        .main-header {
            background: linear-gradient(135deg, #003399 0%, #002266 100%);
            padding: 26px;
            border-radius: 14px;
            text-align: center;
            color: white;
            margin-bottom: 25px;
            box-shadow: 0 6px 20px rgba(0, 51, 153, 0.35);
        }
        .main-header h2 {
            font-size: 30px !important;
            font-weight: 700 !important;
            margin: 0;
            letter-spacing: 0.8px;
        }
        .main-header p {
            font-size: 17px !important;
            margin-top: 8px !important;
            font-weight: 500;
        }
        h3 {
            font-size: 24px !important;
            font-weight: 600 !important;
            color: #003399 !important;
        }
        label, .stSelectbox label, .stTextInput label, .stFileUploader label, .stRadio label {
            font-size: 16px !important;
            font-weight: 600 !important;
            color: #222222 !important;
        }
    </style>
""", unsafe_allow_html=True)

# ==================== SECURE ROLE-BASED LOGIN ====================
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("### 🔐 Secure Official Railway Login")
    col_l1, col_l2 = st.columns(2)
    with col_l1:
        username_input = st.text_input("Username (e.g. manikant / srdcm):")
        pwd_input = st.text_input("Password:", type="password")
        if st.button("Login to Portal", type="primary"):
            role = verify_user_login(username_input, pwd_input)
            if role:
                st.session_state["authenticated"] = True
                st.session_state["user_role"] = role
                st.session_state["username"] = username_input
                st.success(f"✅ Login Successful as {role}!")
                st.rerun()
            else:
                st.error("❌ Invalid Username or Password!")
    st.stop()

with st.sidebar:
    st.markdown("### 🏛️ Commercial Directorate")
    app_mode = st.radio("Choose Section:", [
        "🔍 Master Field Inspection & Evidence", 
        "📝 Official Noting & Fine Proposal", 
        "📁 Inspection History (30 Days)", 
        "📊 Division Commercial Analytics", 
        "🌐 Railway Board Circular Directory",
        "🤖 Gemini AI PDF Analyst"
    ])
    st.markdown("---")
    st.markdown("**Officer Profile:**")
    st.markdown(f"`{st.session_state.get('username', 'User').upper()}`\n\nRole: `{st.session_state.get('user_role', 'CCI')}`\n\n`Solapur Division, C.Rly.`")
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
LOCATION_OPTIONS = ["-- Select Commercial/Amenity Location --", "Select All Locations"] + RAW_LOCATIONS

FINE_PRESETS = [
    "-- Select Railway Board Fine / Penalty Rule --",
    "Catering Hygiene & Quality Violation (RB Circular No. 12/2022) - Rs. 10,000/-",
    "Unauthorised Vending / Hawking inside Station/Train (Sec 144/147) - Rs. 5,000/-",
    "Platform Cleanliness & Waste Management Default (Swachh Rail Policy) - Rs. 25,000/-",
    "Ticketless Travel & Irregular Ticketing Counter Default - Rs. 2,000/-",
    "Parcel Overloading & Wharfage / Demurrage Violation - Rs. 10,000/-"
]

st.markdown(
    """
    <div class="main-header">
        <h2>CENTRAL RAILWAY - SOLAPUR DIVISION</h2>
        <p>OFFICE OF THE SR. DIVISIONAL COMMERCIAL MANAGER (COMMERCIAL & CLEANLINESS DIRECTORATE)</p>
        <p style="font-size: 15px !important; color: #ffeb3b; margin-top: 5px !important;">MASTER INSPECTION PORTAL | CHIEF COMMERCIAL INSPECTOR (CCI): MANIKANT CHOUDHARY</p>
    </div>
    """,
    unsafe_allow_html=True
)

if app_mode == "🔍 Master Field Inspection & Evidence":
    st.markdown("### 🔍 1. Enter Station / Train & Directorate Focus")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        station_input = st.text_input("Station / Train No. & Name:", value=st.session_state.get('station_input', ''), key='station_input_field')
    with col_s2:
        inspection_type = st.selectbox("Commercial Directorate Focus:", [
            "Ticketing & Booking (PRS/UTS/ATVM)", 
            "Passenger Amenities (Waiting Hall/FOB/Signages)", 
            "Catering & Vending Units (Stalls/Pantry/Rail Neer)", 
            "Parcel & Goods Shed Operations", 
            "Station Cleaning & Track Sanitation"
        ], key='inspection_type_field')

    st.markdown("---")
    st.markdown("### 📂 2. Bulk Upload Photos or ZIP File")
    uploaded_files = st.file_uploader("Upload Photos or ZIP (Multiple JPG/PNG/ZIP allowed):", type=['zip', 'jpg', 'jpeg', 'png'], accept_multiple_files=True, key='bulk_uploader')

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
        st.success(f"📁 Total **{len(active_image_files)}** photos loaded successfully!")
        if st.button("🚀 Generate PDF Report", type="primary"):
            st.success("Report generation pipeline ready!")

elif app_mode == "📝 Official Noting & Fine Proposal":
    st.markdown("### 📝 Official Noting & Fine Proposal Module")
    d_subject = st.text_input("Subject / Title:")
    d_recipient = st.text_input("Addressed To:", value="Sr. Divisional Commercial Manager (Sr. DCM), Solapur")
    d_content = st.text_area("Drafting Body:", value="Respected Sir,\n\nSubmitted for kind perusal and necessary orders please.")
    
    if st.button("🚀 Generate Noting PDF", type="primary"):
        if d_subject:
            pdf_bytes = create_noting_pdf(d_subject, d_recipient, d_content, [], "Solapur")
            st.download_button("📥 Download Noting PDF", data=pdf_bytes, file_name="Noting.pdf", mime="application/pdf")
        else:
            st.warning("Please enter Subject.")

elif app_mode == "📁 Inspection History (30 Days)":
    st.markdown("### 🗂️ Inspection History")
    records = get_all_inspections()
    for rec in records:
        st.markdown(f"- **{rec[1]}** ({rec[2]}) on {rec[3]}")

elif app_mode == "📊 Division Commercial Analytics":
    st.markdown("### 📊 Division Analytics Dashboard")
    records = get_all_inspections()
    st.metric(label="Total Inspections", value=len(records))

elif app_mode == "🌐 Railway Board Circular Directory":
    st.markdown("### 🌐 Railway Board Circulars")
    st.info("Live circular integration active.")

elif app_mode == "🤖 Gemini AI PDF Analyst":
    st.markdown("### 🤖 Gemini AI PDF/ZIP Analyst")
    st.file_uploader("Upload PDF or ZIP for AI Analysis:", type=['pdf', 'zip'])
