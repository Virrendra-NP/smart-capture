import streamlit as st
import pandas as pd
from datetime import datetime
import os
from PIL import Image, ExifTags, ImageOps, ImageDraw, ImageFont
import io
import io
import json
import io
import json
import xlsxwriter
from io import BytesIO
try:
    from streamlit_js_eval import get_geolocation
except ImportError:
    get_geolocation = None

# --- PREFERENCE ENGINE & PERSISTENCE ---
PREF_FILE = "user_settings.json"

def load_prefs():
    if os.path.exists(PREF_FILE):
        try:
            with open(PREF_FILE, "r") as f: return json.load(f)
        except: pass
    return {"category": "Building Work", "loc": "Main Building", "dest": "Local Vault (Office PC)", "mode": "📸 LIVE SITE PHOTO"}

def save_prefs(data):
    try:
        with open(PREF_FILE, "w") as f: json.dump(data, f)
    except: pass

prefs = load_prefs()

# --- PAGE CONFIGURATION (Premium App Look) ---
st.set_page_config(
    page_title="Smart Capture",
    page_icon="👷",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Custom iOS Aesthetic Design (Glassmorphism & SF Pro)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
    
    /* Apple SF Pro Style Fonts */
    html, body, [class*="css"] { 
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Inter", sans-serif !important; 
        background-color: #000000;
    }
    
    .stApp { background-color: #000000; color: #f8fafc; }
    
    /* Sidebar as a Professional "Settings Drawer" */
    section[data-testid="stSidebar"] {
        background-color: rgba(30, 41, 59, 0.7) !important;
        backdrop-filter: blur(20px);
        border-right: 1px solid rgba(255,255,255,0.1);
    }
    
    /* Big iOS Style Buttons */
    .stButton>button { 
        width: 100%; 
        height: 75px; 
        font-size: 20px; 
        font-weight: 700; 
        background-color: #007aff !important; 
        border-radius: 14px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3);
        border: none;
    }
    
    /* Input Style (iOS Settings Cell) */
    .stTextInput>div>div>input, .stSelectbox>div>div>div, .stTextArea>div>textarea { 
        background-color: #1c1c1e !important; 
        color: #FFFFFF !important; 
        border: 1px solid #38383a !important; 
        border-radius: 12px !important;
    }
    
    /* Logo / Header */
    .header-text { font-size: 28px; font-weight: 800; color: #007aff; text-align: center; margin-bottom: 20px; }
    </style>
""", unsafe_allow_html=True)

# --- STORAGE SETUP & CLOUD SYNC ---
USER_PROFILE = os.environ.get('USERPROFILE', 'C:/Users/Administrator')
CLOUD_PATHS = {
    "Local Vault (Office PC)": os.path.join(os.getcwd(), "Smart_Capture_Vault"),
    "Google Drive": os.path.join(USER_PROFILE, "Google Drive"),
    "Dropbox": os.path.join(USER_PROFILE, "Dropbox"),
    "OneDrive": os.path.join(USER_PROFILE, "OneDrive"),
    "iCloud (Files)": os.path.join(USER_PROFILE, "iCloudDrive")
}

# Ensure Local Vault exists
for folder in CLOUD_PATHS.values():
    if "Local" in folder: os.makedirs(folder, exist_ok=True)

HISTORY_FILE = "Smart_Capture_Global_Log.csv"

# --- APP HEADER ---
st.title("💠 BUILDING SMART CAPTURE")
st.write("Professional Site Reporting & Cloud Visual Terminal.")

# --- LIVE LOCATION ---
gps_str = "No GPS Access"
if get_geolocation:
    try:
        loc_data = get_geolocation()
        if loc_data and 'coords' in loc_data:
            lat, lon = loc_data['coords']['latitude'], loc_data['coords']['longitude']
            gps_str = f"{round(lat, 5)}, {round(lon, 5)}"
            st.sidebar.success(f"📍 GPS Found: {gps_str}")
    except: pass

# --- APP TABS ---
tab_capture, tab_data = st.tabs(["📸 NEW CAPTURE", "📋 GLOBAL HISTORY"])

# --- SIDEBAR: SETTINGS GEAR (⚙️) ---
with st.sidebar:
    st.markdown("<h2 style='color:#007aff'>⚙️ APP SETTINGS</h2>", unsafe_allow_html=True)
    st.divider()
    
    # 1. Capture Choice (Sticky)
    modes = ["📸 LIVE SITE PHOTO", "📁 BACKUP FROM PHOTOS"]
    m_idx = modes.index(prefs["mode"]) if prefs["mode"] in modes else 0
    c_mode = st.radio("Default Camera Mode:", modes, index=m_idx)
    
    # 2. Work Item (Sticky)
    cat_list = ["Earth Work", "Shuttering / Formwork", "Building Work", "Reinforcement", "Concrete Pouring", "Other (Type Below)"]
    c_idx = cat_list.index(prefs["category"]) if prefs["category"] in cat_list else 5
    category = st.selectbox("Current Work Item:", cat_list, index=c_idx)
    if category == "Other (Type Below)":
        category = st.text_input("📝 Custom Work Type:", value=prefs.get("custom_category", ""))
    
    # 3. Location (Sticky)
    loc = st.text_input("📍 Default Location/Block:", value=prefs["loc"])
    
    # 4. Destination (Sticky)
    d_list = list(CLOUD_PATHS.keys())
    d_idx = d_list.index(prefs["dest"]) if prefs["dest"] in d_list else 0
    dest = st.selectbox("☁️ CLOUD DESTINATION:", d_list, index=d_idx)
    target_dir = CLOUD_PATHS.get(dest, CLOUD_PATHS["Local Vault (Office PC)"])

    st.divider()
    st.markdown("<p style='text-align: center; color: #3b82f6; font-size: 16px;'><b>Professional App Developer: Virrendra</b></p>", unsafe_allow_html=True)

# --- AUTO-SAVE PREFERENCES (Native Style) ---
save_prefs({
    "category": category, 
    "loc": loc, 
    "dest": dest, 
    "mode": c_mode, 
    "custom_category": category if category not in cat_list else ""
})

# --- MAIN PAGE: CAMERA FIRST ---
st.markdown("<div class='header-text'>BUILDING SMART CAPTURE</div>", unsafe_allow_html=True)

tab_capture, tab_data, tab_excel = st.tabs(["📸 SMART CAPTURE", "📋 HISTORY", "📑 EXCEL GENERATOR"])

with tab_capture:
    st.write(f"📝 **Reporting**: {category} at {loc}")
    st.write(f"📁 **Saving to**: {dest}")
    st.caption(f"🎯 **IPAD TIP**: Mode is {c_mode}. Tap below to Snap Photo.")
    
    photo = st.file_uploader("📸 TAP TO CAPTURE SITE EVIDENCE", type=['jpg', 'jpeg', 'png'])
    remarks = st.text_area("✍️ Site Remarks (Optional):", placeholder="Ex: Progress check on slab concrete.")

    if photo is not None and st.button("🚀 SUBMIT & SYNC TO CLOUD"):
        now = datetime.now()
        ts = now.strftime('%Y%c%d_%H%M%S')
        img_filename = f"{ts}_{category[:10].replace(' ','_')}.jpg"
        os.makedirs(target_dir, exist_ok=True)
        img_path = os.path.join(target_dir, img_filename)
        image = Image.open(photo)
        image = ImageOps.exif_transpose(image)
        try:
            draw = ImageDraw.Draw(image)
            w, h = image.size
            dt_str = now.strftime('%d-%b-%Y (%A) %H:%M')
            watermark_text = f"{dt_str} | GPS: {gps_str} | Developer: Virrendra"
            f_size = max(24, int(h * 0.03))
            try: font = ImageFont.truetype("arial.ttf", f_size)
            except: font = ImageFont.load_default()
            pad_x, pad_y = int(w * 0.02), int(h * 0.05)
            draw.text((w - pad_x, h - pad_y), watermark_text, fill="white", font=font, anchor="rs", stroke_width=max(1, f_size//20), stroke_fill="black")
        except: pass
        image.save(img_path, quality=90, optimize=True)
        record = {'Timestamp': now.strftime('%d-%m-%Y %H:%M'), 'Category': category, 'Location': loc, 'GPS': gps_str, 'Destination': dest, 'Remarks': remarks, 'Photo_Name': img_filename}
        if os.path.exists(HISTORY_FILE):
            df = pd.read_csv(HISTORY_FILE)
            df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
        else:
            df = pd.DataFrame([record])
        df.to_csv(HISTORY_FILE, index=False)
        st.success(f"✅ RECORDED SUCCESSFULLY! (Synced to Drive)")
        st.balloons()

with tab_data:
    st.header("📋 GLOBAL SITE LOG")
    if os.path.exists(HISTORY_FILE):
        df_view = pd.read_csv(HISTORY_FILE)
        st.dataframe(df_view.sort_index(ascending=False), use_container_width=True, hide_index=True)
    else:
        st.warning("No records found.")

with tab_excel:
    st.header("🚀 BATCH TO EXCEL")
    if os.path.exists(HISTORY_FILE):
        df_ex = pd.read_csv(HISTORY_FILE)
        df_ex['Display'] = df_ex['Timestamp'] + " - " + df_ex['Category'] + " (" + df_ex['Location'] + ")"
        
        selected_rows = st.multiselect("Select Photos to Transfer:", df_ex.index, format_func=lambda i: df_ex.iloc[i]['Display'])
        
        if st.button("📊 CREATE PROFESSIONAL EXCEL REPORT"):
            if not selected_rows:
                st.error("Please select at least one photo!")
            else:
                output = BytesIO()
                workbook = xlsxwriter.Workbook(output)
                sheet = workbook.add_worksheet("Site Report")
                
                # Formats
                lbl_fmt = workbook.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                val_fmt = workbook.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
                
                # Setup Column Widths (189 pixels is approx 26 units)
                sheet.set_column(0, 3, 15) # Standard
                sheet.set_column(0, 0, 26) # Image Column
                
                row = 0
                for idx in selected_rows:
                    data = df_ex.iloc[idx]
                    
                    # Row 1: Header (Metadata Cells)
                    sheet.set_row(row, 20)
                    sheet.write(row, 0, "ACTIVITY:", lbl_fmt); sheet.write(row, 1, data['Category'], val_fmt)
                    sheet.write(row, 2, "LOCATION:", lbl_fmt); sheet.write(row, 3, data['Location'], val_fmt)
                    
                    # Row 2: Header (Time/GPS)
                    sheet.set_row(row+1, 20)
                    sheet.write(row+1, 0, "TIMESTAMP:", lbl_fmt); sheet.write(row+1, 1, data['Timestamp'], val_fmt)
                    sheet.write(row+1, 2, "GPS INFO:", lbl_fmt); sheet.write(row+1, 3, data['GPS'], val_fmt)
                    
                    # Row 3: Image (Height set to 142 points = 5cm)
                    sheet.set_row(row+2, 142) 
                    
                    img_name = data['Photo_Name']
                    found = False
                    for b_path in CLOUD_PATHS.values():
                        p_file = os.path.join(b_path, img_name)
                        if os.path.exists(p_file):
                            try:
                                # Calculate scale for 5cm (approx 189 pixels)
                                with Image.open(p_file) as img:
                                    w_px, h_px = img.size
                                    x_scale = 189.0 / w_px
                                    y_scale = 189.0 / h_px
                                sheet.insert_image(row+2, 0, p_file, {'x_scale': x_scale, 'y_scale': y_scale, 'x_offset': 5, 'y_offset': 5})
                                found = True
                            except: pass
                            break
                    
                    if not found: sheet.write(row+2, 0, "Photo missing in storage")
                    
                    # NEXT RECORD: Skip 1 row (Photo Row is r+2, Empty Row is r+3, Next is r+4)
                    row += 4 
                
                workbook.close()
                output.seek(0)
                st.download_button("⬇️ DOWNLOAD YOUR SITE REPORT", output, f"SiteReport_{datetime.now().strftime('%d_%b')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No data found to transform.")

st.divider()
st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px;'>🛡️ <b>SMART CAPTURE</b> | Professional Site Dashboard</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #3b82f6; font-size: 16px;'><b>App Developer: Virrendra</b></p>", unsafe_allow_html=True)
