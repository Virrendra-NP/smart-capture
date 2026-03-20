import streamlit as st
import pandas as pd
from datetime import datetime
import os
from PIL import Image, ExifTags, ImageOps, ImageDraw, ImageFont
import io
import io
import json
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

tab_capture, tab_data = st.tabs(["📸 SMART CAPTURE", "📋 GLOBAL HISTORY"])

with tab_capture:
    # Mode Summary
    st.write(f"📝 **Reporting**: {category} at {loc}")
    st.write(f"📁 **Saving to**: {dest}")

    st.caption(f"🎯 **IPAD TIP**: Mode is {c_mode}. Tap below to Snap Photo.")
    
    # HERO CAMERA BUTTON
    photo = st.file_uploader("📸 TAP TO CAPTURE SITE EVIDENCE", type=['jpg', 'jpeg', 'png'])
    
    remarks = st.text_area("✍️ Site Remarks (Optional):", placeholder="Ex: Progress check on slab concrete.")

    if photo is not None and st.button("🚀 SUBMIT & SYNC TO CLOUD"):
        # Processing Logic
        now = datetime.now()
        ts = now.strftime('%Y%c%d_%H%M%S')
        img_filename = f"{ts}_{category[:10].replace(' ','_')}.jpg"
        
        # Ensure target dir exists
        os.makedirs(target_dir, exist_ok=True)
        img_path = os.path.join(target_dir, img_filename)
        
        # Save photo WITH ROTATION FIX & WATERMARK
        image = Image.open(photo)
        image = ImageOps.exif_transpose(image)
        
        # Watermark...
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
        
        # Record data...
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
    st.header("Global Site History")
    if os.path.exists(HISTORY_FILE):
        df_view = pd.read_csv(HISTORY_FILE)
        st.dataframe(df_view.sort_index(ascending=False), use_container_width=True, hide_index=True)
        st.download_button("⬇️ Download Master Log (CSV)", df_view.to_csv(index=False).encode('utf-8'), "Site_Log.csv", "text/csv")
    else:
        st.warning("No data yet.")

st.divider()
st.markdown("<p style='text-align: center; color: #64748b; font-size: 14px;'>🛡️ <b>SMART CAPTURE</b> | Site Intelligence Console</p>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #3b82f6; font-size: 16px;'><b>Professional App Developer: Virrendra</b></p>", unsafe_allow_html=True)
