import streamlit as st
import pandas as pd
from datetime import datetime
import os
from PIL import Image, ExifTags, ImageOps, ImageDraw, ImageFont
import io
try:
    from streamlit_js_eval import get_geolocation
except ImportError:
    get_geolocation = None

# --- PAGE CONFIGURATION (Premium App Look) ---
st.set_page_config(
    page_title="Smart Capture",
    page_icon="👷",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom Design for Site Use (High Contrast, Large Buttons)
st.markdown("""
    <style>
    /* Dark Slate Premium Background */
    .stApp { background-color: #0f172a; color: #f1f5f9; font-family: 'Segoe UI', sans-serif; }
    
    /* Header Styling */
    h1 { color: #3b82f6; text-align: center; border-bottom: 2px solid #334155; padding-bottom: 10px; }
    
    /* Big Submit Button */
    .stButton>button { 
        width: 100%; 
        height: 70px; 
        font-size: 20px; 
        font-weight: bold; 
        background-color: #3b82f6 !important; 
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s;
    }
    .stButton>button:active { transform: scale(0.98); }
    
    /* Camera Block */
    .stCamera { border: 3px solid #3b82f6; border-radius: 15px; overflow: hidden; }
    
    /* Input Highlights */
    .stTextInput>div>div>input, .stSelectbox>div>div>div { background-color: #1e293b !important; color: white !important; border: 1px solid #334155 !important; }
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

with tab_capture:
    st.header("Step 1: Activity & Cloud Destination")
    
    col1, col2 = st.columns(2)
    with col1:
        category = st.selectbox("Work Item Type:", ["Earth Work", "Shuttering / Formwork", "Building Work", "Reinforcement", "Other (Type Below)"])
        if category == "Other (Type Below)":
            custom_work = st.text_input("📝 Custom Work Item:")
            category = custom_work if custom_work else "Custom Item"
            
    with col2:
        dest = st.selectbox("☁️ SAVE TO CLOUD DRIVE:", list(CLOUD_PATHS.keys()))
        target_dir = CLOUD_PATHS.get(dest, CLOUD_PATHS["Local Vault (Office PC)"])
        if not os.path.exists(target_dir):
            st.warning(f"⚠️ {dest} folder not found on PC. Saving to Local Vault instead.")
            target_dir = CLOUD_PATHS["Local Vault (Office PC)"]
    
    col3, col4 = st.columns(2)
    with col3:
        loc = st.text_input("📍 Site Location / Block:", value="Main Building")
    with col4:
        remarks = st.text_area("📝 Remarks / Descriptions:", placeholder="Ex: Progress 40% complete.")

    st.header("Step 2: Start Site Capture")
    st.info("🎯 **IPAD/IPHONE TIP**: Tap the button below and select **'Take Photo'**. This opens the Real Apple Camera where you can switch between **Front and Rear** cameras normally!")
    
    # ONE BUTTON FOR ALL: Opens native camera or library
    photo = st.file_uploader("📸 SNAP OR SELECT SITE PHOTO", type=['jpg', 'jpeg', 'png'])
    
    if photo is not None and st.button("🚀 SUBMIT & SYNC TO CLOUD"):
        now = datetime.now()
        ts = now.strftime('%Y%m%d_%H%M%S')
        img_filename = f"{ts}_{category[:10].replace(' ','_')}.jpg"
        
        # Ensure target dir exists
        os.makedirs(target_dir, exist_ok=True)
        img_path = os.path.join(target_dir, img_filename)
        
        # Save photo WITH ROTATION FIX & WATERMARK
        image = Image.open(photo)
        image = ImageOps.exif_transpose(image)
        
        # --- SMART WATERMARK ENGINE ---
        try:
            draw = ImageDraw.Draw(image)
            w, h = image.size
            
            # Formatted Data: "19-Mar-2026 (Thursday) 23:12 | GPS: 12.345, 67.890 | Virrendra"
            dt_str = now.strftime('%d-%b-%Y (%A) %H:%M')
            watermark_text = f"{dt_str} | GPS: {gps_str} | Developer: Virrendra"
            
            # Dynamic Font Size (approx 3% of image height)
            f_size = max(24, int(h * 0.03))
            try: font = ImageFont.truetype("arial.ttf", f_size)
            except: font = ImageFont.load_default()
            
            # Position: Bottom Right (with 2% padding)
            pad_x, pad_y = int(w * 0.02), int(h * 0.05)
            
            # Draw with high-contrast STROKE (so it works on white/black photos)
            draw.text((w - pad_x, h - pad_y), watermark_text, 
                      fill="white", font=font, anchor="rs", 
                      stroke_width=max(1, f_size//20), stroke_fill="black")
        except Exception as e:
            st.error(f"Watermark Error: {e}")

        # Save Final Processed Image
        image.save(img_path, quality=90, optimize=True)
        
        # Record data
        record = {
            'Timestamp': now.strftime('%d-%m-%Y %H:%M'),
            'Category': category,
            'Location': loc,
            'GPS': gps_str,
            'Destination': dest,
            'Remarks': remarks,
            'Photo_Name': img_filename
        }
        
        if os.path.exists(HISTORY_FILE):
            df = pd.read_csv(HISTORY_FILE)
            df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
        else:
            df = pd.DataFrame([record])
        
        df.to_csv(HISTORY_FILE, index=False)
        st.success(f"✅ RECORDED! Saved to: {dest} (Watermark Created!)")
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
