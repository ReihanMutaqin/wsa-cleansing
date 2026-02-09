import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# 1. SET THEME & DARK MODE CSS
st.set_page_config(page_title="WSA Pro Analytics", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    .metric-card {
        background-color: #1e2129;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #00d4ff;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        margin-bottom: 10px;
    }
    h1, h2, h3 { color: #00d4ff !important; }
    .stDownloadButton button {
        background-color: #00d4ff !important;
        color: #0e1117 !important;
        font-weight: bold;
        border-radius: 8px;
        width: 100%;
    }
    /* Mengatur tampilan dropzone */
    [data-testid="stFileUploader"] {
        background-color: #1a1c24;
        border: 2px dashed #00d4ff;
        border-radius: 10px;
        padding: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_resource
def get_client():
    # File JSON terbaru Anda
    json_file = "project-pengolahan-data-561c0b891db8.json"
    with open(json_file) as f:
        info = json.load(f)
    info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)

# --- SIDEBAR ---
with st.sidebar:
    st.title("🌑 WSA Control")
    st.markdown("---")
    selected_months = st.multiselect(
        "📅 Pilih Bulan:", 
        options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], 
        default=[1, 2],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# --- MAIN CONTENT ---
st.title("⚡ WSA FULFILLMENT DASHBOARD")
st.write("Drag & Drop file report (Excel atau CSV) di bawah ini.")

try:
    client = get_client()
    # Support multiple extensions
    uploaded_file = st.file_uploader("Seret file ke sini atau klik untuk upload", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # LOGIKA PEMBACAAN FILE (XLSX vs CSV)
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # 1. Filter WSA/AO/PDA
        if 'SC Order No/Track ID/CSRM No' in df.columns:
            df = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
            
            # 2. Filter Bulan
            df['Date Created DT'] = pd.to_datetime(df['Date Created'], errors='coerce')
            if selected_months:
                df = df[df['Date Created DT'].dt.month.isin(selected_months)]

            # 3. FILTER STATUS OTOMATIS
            if 'Status' in df.columns:
                all_status = df['Status'].unique().tolist()
                with st.sidebar:
                    st.subheader("📌 Filter Status")
                    selected_status = st.multiselect("Pilih Status:", options=all_status, default=all_status)
                df = df[df['Status'].isin(selected_status)]

            # 4. CEK DUPLIKAT GDOC
            sheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT").sheet1
            google_df = pd.DataFrame(sheet.get_all_records())
            
            if not google_df.empty and 'SC Order No/Track ID/CSRM No' in google_df.columns:
                existing = google_df['SC Order No/Track ID/CSRM No'].astype(str).unique()
                df_final = df[~df['SC Order No/Track ID/CSRM No'].astype(str).isin(existing)].copy()
            else:
                df_final = df.copy()

            # --- UI DISPLAY (METRICS) ---
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="metric-card">Total Data<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="metric-card">Data Unik Baru<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
            with c3:
                st.markdown(f'<div class="metric-card">Status Koneksi<br><h2 style="color:#00ff88">ONLINE</h2></div>', unsafe_allow_html=True)

            st.markdown("### 📋 Preview Data")
            df_final['Date Created'] = df_final['Date Created'].astype(str).str.split('.').str[0]
            st.dataframe(df_final.drop(columns=['Date Created DT']), use_container_width=True)

            # --- DOWNLOAD ---
            st.markdown("---")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.drop(columns=['Date Created DT']).to_excel(writer, index=False)
            
            st.download_button("📥 DOWNLOAD DATA CLEANSING", output.getvalue(), "wsa_report_final.xlsx")
        else:
            st.error("Kolom 'SC Order No/Track ID/CSRM No' tidak ditemukan dalam file!")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
