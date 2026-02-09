import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
from datetime import datetime

# --- KONEKSI AMAN VIA SECRETS ---
@st.cache_resource
def get_gspread_client():
    # Mengambil data dari Streamlit Secrets
    info = st.secrets["gcp_service_account"]
    
    # Memperbaiki format private_key (mengubah \n menjadi baris baru asli)
    info_dict = dict(info)
    info_dict['private_key'] = info_dict['private_key'].replace('\\n', '\n')
    
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info_dict, scope)
    return gspread.authorize(creds)

# --- MULAI UI STREAMLIT (TEMA GELAP) ---
st.set_page_config(page_title="WSA Pro Analytics", layout="wide")

# CSS untuk mempercantik UI
st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    .metric-card {
        background-color: #1e2129;
        padding: 20px;
        border-radius: 12px;
        border-top: 4px solid #00d4ff;
        text-align: center;
        margin-bottom: 20px;
    }
    h1, h2, h3 { color: #00d4ff !important; }
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d4ff, #008fb3) !important;
        color: #000000 !important;
        font-weight: bold;
        width: 100%;
        border-radius: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# Logika Bulan Real-time
curr_month = datetime.now().month
prev_month = curr_month - 1 if curr_month > 1 else 12

with st.sidebar:
    st.title("⚙️ WSA Config")
    selected_months = st.multiselect("📅 Periode Bulan:", options=list(range(1, 13)), default=[prev_month, curr_month],
                                     format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1])

st.title("⚡ WSA FULFILLMENT ANALYTICS")
client = get_gspread_client()

if client:
    uploaded_file = st.file_uploader("DROP FILE XLSX / CSV DI SINI", type=["xlsx", "xls", "csv"])
    if uploaded_file:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        col_id = 'SC Order No/Track ID/CSRM No'
        if col_id in df.columns:
            # Filter Dasar
            df = df[df[col_id].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
            df['Date Created DT'] = pd.to_datetime(df['Date Created'], errors='coerce')
            
            # Filter Bulan
            if selected_months:
                df = df[df['Date Created DT'].dt.month.isin(selected_months)]
            
            # Filter Status Dinamis
            if 'Status' in df.columns:
                all_status = df['Status'].unique().tolist()
                selected_status = st.sidebar.multiselect("📍 Filter Status:", options=all_status, default=all_status)
                df = df[df['Status'].isin(selected_status)]

            # Cek Duplikat ke Google Sheets
            sheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT").sheet1
            google_df = pd.DataFrame(sheet.get_all_records())
            
            if not google_df.empty and col_id in google_df.columns:
                existing = google_df[col_id].astype(str).unique()
                df_final = df[~df[col_id].astype(str).isin(existing)].copy()
            else:
                df_final = df.copy()

            # UI Metrics
            c1, c2, c3 = st.columns(3)
            with c1: st.markdown(f'<div class="metric-card">📂 Total Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
            with c2: st.markdown(f'<div class="metric-card">✨ Data Baru Unik<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
            with c3: st.markdown(f'<div class="metric-card">🔗 Status GDoc<br><h2 style="color:#00ff88">Connected</h2></div>', unsafe_allow_html=True)

            # Preview & Download
            df_final['Date Created'] = df_final['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M:%S')
            st.dataframe(df_final.drop(columns=['Date Created DT']), use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.drop(columns=['Date Created DT']).to_excel(writer, index=False)
            st.download_button("📥 DOWNLOAD HASIL CLEANSING", output.getvalue(), "WSA_Report_Cleaned.xlsx")
