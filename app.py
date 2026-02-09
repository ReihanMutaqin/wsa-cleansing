import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
from datetime import datetime

# 1. SET THEME & DARK MODE
st.set_page_config(page_title="WSA Pro Dashboard", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    .metric-card {
        background-color: #1e2129;
        padding: 20px;
        border-radius: 12px;
        border-top: 4px solid #00d4ff;
        box-shadow: 0 4px 15px rgba(0,0,0,0.4);
        margin-bottom: 20px;
        text-align: center;
    }
    h1, h2, h3 { color: #00d4ff !important; }
    [data-testid="stFileUploader"] {
        background-color: #1a1c24;
        border: 2px dashed #00d4ff;
        border-radius: 15px;
        padding: 30px;
    }
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d4ff, #008fb3) !important;
        color: #000000 !important;
        font-weight: bold !important;
        width: 100%;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_resource
def get_client():
    json_file = "project-pengolahan-data-561c0b891db8.json"
    with open(json_file) as f:
        info = json.load(f)
    info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)

# --- LOGIKA BULAN REAL-TIME ---
current_month = datetime.now().month
# Jika sekarang Januari (1), maka bulan sebelumnya adalah Desember (12)
last_month = current_month - 1 if current_month > 1 else 12
default_months = [last_month, current_month]

# --- SIDEBAR ---
with st.sidebar:
    st.title("⚙️ WSA Config")
    st.markdown("---")
    selected_months = st.multiselect(
        "📅 Periode Bulan:", 
        options=list(range(1, 13)), 
        default=default_months, # Otomatis bulan sekarang & sebelumnya
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# --- MAIN CONTENT ---
st.title("⚡ WSA FULFILLMENT ANALYTICS")
st.write(f"Sistem mendeteksi bulan sekarang: **{datetime.now().strftime('%B')}**")

try:
    client = get_client()
    uploaded_file = st.file_uploader("DROP FILE XLSX / CSV DI SINI", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # Baca file sesuai format
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        col_target = 'SC Order No/Track ID/CSRM No'
        if col_target in df.columns:
            # 1. Filter Dasar
            df = df[df[col_target].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
            
            # 2. Filter Bulan Berdasarkan Pilihan (Default: Now & Last Month)
            df['Date Created DT'] = pd.to_datetime(df['Date Created'], errors='coerce')
            if selected_months:
                df = df[df['Date Created DT'].dt.month.isin(selected_months)]

            # 3. Filter Status Dinamis
            if 'Status' in df.columns:
                with st.sidebar:
                    st.markdown("---")
                    st.subheader("📍 Status Filter")
                    all_status = df['Status'].unique().tolist()
                    selected_status = st.multiselect("Pilih Status:", options=all_status, default=all_status)
                df = df[df['Status'].isin(selected_status)]

            # 4. Sinkronisasi Google Sheets
            sheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT").sheet1
            google_df = pd.DataFrame(sheet.get_all_records())
            
            if not google_df.empty and col_target in google_df.columns:
                existing = google_df[col_target].astype(str).unique()
                df_final = df[~df[col_target].astype(str).isin(existing)].copy()
            else:
                df_final = df.copy()

            # --- DISPLAY METRICS ---
            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(f'<div class="metric-card">📂 Total Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="metric-card">✨ Data Unik Baru<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
            with c3:
                st.markdown(f'<div class="metric-card">🔗 GDoc Status<br><h2 style="color:#00ff88">Connected</h2></div>', unsafe_allow_html=True)

            # Preview
            st.subheader("📋 Preview Data")
            df_final['Date Created'] = df_final['Date Created'].astype(str).str.split('.').str[0]
            st.dataframe(df_final.drop(columns=['Date Created DT']), use_container_width=True)

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.drop(columns=['Date Created DT']).to_excel(writer, index=False)
            
            st.download_button("📥 DOWNLOAD HASIL PROSES", output.getvalue(), "WSA_Report_Cleaned.xlsx")
        else:
            st.error(f"Kolom '{col_target}' tidak ditemukan!")

except Exception as e:
    st.error(f"Sistem Error: {e}")
