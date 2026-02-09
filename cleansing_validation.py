import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from datetime import datetime

# --- KONFIGURASI HALAMAN & TEMA GELAP ---
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
    }
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d4ff, #008fb3) !important;
        color: #000000 !important;
        font-weight: bold !important;
        width: 100%;
        border: none;
        padding: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- KONEKSI KE GOOGLE SHEETS VIA SECRETS ---
@st.cache_resource
def get_gspread_client():
    try:
        # Mengambil data dari Streamlit Secrets (Bukan file file .json)
        secret_info = st.secrets["gcp_service_account"]
        
        # Memperbaiki format private key (mengembalikan karakter \n)
        info_dict = dict(secret_info)
        info_dict['private_key'] = info_dict['private_key'].replace('\\n', '\n')
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info_dict, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Koneksi GDoc Gagal. Pastikan Secrets sudah diisi! Error: {e}")
        return None

# --- LOGIKA BULAN OTOMATIS ---
curr_month = datetime.now().month
prev_month = curr_month - 1 if curr_month > 1 else 12
default_months = [prev_month, curr_month]

# --- SIDEBAR CONTROL ---
with st.sidebar:
    st.title("⚙️ WSA Control")
    st.markdown("---")
    selected_months = st.multiselect(
        "📅 Periode Bulan:", 
        options=list(range(1, 13)), 
        default=default_months,
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

st.title("⚡ WSA FULFILLMENT ANALYTICS")
st.write(f"Mode: **Otomasi Validasi GDoc** | Periode aktif: {datetime.now().strftime('%B %Y')}")

# --- PROSES UTAMA ---
client = get_gspread_client()

if client:
    uploaded_file = st.file_uploader("DROP FILE XLSX / CSV DI SINI", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # Baca file
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        col_id = 'SC Order No/Track ID/CSRM No'
        
        if col_id in df.columns:
            # 1. Filter AO/PDA/WSA & CRM Order Type
            df = df[df[col_id].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
            if 'CRM Order Type' in df.columns:
                df = df[df['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]

            # 2. Filter Bulan & Format Tanggal
            df['Date Created DT'] = pd.to_datetime(df['Date Created'], errors='coerce')
            if selected_months:
                df = df[df['Date Created DT'].dt.month.isin(selected_months)]

            # 3. Filter Status Dinamis (Sidebar)
            if 'Status' in df.columns:
                with st.sidebar:
                    st.markdown("---")
                    st.subheader("📍 Status Filter")
                    all_status = df['Status'].unique().tolist()
                    selected_status = st.multiselect("Pilih Status:", options=all_status, default=all_status)
                df = df[df['Status'].isin(selected_status)]

            # 4. Pembersihan Data (Logic asli Anda)
            df['SC Order No/Track ID/CSRM No'] = df[col_id].apply(lambda x: str(x).split('_')[0])
            df['Date Created'] = df['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M:%S')
            df['Booking Date'] = df['Booking Date'].astype(str).str.split('.').str[0]
            
            # 5. Sinkronisasi & Cek Duplikat ke Google Sheets
            try:
                sheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT").sheet1
                google_df = pd.DataFrame(sheet.get_all_records())
                
                if not google_df.empty and col_id in google_df.columns:
                    existing = google_df[col_id].astype(str).unique()
                    df_final = df[~df[col_id].astype(str).isin(existing)].copy()
                else:
                    df_final = df.copy()

                # --- DISPLAY ---
                c1, c2, c3 = st.columns(3)
                with c1: st.markdown(f'<div class="metric-card">📂 Total Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                with c2: st.markdown(f'<div class="metric-card">✨ Data Unik Baru<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                with c3: st.markdown(f'<div class="metric-card">🔗 GDoc Status<br><h2 style="color:#00ff88">Connected</h2></div>', unsafe_allow_html=True)

                st.subheader("📋 Preview Data Validasi")
                cols_to_show = ['Workzone', 'Date Created', 'SC Order No/Track ID/CSRM No', 'Service No.', 'Workorder', 'Customer Name', 'Address', 'Contact Number', 'CRM Order Type','Booking Date']
                st.dataframe(df_final[cols_to_show].sort_values('Workzone'), use_container_width=True)

                # --- DOWNLOAD ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[cols_to_show].to_excel(writer, index=False)
                
                st.download_button("📥 DOWNLOAD DATA UNIK", output.getvalue(), "WSA_Report_Cleaned.xlsx")

            except Exception as e:
                st.error(f"Gagal akses Google Sheets: {e}")
        else:
            st.error(f"Kolom '{col_id}' tidak ditemukan!")
