import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
from datetime import datetime

# --- 1. KONFIGURASI HALAMAN & TEMA ---
st.set_page_config(page_title="WSA Multi-Tool Analytics", layout="wide")

st.markdown("""
    <style>
    /* Dark Mode Theme */
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    
    /* Card Statistik */
    .metric-card {
        background-color: #1e2129;
        padding: 20px;
        border-radius: 12px;
        border-top: 4px solid #00d4ff;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    
    /* Typography */
    h1, h2, h3 { color: #00d4ff !important; }
    
    /* Tombol Download */
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d4ff, #008fb3) !important;
        color: #000000 !important;
        font-weight: bold;
        width: 100%;
        border: none;
        padding: 12px;
        border-radius: 8px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- 2. KONEKSI GOOGLE SHEETS (VIA SECRETS) ---
@st.cache_resource
def get_gspread_client():
    try:
        # Mengambil data dari Secrets Streamlit Cloud
        # Pastikan Anda sudah setting di Dashboard Streamlit
        info = dict(st.secrets["gcp_service_account"])
        
        # Fix format private key (\n)
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n')
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Koneksi Gagal. Cek Secrets Anda! Error: {e}")
        return None

# --- 3. SIDEBAR MENU & FILTER ---
with st.sidebar:
    st.title("⚙️ Control Panel")
    
    # Menu Pilihan (WSA / MODOROSO / WAPPR)
    menu = st.radio("Pilih Jenis Proses:", ["WSA", "MODOROSO", "WAPPR"])
    st.markdown("---")
    
    # Filter Bulan Otomatis (Default: Bulan Lalu & Sekarang)
    curr_month = datetime.now().month
    prev_month = curr_month - 1 if curr_month > 1 else 12
    
    selected_months = st.multiselect(
        "📅 Filter Bulan:", 
        options=list(range(1, 13)), 
        default=[prev_month, curr_month],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# --- 4. LOGIKA UTAMA ---
st.title(f"🚀 Dashboard Pemrosesan - {menu}")
st.caption("Sistem Validasi Data & Cek Duplikat Google Sheets Otomatis")

client = get_gspread_client()

if client:
    uploaded_file = st.file_uploader(f"Upload Report {menu} (XLSX/CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # Baca File
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
            
        try:
            with st.spinner(f"Sedang memproses logika {menu}..."):
                
                # --- A. LOGIKA FILTER SPECIFIC (BERDASARKAN MENU) ---
                if menu == "WSA":
                    # Logic: Filter AO/PDA/WSA + CRM Type Create/Migrate
                    # Cek Duplikat: SC Order No
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                    if 'CRM Order Type' in df.columns:
                        df = df[df['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]
                    
                    check_col = 'SC Order No/Track ID/CSRM No'
                    target_sheet = "Sheet1"

                elif menu == "MODOROSO":
                    # Logic: Filter -MO atau -DO
                    # Cek Duplikat: Workorder
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('-MO|-DO', na=False)].copy()
                    
                    check_col = 'Workorder' 
                    target_sheet = "MODOROSO_JAKTIMSEL"

                elif menu == "WAPPR":
                    # Logic: Filter AO/PDA + Status WAPPR
                    # Cek Duplikat: Workorder
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA', na=False)].copy()
                    if 'Status' in df.columns:
                        df = df[df['Status'] == 'WAPPR']
                    
                    check_col = 'Workorder'
                    target_sheet = "Sheet1"

                # --- B. PEMBERSIHAN DATA UMUM ---
                # Bersihkan kolom Date Created (hilangkan .0)
                if 'Date Created' in df.columns:
                    df['Date Created DT'] = pd.to_datetime(df['Date Created'].astype(str).str.replace('.0', '', regex=False), errors='coerce')
                    
                    # Filter Bulan
                    if selected_months:
                        df = df[df['Date Created DT'].dt.month.isin(selected_months)]
                    
                    # Format untuk Display
                    df['Date Created Display'] = df['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M:%S')
                
                # Bersihkan Booking Date
                if 'Booking Date' in df.columns:
                    df['Booking Date'] = df['Booking Date'].astype(str).str.split('.').str[0]
                
                # Bersihkan SC Order (Ambil bagian depan sebelum underscore)
                if 'SC Order No/Track ID/CSRM No' in df.columns:
                     df['SC Order No/Track ID/CSRM No'] = df['SC Order No/Track ID/CSRM No'].apply(lambda x: str(x).split('_')[0])

                # --- C. CEK DUPLIKAT KE GOOGLE SHEETS ---
                spreadsheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT")
                
                # Pilih Sheet Sesuai Menu
                try:
                    if menu == "MODOROSO":
                        worksheet = spreadsheet.worksheet(target_sheet)
                    else:
                        worksheet = spreadsheet.sheet1
                except gspread.WorksheetNotFound:
                    st.error(f"Sheet '{target_sheet}' tidak ditemukan di Google Sheets!")
                    st.stop()
                
                # Ambil Data GDoc
                google_data = worksheet.get_all_records()
                google_df = pd.DataFrame(google_data)

                # Proses Deduplikasi
                if not google_df.empty and check_col in google_df.columns:
                    existing_ids = google_df[check_col].astype(str).unique()
                    df_final = df[~df[check_col].astype(str).isin(existing_ids)].copy()
                else:
                    df_final = df.copy()

                # --- D. TAMPILAN DASHBOARD ---
                m1, m2, m3 = st.columns(3)
                m1.markdown(f'<div class="metric-card">📂 Data Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                m2.markdown(f'<div class="metric-card">✨ Data Unik Baru<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                m3.markdown(f'<div class="metric-card">🔗 Target GDoc<br><h2>{target_sheet}</h2></div>', unsafe_allow_html=True)

                # --- E. PREVIEW & DOWNLOAD ---
                st.subheader(f"📋 Preview Data ({menu})")
                
                # Kolom Output Standar
                cols_target = ['Workzone', 'Date Created Display', 'SC Order No/Track ID/CSRM No', 
                               'Service No.', 'Workorder', 'Customer Name', 'Address', 
                               'Contact Number', 'CRM Order Type']
                
                # Tambah Booking Date jika ada
                if 'Booking Date' in df_final.columns:
                    cols_target.append('Booking Date')
                
                # Filter hanya kolom yang ada
                cols_final = [c for c in cols_target if c in df_final.columns]
                
                # Sortir
                if 'Workzone' in df_final.columns:
                    df_final = df_final.sort_values('Workzone')

                st.dataframe(df_final[cols_final], use_container_width=True)

                # Tombol Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[cols_final].to_excel(writer, index=False)
                
                st.download_button(
                    label=f"📥 Download Hasil {menu}",
                    data=output.getvalue(),
                    file_name=f"Cleaned_{menu}_{datetime.now().strftime('%d%m%Y')}.xlsx"
                )

        except Exception as e:
            st.error(f"Terjadi kesalahan teknis: {e}")
