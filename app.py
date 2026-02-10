import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json
from datetime import datetime

# --- CONFIG PAGE ---
st.set_page_config(page_title="WSA Multi-Tool", layout="wide")

st.markdown("""
    <style>
    .stApp { background-color: #0e1117; color: #ffffff; }
    [data-testid="stSidebar"] { background-color: #1a1c24; }
    .metric-card {
        background-color: #1e2129;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #00d4ff;
        box-shadow: 0 4px 6px rgba(0,0,0,0.3);
    }
    .status-box { padding: 10px; border-radius: 5px; margin-bottom: 20px; font-weight: bold; }
    .success-box { background-color: #1c4f2e; color: #4caf50; border: 1px solid #4caf50; }
    .error-box { background-color: #4f1c1c; color: #ff4b4b; border: 1px solid #ff4b4b; }
    h1, h2, h3 { color: #00d4ff !important; }
    .stDownloadButton button {
        background: linear-gradient(45deg, #00d4ff, #008fb3) !important;
        color: #000000 !important;
        font-weight: bold;
        width: 100%;
        border-radius: 8px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- KONEKSI GOOGLE SHEETS ---
@st.cache_resource
def get_gspread_client():
    try:
        # Mengambil data dari Secrets Streamlit Cloud
        info = dict(st.secrets["gcp_service_account"])
        
        # Perbaikan format private key (mengembalikan karakter baris baru)
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n')
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
        return gspread.authorize(creds)
    except Exception as e:
        return None

# --- SIDEBAR MENU ---
with st.sidebar:
    st.title("⚙️ Control Panel")
    menu = st.radio("Pilih Operasi:", ["WSA", "MODOROSO", "WAPPR"])
    st.markdown("---")
    
    curr_month = datetime.now().month
    prev_month = curr_month - 1 if curr_month > 1 else 12
    selected_months = st.multiselect(
        "📅 Filter Bulan:", 
        options=list(range(1, 13)), 
        default=[prev_month, curr_month],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# --- MAIN LOGIC ---
st.title(f"🚀 Dashboard {menu}")

# Tentukan Nama Sheet Tujuan
if menu == "MODOROSO":
    target_sheet_name = "MODOROSO_JAKTIMSEL"
else:
    target_sheet_name = "Sheet1"

client = get_gspread_client()

# --- CEK KONEKSI (VISUAL) ---
if client:
    try:
        # Coba buka Spreadsheet dan Sheet spesifik
        sh = client.open("Salinan dari NEW GDOC WSA FULFILLMENT")
        ws = sh.worksheet(target_sheet_name)
        
        st.markdown(f"""
        <div class="status-box success-box">
            ✅ SISTEM ONLINE | Terhubung ke: {target_sheet_name}
        </div>
        """, unsafe_allow_html=True)
        connection_status = True

    except Exception as e:
        st.markdown(f"""
        <div class="status-box error-box">
            ❌ GAGAL AKSES SHEET | Cek nama sheet: {target_sheet_name}<br>
            <small>Error: {e}</small>
        </div>
        """, unsafe_allow_html=True)
        connection_status = False
else:
    st.error("Gagal membaca Secrets. Pastikan Anda sudah update Secrets di Streamlit Cloud.")
    connection_status = False

# --- UPLOAD FILE ---
if connection_status:
    uploaded_file = st.file_uploader(f"Upload Data {menu} (XLSX/CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
            
        try:
            with st.spinner(f"Memproses data {menu}..."):
                
                # A. LOGIKA FILTERING
                if menu == "WSA":
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                    if 'CRM Order Type' in df.columns:
                        df = df[df['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]
                    check_col = 'SC Order No/Track ID/CSRM No'

                elif menu == "MODOROSO":
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('-MO|-DO', na=False)].copy()
                    check_col = 'Workorder' 

                elif menu == "WAPPR":
                    df = df_raw[df_raw['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA', na=False)].copy()
                    if 'Status' in df.columns:
                        df = df[df['Status'] == 'WAPPR']
                    check_col = 'Workorder'

                # B. CLEANING DATA
                if 'Date Created' in df.columns:
                    df['Date Created DT'] = pd.to_datetime(df['Date Created'].astype(str).str.replace('.0', '', regex=False), errors='coerce')
                    if selected_months:
                        df = df[df['Date Created DT'].dt.month.isin(selected_months)]
                    df['Date Created Display'] = df['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M:%S')
                
                if 'Booking Date' in df.columns:
                    df['Booking Date'] = df['Booking Date'].astype(str).str.split('.').str[0]
                
                if 'SC Order No/Track ID/CSRM No' in df.columns:
                     df['SC Order No/Track ID/CSRM No'] = df['SC Order No/Track ID/CSRM No'].apply(lambda x: str(x).split('_')[0])

                # C. CEK DUPLIKAT (Pakai ws yang sudah dibuka di atas)
                google_data = ws.get_all_records()
                google_df = pd.DataFrame(google_data)

                if not google_df.empty and check_col in google_df.columns:
                    existing_ids = google_df[check_col].astype(str).unique()
                    df_final = df[~df[check_col].astype(str).isin(existing_ids)].copy()
                else:
                    df_final = df.copy()

                # D. DISPLAY HASIL
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="metric-card">📂 Data Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="metric-card">✨ Data Unik<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="metric-card">🔗 Sheet Target<br><h2>{target_sheet_name}</h2></div>', unsafe_allow_html=True)

                st.subheader("📋 Preview Data")
                cols_target = ['Workzone', 'Date Created Display', 'SC Order No/Track ID/CSRM No', 
                               'Service No.', 'Workorder', 'Customer Name', 'Address', 
                               'Contact Number', 'CRM Order Type']
                if 'Booking Date' in df_final.columns: cols_target.append('Booking Date')
                
                cols_final = [c for c in cols_target if c in df_final.columns]
                if 'Workzone' in df_final.columns: df_final = df_final.sort_values('Workzone')

                st.dataframe(df_final[cols_final], use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[cols_final].to_excel(writer, index=False)
                
                st.download_button(
                    label=f"📥 Download Hasil {menu}",
                    data=output.getvalue(),
                    file_name=f"Cleaned_{menu}_{datetime.now().strftime('%d%m%Y')}.xlsx"
                )

        except Exception as e:
            st.error(f"Error Processing: {e}")
