import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from datetime import datetime

# ==========================================
# 1. KONFIGURASI HALAMAN & CSS
# ==========================================
st.set_page_config(page_title="WSA Multi-Tool Pro", layout="wide")

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

# ==========================================
# 2. KONEKSI GOOGLE SHEETS
# ==========================================
@st.cache_resource
def get_gspread_client():
    try:
        info = dict(st.secrets["gcp_service_account"])
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n')
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds)
    except Exception as e:
        return None

# ==========================================
# 3. FUNGSI LOGIKA TERPISAH (MODULAR)
# ==========================================

# --- Helper: Bersihkan Format Standar ---
def clean_common_data(df):
    # Bersihkan Workorder (hapus .0)
    if 'Workorder' in df.columns:
        df['Workorder'] = df['Workorder'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    
    # Bersihkan Booking Date
    if 'Booking Date' in df.columns:
        df['Booking Date'] = df['Booking Date'].astype(str).str.split('.').str[0]
        
    return df

# --- LOGIKA 1: WSA (VALIDATION) ---
# Menggunakan Logika Awal/Original (Case Sensitive, Create/Migrate, Fill Contact)
def proses_wsa(df):
    col_sc = 'SC Order No/Track ID/CSRM No'
    
    # 1. Filter Regex Strict (Case Sensitive)
    df = df[df[col_sc].astype(str).str.contains('AO|PDA|WSA', na=False)]
    
    # 2. Filter Type: CREATE / MIGRATE
    if 'CRM Order Type' in df.columns:
        df = df[df['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]
    
    # 3. Fitur Khusus: Isi Contact Number Kosong
    if 'Contact Number' in df.columns and 'Customer Name' in df.columns:
        c_map = df.loc[df['Contact Number'].notna() & (df['Contact Number'] != ''), ['Customer Name', 'Contact Number']].drop_duplicates('Customer Name')
        c_dict = dict(zip(c_map['Customer Name'], c_map['Contact Number']))
        
        def fill_contact(row):
            val = str(row['Contact Number'])
            if pd.isna(row['Contact Number']) or val.strip() == '' or val.lower() == 'nan':
                return c_dict.get(row['Customer Name'], row['Contact Number'])
            return row['Contact Number']
        
        df['Contact Number'] = df.apply(fill_contact, axis=1)
    
    # Return: Dataframe hasil filter & Kolom Kunci Cek Duplikat
    return df, col_sc

# --- LOGIKA 2: MODOROSO ---
# Menggunakan Logika Baru (Cari -MO/-DO di SC Order, Auto Rename CRM Type)
def proses_modoroso(df):
    col_sc = 'SC Order No/Track ID/CSRM No'
    
    # 1. Filter: Cari "-MO" atau "-DO" (Case Insensitive)
    df = df[df[col_sc].astype(str).str.contains(r'-MO|-DO', na=False, case=False)]
    
    # 2. Rename Type Otomatis
    if 'CRM Order Type' in df.columns:
        def detect_mo_do(val):
            s = str(val).upper()
            if '-MO' in s: return 'MO'
            if '-DO' in s: return 'DO'
            return 'MO' # Default fallback
        
        df['CRM Order Type'] = df[col_sc].apply(detect_mo_do)
    
    # Return: Dataframe hasil filter & Kolom Kunci Cek Duplikat (Workorder)
    return df, 'Workorder'

# --- LOGIKA 3: WAPPR ---
# Menggunakan Logika Standar WAPPR
def proses_wappr(df):
    col_sc = 'SC Order No/Track ID/CSRM No'
    
    # 1. Filter Regex AO/PDA
    df = df[df[col_sc].astype(str).str.contains('AO|PDA', na=False)]
    
    # 2. Filter Status WAPPR
    if 'Status' in df.columns:
        df = df[df['Status'].astype(str).str.strip().str.upper() == 'WAPPR']
    
    # Return: Dataframe hasil filter & Kolom Kunci Cek Duplikat (Workorder)
    return df, 'Workorder'

# ==========================================
# 4. UI SIDEBAR & NAVIGASI
# ==========================================
with st.sidebar:
    st.title("⚙️ Control Panel")
    menu = st.radio("Pilih Operasi:", ["WSA (Validation)", "MODOROSO", "WAPPR"])
    st.markdown("---")
    
    curr_month = datetime.now().month
    prev_month = curr_month - 1 if curr_month > 1 else 12
    
    selected_months = st.multiselect(
        "📅 Filter Bulan Data:", 
        options=list(range(1, 13)), 
        default=[prev_month, curr_month],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# ==========================================
# 5. EKSEKUSI UTAMA (MAIN APP)
# ==========================================
st.title(f"🚀 Dashboard {menu}")

client = get_gspread_client()
ws = None

# A. SETUP SHEET TARGET
if client:
    try:
        sh = client.open("Salinan dari NEW GDOC WSA FULFILLMENT")
        
        if menu == "MODOROSO":
            target_sheet_name = "MODOROSO_JAKTIMSEL"
            try:
                ws = sh.worksheet(target_sheet_name)
            except:
                st.error(f"Sheet '{target_sheet_name}' tidak ditemukan! Cek nama tab GDoc.")
                ws = None
        else:
            # WSA & WAPPR: Ambil Sheet Index 0
            ws = sh.get_worksheet(0)
            target_sheet_name = ws.title

        if ws:
            st.markdown(f'<div class="status-box success-box">✅ SISTEM ONLINE | Sheet: {target_sheet_name}</div>', unsafe_allow_html=True)
            connection_status = True
        else:
            connection_status = False

    except Exception as e:
        st.markdown(f'<div class="status-box error-box">❌ GAGAL AKSES SHEET | {e}</div>', unsafe_allow_html=True)
        connection_status = False
else:
    st.error("Gagal membaca Secrets API.")
    connection_status = False

# B. UPLOAD & PROSES FILE
if connection_status and ws:
    uploaded_file = st.file_uploader(f"Upload Data {menu} (XLSX/CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        df_raw = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            
        try:
            with st.spinner(f"Memproses {menu}..."):
                
                # 1. Bersihkan Data Umum
                df = clean_common_data(df_raw.copy())

                # 2. Panggil Fungsi Logika Berdasarkan Menu (DIPISAH DI SINI)
                if menu == "WSA (Validation)":
                    df_filtered, check_col = proses_wsa(df)
                elif menu == "MODOROSO":
                    df_filtered, check_col = proses_modoroso(df)
                elif menu == "WAPPR":
                    df_filtered, check_col = proses_wappr(df)

                # 3. Filter Bulan (Common Logic)
                if 'Date Created' in df_filtered.columns:
                    df_filtered['Date Created DT'] = pd.to_datetime(df_filtered['Date Created'].astype(str).str.replace(r'\.0$', '', regex=True), errors='coerce')
                    
                    data_count_before = len(df_filtered)
                    if selected_months:
                        df_filtered = df_filtered[df_filtered['Date Created DT'].dt.month.isin(selected_months)]
                    
                    # Warning jika data habis karena bulan
                    if data_count_before > 0 and len(df_filtered) == 0:
                        st.warning(f"⚠️ {data_count_before} data ditemukan, tapi hilang karena filter bulan tidak sesuai.")

                    df_filtered['Date Created Display'] = df_filtered['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M')
                    df_filtered['Date Created'] = df_filtered['Date Created Display']

                # 4. Cek Duplikat ke GDoc (Common Logic)
                google_data = ws.get_all_records()
                google_df = pd.DataFrame(google_data)
                
                if not google_df.empty and check_col in google_df.columns:
                    existing_ids = google_df[check_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique()
                    
                    # Khusus WSA, kita samakan format SC Order sebelum cek
                    col_sc = 'SC Order No/Track ID/CSRM No'
                    if col_sc in df_filtered.columns:
                        df_filtered[col_sc] = df_filtered[col_sc].astype(str).apply(lambda x: x.split('_')[0])

                    df_final = df_filtered[~df_filtered[check_col].astype(str).str.strip().isin(existing_ids)].copy()
                else:
                    col_sc = 'SC Order No/Track ID/CSRM No'
                    if col_sc in df_filtered.columns:
                        df_filtered[col_sc] = df_filtered[col_sc].astype(str).apply(lambda x: x.split('_')[0])
                    df_final = df_filtered.copy()

                # 5. Tampilkan Hasil (Urutan Kolom Sama)
                target_order = ['Date Created', 'Workorder', 'SC Order No/Track ID/CSRM No', 
                                'Service No.', 'CRM Order Type', 'Status', 'Address', 
                                'Customer Name', 'Workzone', 'Booking Date', 'Contact Number']
                
                cols_final = [c for c in target_order if c in df_final.columns]
                
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="metric-card">📂 Data Filtered<br><h2>{len(df_filtered)}</h2></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="metric-card">✨ Data Unik<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="metric-card">🔗 Validasi By<br><h5>{check_col}</h5></div>', unsafe_allow_html=True)

                st.subheader("📋 Preview Data Unik")
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
            st.error(f"Terjadi kesalahan: {e}")
