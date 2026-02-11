import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from datetime import datetime

# --- 1. CONFIG PAGE ---
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

# --- 2. KONEKSI GOOGLE SHEETS ---
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

# --- 3. SIDEBAR MENU ---
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

# --- 4. MAIN LOGIC ---
st.title(f"🚀 Dashboard {menu}")

client = get_gspread_client()
ws = None

if client:
    try:
        sh = client.open("Salinan dari NEW GDOC WSA FULFILLMENT")
        
        if menu == "MODOROSO":
            target_sheet_name = "MODOROSO_JAKTIMSEL"
            try:
                ws = sh.worksheet(target_sheet_name)
            except:
                st.error(f"Sheet '{target_sheet_name}' tidak ditemukan! Pastikan nama tab di GDoc sesuai.")
                ws = None
        else:
            # WSA dan WAPPR: AMBIL SHEET PERTAMA (INDEX 0)
            ws = sh.get_worksheet(0)
            target_sheet_name = ws.title

        if ws:
            st.markdown(f'<div class="status-box success-box">✅ SISTEM ONLINE | Terhubung ke Sheet: {target_sheet_name}</div>', unsafe_allow_html=True)
            connection_status = True
        else:
            connection_status = False

    except Exception as e:
        st.markdown(f'<div class="status-box error-box">❌ GAGAL AKSES SHEET | {e}</div>', unsafe_allow_html=True)
        connection_status = False
else:
    st.error("Gagal membaca Secrets API.")
    connection_status = False

# --- 5. PROSES DATA ---
if connection_status and ws:
    uploaded_file = st.file_uploader(f"Upload Data {menu} (XLSX/CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # Baca File
        df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            
        try:
            with st.spinner(f"Memproses {menu}..."):
                # STANDARISASI KOLOM
                col_sc = 'SC Order No/Track ID/CSRM No'
                
                # --- A. LOGIKA FILTERING ---
                
                # === 1. WSA (VALIDATION) ===
                if menu == "WSA (Validation)":
                    # Filter Regex
                    df = df[df[col_sc].astype(str).str.contains('AO|PDA|WSA', na=False, case=False)]
                    # Filter Type: CREATE / MIGRATE
                    if 'CRM Order Type' in df.columns:
                        df = df[df['CRM Order Type'].astype(str).str.strip().str.upper().isin(['CREATE', 'MIGRATE'])]
                    
                    # Fill Contact
                    if 'Contact Number' in df.columns and 'Customer Name' in df.columns:
                        c_map = df.loc[df['Contact Number'].notna() & (df['Contact Number'] != ''), ['Customer Name', 'Contact Number']].drop_duplicates('Customer Name')
                        c_dict = dict(zip(c_map['Customer Name'], c_map['Contact Number']))
                        df['Contact Number'] = df.apply(lambda r: c_dict.get(r['Customer Name'], r['Contact Number']) if pd.isna(r['Contact Number']) or str(r['Contact Number']).strip() == '' else r['Contact Number'], axis=1)
                    check_col = col_sc
                
                # === 2. MODOROSO (UPDATED) ===
                elif menu == "MODOROSO":
                    # Filter Regex: MO atau DO
                    df = df[df[col_sc].astype(str).str.contains('MO|DO', na=False, case=False)]
                    
                    # NEW: Filter Type hanya MODIFY dan DISCONNECT
                    if 'CRM Order Type' in df.columns:
                        # Menggunakan Upper() agar 'Modify' dan 'MODIFY' sama-sama terbaca
                        # Menambahkan 'DISCONECT' (typo) dan 'DISCONNECT' (baku) untuk keamanan
                        allowed_types = ['MODIFY', 'DISCONNECT', 'DISCONECT']
                        df = df[df['CRM Order Type'].astype(str).str.strip().str.upper().isin(allowed_types)]
                        
                    check_col = 'Workorder'

                # === 3. WAPPR ===
                elif menu == "WAPPR":
                    # Filter Regex
                    df = df[df[col_sc].astype(str).str.contains('AO|PDA', na=False, case=False)]
                    # Filter Status: WAPPR
                    if 'Status' in df.columns:
                        df = df[df['Status'].astype(str).str.strip().str.upper() == 'WAPPR']
                    check_col = 'Workorder'

                # --- B. FILTER BULAN ---
                if 'Date Created' in df.columns:
                    # Hapus .0 -> ubah ke datetime
                    df['Date Created DT'] = pd.to_datetime(df['Date Created'].astype(str).str.replace(r'\.0$', '', regex=True), errors='coerce')
                    
                    data_before_month = len(df)
                    if selected_months:
                        df = df[df['Date Created DT'].dt.month.isin(selected_months)]
                    
                    if data_before_month > 0 and len(df) == 0:
                        st.warning(f"⚠️ PERHATIAN: Ada {data_before_month} data yang cocok, TAPI hilang karena filternya beda bulan. Cek menu 'Filter Bulan' di kiri!")

                    df['Date Created Display'] = df['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M')
                    df['Date Created'] = df['Date Created Display']

                # --- C. CLEANING WORKORDER ---
                if 'Workorder' in df.columns:
                    df['Workorder'] = df['Workorder'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

                # --- D. CEK DUPLIKAT ---
                google_data = ws.get_all_records()
                google_df = pd.DataFrame(google_data)
                
                if not google_df.empty and check_col in google_df.columns:
                    existing_ids = google_df[check_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().unique()
                    
                    # Split SC Order visual hanya setelah filter selesai
                    if col_sc in df.columns:
                        df[col_sc] = df[col_sc].astype(str).apply(lambda x: x.split('_')[0])
                    
                    df_final = df[~df[check_col].astype(str).str.strip().isin(existing_ids)].copy()
                else:
                    if col_sc in df.columns:
                        df[col_sc] = df[col_sc].astype(str).apply(lambda x: x.split('_')[0])
                    df_final = df.copy()

                # --- E. OUTPUT DISPLAY ---
                target_order = ['Date Created', 'Workorder', col_sc, 'Service No.', 'CRM Order Type', 'Status', 'Address', 'Customer Name', 'Workzone', 'Booking Date', 'Contact Number']
                cols_final = [c for c in target_order if c in df_final.columns]
                
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="metric-card">📂 Data Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="metric-card">✨ Data Unik Baru<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="metric-card">🔗 Validasi By<br><h5>{check_col}</h5></div>', unsafe_allow_html=True)

                st.subheader("📋 Preview Data Unik")
                if 'Workzone' in df_final.columns: df_final = df_final.sort_values('Workzone')
                st.dataframe(df_final[cols_final], use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final[cols_final].to_excel(writer, index=False)
                st.download_button(label=f"📥 Download Hasil {menu}", data=output.getvalue(), file_name=f"Cleaned_{menu}_{datetime.now().strftime('%d%m%Y')}.xlsx")

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
