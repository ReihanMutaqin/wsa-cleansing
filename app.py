import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
from datetime import datetime

# --- 1. CONFIG PAGE ---
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

# --- 2. KONEKSI GOOGLE SHEETS ---
@st.cache_resource
def get_gspread_client():
    try:
        # Mengambil dari Secrets Streamlit
        info = dict(st.secrets["gcp_service_account"])
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n')
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
        return gspread.authorize(creds)
    except Exception as e:
        return None

# --- 3. SIDEBAR MENU ---
with st.sidebar:
    st.title("⚙️ Control Panel")
    # Menu disesuaikan dengan nama script asli
    menu = st.radio("Pilih Operasi:", ["WSA (Validation)", "MODOROSO", "WAPPR"])
    st.markdown("---")
    
    # Filter Bulan (User Control)
    curr_month = datetime.now().month
    prev_month = curr_month - 1 if curr_month > 1 else 12
    selected_months = st.multiselect(
        "📅 Filter Bulan:", 
        options=list(range(1, 13)), 
        default=[prev_month, curr_month],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

# --- 4. MAIN LOGIC ---
st.title(f"🚀 Dashboard {menu}")

client = get_gspread_client()
ws = None
target_sheet_name = ""

# --- A. KONEKSI & PILIH SHEET ---
if client:
    try:
        sh = client.open("Salinan dari NEW GDOC WSA FULFILLMENT")
        
        if menu == "MODOROSO":
            # Script asli: worksheet("MODOROSO_JAKTIMSEL")
            target_sheet_name = "MODOROSO_JAKTIMSEL"
            ws = sh.worksheet(target_sheet_name)
        else:
            # Script asli WSA & WAPPR: sheet1 (index 0)
            ws = sh.get_worksheet(0)
            target_sheet_name = ws.title

        st.markdown(f"""
        <div class="status-box success-box">
            ✅ SISTEM ONLINE | Terhubung ke Sheet: {target_sheet_name}
        </div>
        """, unsafe_allow_html=True)
        connection_status = True

    except Exception as e:
        st.markdown(f"""
        <div class="status-box error-box">
            ❌ GAGAL AKSES SHEET | {e}
        </div>
        """, unsafe_allow_html=True)
        connection_status = False
else:
    st.error("Kunci API (Secrets) bermasalah/kosong.")
    connection_status = False


# --- B. PROSES DATA ---
if connection_status and ws:
    uploaded_file = st.file_uploader(f"Upload Data {menu} (XLSX/CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        # BACA FILE
        if uploaded_file.name.lower().endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)
            
        try:
            with st.spinner(f"Sedang memproses logika {menu}..."):
                
                df = df_raw.copy()

                # --- STEP 1: PRE-PROCESSING UMUM (DATE & SC ORDER) ---
                if 'Date Created' in df.columns:
                    # Hapus .0 dan ubah ke datetime
                    df['Date Created DT'] = pd.to_datetime(df['Date Created'].astype(str).str.replace('.0', '', regex=False), errors='coerce')
                    
                    # Filter Bulan
                    if selected_months:
                        df = df[df['Date Created DT'].dt.month.isin(selected_months)]
                    
                    # Format Display sesuai script asli: '%d/%m/%Y %H:%M:%S'
                    df['Date Created'] = df['Date Created DT'].dt.strftime('%d/%m/%Y %H:%M:%S')

                if 'Booking Date' in df.columns:
                    df['Booking Date'] = df['Booking Date'].astype(str).str.split('.').str[0]
                
                
                # --- STEP 2: LOGIKA SPESIFIK (FILTER & LOGIC) ---
                
                # === MENU 1: WSA (Validation) ===
                if menu == "WSA (Validation)":
                    # 1. Filter Regex: AO | PDA | WSA
                    df = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)]
                    
                    # 2. Filter CRM Order Type: CREATE | MIGRATE
                    if 'CRM Order Type' in df.columns:
                        df = df[df['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]
                    
                    # 3. Fitur Khusus WSA: Isi Contact Number Kosong (Logic Asli)
                    if 'Contact Number' in df.columns and 'Customer Name' in df.columns:
                        contact_map = df.loc[df['Contact Number'].notna() & (df['Contact Number'] != ''), 
                                           ['Customer Name', 'Contact Number']].drop_duplicates('Customer Name')
                        contact_dict = dict(zip(contact_map['Customer Name'], contact_map['Contact Number']))
                        
                        def fill_contact(row):
                            if pd.isna(row['Contact Number']) or str(row['Contact Number']).strip() == '':
                                return contact_dict.get(row['Customer Name'], row['Contact Number'])
                            return row['Contact Number']
                        
                        df['Contact Number'] = df.apply(fill_contact, axis=1)
                    
                    # 4. Kolom Cek Duplikat: SC Order No (Logic Asli)
                    check_col = 'SC Order No/Track ID/CSRM No'
                    
                    # 5. Kolom Output (Ada Booking Date)
                    output_cols_list = ['Date Created', 'Workorder','SC Order No/Track ID/CSRM No', 
                                        'Service No.', 'CRM Order Type', 'Status', 'Address', 
                                        'Customer Name', 'Workzone', 'Booking Date','Contact Number']
                
                # === MENU 2: MODOROSO ===
                elif menu == "MODOROSO":
                    # 1. Filter Regex: -MO | -DO
                    df = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('-MO|-DO', na=False)]
                    
                    # 2. Kolom Cek Duplikat: Workorder (Logic Asli)
                    check_col = 'Workorder'
                    
                    # 3. Kolom Output (Logic Asli TIDAK ADA Booking Date)
                    output_cols_list = ['Date Created', 'Workorder','SC Order No/Track ID/CSRM No', 
                                        'Service No.', 'CRM Order Type', 'Status', 'Address', 
                                        'Customer Name', 'Workzone', 'Booking Date','Contact Number']

                # === MENU 3: WAPPR ===
                elif menu == "WAPPR":
                    # 1. Filter Regex: AO | PDA
                    df = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA', na=False)]
                    
                    # 2. Filter Status: WAPPR
                    if 'Status' in df.columns:
                        df = df[df['Status'] == 'WAPPR']
                    
                    # 3. Kolom Cek Duplikat: Workorder (Logic Asli)
                    check_col = 'Workorder'
                    
                    # 4. Kolom Output (Ada Booking Date)
                    output_cols_list = ['Date Created', 'Workorder','SC Order No/Track ID/CSRM No', 
                                        'Service No.', 'CRM Order Type', 'Status', 'Address', 
                                        'Customer Name', 'Workzone', 'Booking Date','Contact Number']

                # --- STEP 3: RAPIKAN SC ORDER (Split Underscore) ---
                # Dilakukan setelah filter regex agar filter tetap akurat
                if 'SC Order No/Track ID/CSRM No' in df.columns:
                    df['SC Order No/Track ID/CSRM No'] = df['SC Order No/Track ID/CSRM No'].apply(lambda x: str(x).split('_')[0])

                # --- STEP 4: CEK DUPLIKAT KE GOOGLE SHEETS ---
                google_data = ws.get_all_records()
                google_df = pd.DataFrame(google_data)

                if not google_df.empty and check_col in google_df.columns:
                    # Ambil list ID yang sudah ada di GDoc
                    existing_ids = google_df[check_col].astype(str).unique()
                    
                    # Filter data Excel yang TIDAK ada di GDoc
                    df_final = df[~df[check_col].astype(str).isin(existing_ids)].copy()
                else:
                    df_final = df.copy()

                # --- STEP 5: TAMPILAN DASHBOARD ---
                c1, c2, c3 = st.columns(3)
                c1.markdown(f'<div class="metric-card">📂 Data Filtered<br><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                c2.markdown(f'<div class="metric-card">✨ Data Unik (Ready)<br><h2>{len(df_final)}</h2></div>', unsafe_allow_html=True)
                c3.markdown(f'<div class="metric-card">🔗 Validasi By<br><h5>{check_col}</h5></div>', unsafe_allow_html=True)

                # --- STEP 6: PREVIEW & DOWNLOAD ---
                st.subheader("📋 Preview Data Unik")
                
                # Filter hanya kolom yang tersedia di dataframe
                cols_final = [c for c in output_cols_list if c in df_final.columns]
                
                # Sortir berdasarkan Workzone
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





