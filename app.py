import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# 1. Konfigurasi Halaman & Tema
st.set_page_config(
    page_title="WSA Pro Dashboard",
    page_icon="⚡",
    layout="wide"
)

# 2. Custom CSS untuk Tampilan Mewah
st.markdown("""
    <style>
    /* Mengubah font dan background */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    .main {
        background-color: #f0f2f6;
    }
    /* Style Card Statistik */
    .metric-container {
        background-color: white;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        border-bottom: 5px solid #ff4b4b;
    }
    /* Tombol Download */
    .stDownloadButton button {
        background-color: #28a745 !important;
        color: white !important;
        border-radius: 10px !important;
        width: 100%;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# 3. Fungsi Koneksi
@st.cache_resource
def get_gspread_client():
    json_file = "project-pengolahan-data-561c0b891db8.json"
    try:
        with open(json_file) as f:
            info = json.load(f)
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Koneksi Gagal: {e}")
        return None

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Filter & Kontrol")
    
    # Filter Bulan
    selected_months = st.multiselect(
        "📅 Pilih Bulan:",
        options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        default=[1, 2],
        format_func=lambda x: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][x-1]
    )
    
    # Tempat penampung filter status (akan diisi setelah file diupload)
    status_filter = st.multiselect("📌 Filter Status:", options=[], help="Upload file dulu untuk memunculkan pilihan status")
    
    st.markdown("---")
    if st.button("🔄 Reset Aplikasi"):
        st.cache_resource.clear()
        st.rerun()

# --- MAIN CONTENT ---
st.title("🚀 WSA FULFILLMENT DASHBOARD")
st.caption("Automated Cleansing & Validation System v2.0")

client = get_gspread_client()

if client:
    uploaded_file = st.file_uploader("Upload file Report Excel (XLSX)", type=["xlsx"])

    if uploaded_file:
        with st.spinner('Sedang memproses database...'):
            df = pd.read_excel(uploaded_file)
            
            if 'SC Order No/Track ID/CSRM No' in df.columns:
                # A. Filter Dasar (AO/PDA/WSA)
                df_clean = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                
                # B. Konversi Tanggal & Filter Bulan
                df_clean['Date Created DT'] = pd.to_datetime(df_clean['Date Created'], errors='coerce')
                if selected_months:
                    df_clean = df_clean[df_clean['Date Created DT'].dt.month.isin(selected_months)]
                
                # C. Fitur Baru: Filter Status
                if 'Status' in df_clean.columns:
                    all_statuses = df_clean['Status'].unique().tolist()
                    # Menampilkan pilihan status di sidebar secara dinamis
                    st.sidebar.markdown("### Status Terdeteksi:")
                    status_choice = st.sidebar.multiselect("Pilih Status:", options=all_statuses, default=all_statuses)
                    df_clean = df_clean[df_clean['Status'].isin(status_choice)]

                # D. Validasi Duplikat ke Google Sheets
                spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"
                sheet = client.open(spreadsheet_name).sheet1
                google_df = pd.DataFrame(sheet.get_all_records())
                
                if not google_df.empty and 'SC Order No/Track ID/CSRM No' in google_df.columns:
                    existing = google_df['SC Order No/Track ID/CSRM No'].astype(str).unique()
                    unique_df = df_clean[~df_clean['SC Order No/Track ID/CSRM No'].astype(str).isin(existing)].copy()
                else:
                    unique_df = df_clean.copy()

                # --- UI: METRIC CARDS ---
                m1, m2, m3 = st.columns(3)
                with m1:
                    st.markdown(f'<div class="metric-container"><h5>Total Row</h5><h2>{len(df)}</h2></div>', unsafe_allow_html=True)
                with m2:
                    st.markdown(f'<div class="metric-container"><h5>WSA Filtered</h5><h2>{len(df_clean)}</h2></div>', unsafe_allow_html=True)
                with m3:
                    st.markdown(f'<div class="metric-container"><h5>Ready to Export</h5><h2>{len(unique_df)}</h2></div>', unsafe_allow_html=True)

                st.markdown("### 📋 Preview Data Unik")
                
                # Bersihkan tampilan kolom Date
                unique_df['Date Created'] = unique_df['Date Created'].astype(str).str.split('.').str[0]
                final_out = unique_df.drop(columns=['Date Created DT'])
                
                st.dataframe(final_out, use_container_width=True, height=400)

                # --- DOWNLOAD ---
                st.markdown("---")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_out.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 DOWNLOAD HASIL CLEANSING (EXCEL)",
                    data=output.getvalue(),
                    file_name="wsa_report_pro.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("Kolom 'SC Order No/Track ID/CSRM No' tidak ditemukan!")
else:
    st.warning("Menghubungkan ke server Google...")
