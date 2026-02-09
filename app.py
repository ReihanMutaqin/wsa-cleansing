import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# Konfigurasi Halaman (Harus di paling atas)
st.set_page_config(
    page_title="WSA Data Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS UNTUK TAMPILAN PROFESIONAL ---
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #007bff;
        color: white;
        border: none;
    }
    .stButton>button:hover {
        background-color: #0056b3;
        color: white;
    }
    div[data-testid="stMetricValue"] {
        font-size: 28px;
        color: #007bff;
    }
    .reportview-container .main .block-container {
        padding-top: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

# --- FUNGSI CORE ---
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
        st.error(f"Gagal koneksi: {e}")
        return None

# --- SIDEBAR ---
with st.sidebar:
    st.image("https://www.gstatic.com/images/branding/product/2x/sheets_2020q4_48dp.png", width=50)
    st.title("Control Panel")
    st.markdown("---")
    selected_months = st.multiselect(
        "📅 Pilih Periode Bulan:",
        options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
        default=[1, 2],
        format_func=lambda x: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][x-1]
    )
    st.markdown("---")
    st.info("Aplikasi ini otomatis memfilter duplikat berdasarkan data di Google Sheets.")

# --- MAIN CONTENT ---
st.title("🚀 WSA Data Cleansing Dashboard")
st.markdown("Sistem otomasi validasi data Order (AO/PDA/WSA) dan sinkronisasi Google Sheets.")

client = get_gspread_client()

if client:
    uploaded_file = st.file_uploader("Seret dan lepas file Report Excel di sini", type=["xlsx"])

    if uploaded_file:
        with st.spinner('Menganalisis data...'):
            data = pd.read_excel(uploaded_file)
            
            if 'SC Order No/Track ID/CSRM No' in data.columns:
                # Proses Data
                data_filtered = data[data['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                data_filtered['Date Created DT'] = pd.to_datetime(data_filtered['Date Created'], errors='coerce')
                
                if selected_months:
                    data_filtered = data_filtered[data_filtered['Date Created DT'].dt.month.isin(selected_months)]
                
                # Cek Duplikat ke Google Sheets
                spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"
                sheet = client.open(spreadsheet_name).sheet1
                google_data = pd.DataFrame(sheet.get_all_records())
                
                if not google_data.empty and 'SC Order No/Track ID/CSRM No' in google_data.columns:
                    existing_orders = google_data['SC Order No/Track ID/CSRM No'].astype(str).unique()
                    unique_data = data_filtered[~data_filtered['SC Order No/Track ID/CSRM No'].astype(str).isin(existing_orders)].copy()
                else:
                    unique_data = data_filtered.copy()

                # --- STATISTIK (METRIC CARDS) ---
                st.markdown("### 📊 Ringkasan Analisis")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Data Upload", f"{len(data)} baris")
                col2.metric("Data Filtered (WSA/AO/PDA)", f"{len(data_filtered)} baris")
                col3.metric("Data Baru (Unik)", f"{len(unique_data)} baris", delta=f"{len(unique_data)} New")

                st.markdown("---")

                # --- DATA PREVIEW ---
                st.subheader("🔍 Preview Data Baru")
                unique_data['Date Created'] = unique_data['Date Created'].astype(str).str.split('.').str[0]
                final_display = unique_data.drop(columns=['Date Created DT'])
                st.dataframe(final_display, use_container_width=True)

                # --- DOWNLOAD SECTION ---
                st.markdown("### 💾 Export Hasil")
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_display.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Download Excel Hasil Cleansing",
                    data=output.getvalue(),
                    file_name="wsa_clean_report.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("Format kolom tidak sesuai. Pastikan ada kolom 'SC Order No/Track ID/CSRM No'.")
else:
    st.warning("Menunggu koneksi ke Database Google...")
