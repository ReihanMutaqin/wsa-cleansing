import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io

st.set_page_config(page_title="WSA Data Cleansing", layout="wide")

st.title("📊 WSA Data Cleansing Tool")
st.info("Upload file report.xlsx untuk memfilter data baru yang belum ada di Google Sheets.")

# Setup Credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Fungsi untuk memuat koneksi ke Google Sheets
@st.cache_resource
def get_gspread_client():
    # Pastikan nama file JSON sesuai dengan yang ada di GitHub Anda
    creds = ServiceAccountCredentials.from_json_keyfile_name("project-pengolahan-data-62bff7107e8e.json", scope)
    return gspread.authorize(creds)

try:
    client = get_gspread_client()
    
    uploaded_file = st.file_uploader("Upload file Report Excel", type=["xlsx"])

    if uploaded_file:
        with st.spinner('Memproses data...'):
            # 1. Baca data input
            data = pd.read_excel(uploaded_file)
            
            # 2. Filter AO/PDA/WSA
            data_filtered = data[data['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
            
            # 3. Format Tanggal agar bersih
            data_filtered['Date Created'] = data_filtered['Date Created'].astype(str).str.split('.').str[0]
            
            # 4. Ambil data dari Google Sheets
            spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"
            sheet = client.open(spreadsheet_name).sheet1
            google_data = pd.DataFrame(sheet.get_all_records())

            # 5. Cari data yang belum ada di GDoc
            if 'SC Order No/Track ID/CSRM No' in google_data.columns:
                existing_orders = google_data['SC Order No/Track ID/CSRM No'].astype(str).tolist()
                unique_data = data_filtered[~data_filtered['SC Order No/Track ID/CSRM No'].astype(str).isin(existing_orders)]
            else:
                unique_data = data_filtered

            # 6. Tampilkan Hasil
            st.success(f"Ditemukan {len(unique_data)} data baru yang siap divalidasi.")
            st.dataframe(unique_data)

            # 7. Tombol Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                unique_data.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Download Data Validasi (Excel)",
                data=output.getvalue(),
                file_name="data_validasi_hasil_clean.xlsx",
                mime="application/vnd.ms-excel"
            )

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")