import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

st.set_page_config(page_title="WSA Data Cleansing", layout="wide")

st.title("📊 WSA Data Cleansing Tool")

# Setup Credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def get_gspread_client():
    # Pastikan nama file sesuai dengan di GitHub Anda
    json_file = "pengolahan.json"
    
    with open(json_file) as f:
        info = json.load(f)
    
    # --- BAGIAN PERBAIKAN JWT SIGNATURE ---
    # Memastikan private_key memiliki format \n yang benar dan tidak ada spasi rusak
    if 'private_key' in info:
        info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
    
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
    return gspread.authorize(creds)

try:
    client = get_gspread_client()
    st.success("✅ Koneksi ke Google Sheets Berhasil!")
    
    uploaded_file = st.file_uploader("Upload file Report Excel", type=["xlsx"])

    if uploaded_file:
        with st.spinner('Sedang memproses data...'):
            # Membaca file excel
            data = pd.read_excel(uploaded_file)
            
            # Filter AO/PDA/WSA
            if 'SC Order No/Track ID/CSRM No' in data.columns:
                data_filtered = data[data['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                
                # Format Tanggal
                if 'Date Created' in data_filtered.columns:
                    data_filtered['Date Created'] = data_filtered['Date Created'].astype(str).str.split('.').str[0]
                
                # Ambil data dari Google Sheets
                spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"
                sheet = client.open(spreadsheet_name).sheet1
                google_data = pd.DataFrame(sheet.get_all_records())

                # Bandingkan Duplikat
                if not google_data.empty and 'SC Order No/Track ID/CSRM No' in google_data.columns:
                    existing_orders = google_data['SC Order No/Track ID/CSRM No'].astype(str).unique()
                    unique_data = data_filtered[~data_filtered['SC Order No/Track ID/CSRM No'].astype(str).isin(existing_orders)]
                else:
                    unique_data = data_filtered

                st.write(f"### Hasil: {len(unique_data)} Data Baru Unik")
                st.dataframe(unique_data)

                # Tombol Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    unique_data.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Download Data Validasi (Excel)",
                    data=output.getvalue(),
                    file_name="data_cleansing_result.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("Kolom 'SC Order No/Track ID/CSRM No' tidak ditemukan di file Excel Anda.")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
    st.info("Jika error JWT menetap, coba hapus file pengolahan.json di GitHub lalu upload ulang file aslinya.")
