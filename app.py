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
    # Menggunakan file JSON terbaru Anda
    json_file = "project-pengolahan-data-561c0b891db8.json"
    
    try:
        with open(json_file) as f:
            info = json.load(f)
        
        if 'private_key' in info:
            info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(info, scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Gagal membaca file JSON: {e}")
        return None

# Sidebar untuk Filter
st.sidebar.header("Filter Data")
selected_months = st.sidebar.multiselect(
    "Pilih Bulan (Date Created):",
    options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
    default=[1, 2], # Default seperti di script mentah Anda (Januari & Februari)
    format_func=lambda x: ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"][x-1]
)

try:
    client = get_gspread_client()
    
    if client:
        st.sidebar.success("✅ Koneksi GDoc Aktif")
        
        uploaded_file = st.file_uploader("Upload file Report Excel", type=["xlsx"])

        if uploaded_file:
            with st.spinner('Sedang memproses data...'):
                data = pd.read_excel(uploaded_file)
                
                # 1. Filter AO/PDA/WSA
                if 'SC Order No/Track ID/CSRM No' in data.columns:
                    data_filtered = data[data['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
                    
                    # 2. Konversi ke Datetime untuk Filter Bulan
                    data_filtered['Date Created DT'] = pd.to_datetime(data_filtered['Date Created'], errors='coerce')
                    
                    # 3. Filter Berdasarkan Bulan yang dipilih di Sidebar
                    if selected_months:
                        data_filtered = data_filtered[data_filtered['Date Created DT'].dt.month.isin(selected_months)]
                    
                    # 4. Format Tampilan Tanggal (hilangkan .0 seperti permintaan awal)
                    data_filtered['Date Created'] = data_filtered['Date Created'].astype(str).str.split('.').str[0]
                    
                    # Hapus kolom bantuan datetime agar tidak muncul di hasil download
                    display_data = data_filtered.drop(columns=['Date Created DT'])

                    # 5. Ambil data dari Google Sheets untuk cek duplikat
                    spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"
                    sheet = client.open(spreadsheet_name).sheet1
                    google_data = pd.DataFrame(sheet.get_all_records())

                    # 6. Bandingkan Duplikat
                    if not google_data.empty and 'SC Order No/Track ID/CSRM No' in google_data.columns:
                        existing_orders = google_data['SC Order No/Track ID/CSRM No'].astype(str).unique()
                        unique_data = display_data[~display_data['SC Order No/Track ID/CSRM No'].astype(str).isin(existing_orders)]
                    else:
                        unique_data = display_data

                    st.write(f"### Hasil: {len(unique_data)} Data Baru (Bulan: {', '.join([str(m) for m in selected_months])})")
                    st.dataframe(unique_data)

                    # 7. Tombol Download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        unique_data.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 Download Data Validasi (Excel)",
                        data=output.getvalue(),
                        file_name=f"data_cleansing_bulan_{'_'.join([str(m) for m in selected_months])}.xlsx",
                        mime="application/vnd.ms-excel"
                    )
                else:
                    st.error("Kolom 'SC Order No/Track ID/CSRM No' tidak ditemukan.")

except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
