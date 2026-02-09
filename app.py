import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import io
import json

# 1. SET THEME LANGSUNG DI KODE
st.set_page_config(page_title="WSA Dashboard", layout="wide")

# CSS untuk mengubah warna Dashboard agar tidak "Polosan"
st.markdown("""
    <style>
    .stApp { background-color: #f4f7f9; }
    div[data-testid="stMetricValue"] { color: #007bff; font-weight: bold; }
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    .st-emotion-cache-1av53o2 { background-color: #ffffff; padding: 20px; border-radius: 10px; box-shadow: 0px 4px 12px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

@st.cache_resource
def get_client():
    # Gunakan file JSON terbaru Anda
    json_file = "project-pengolahan-data-561c0b891db8.json"
    with open(json_file) as f:
        info = json.load(f)
    info['private_key'] = info['private_key'].replace('\\n', '\n').strip()
    creds = ServiceAccountCredentials.from_json_keyfile_dict(info, ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)

# --- UI SIDEBAR ---
with st.sidebar:
    st.title("🛡️ WSA Control")
    st.subheader("Filter Periode")
    selected_months = st.multiselect(
        "Pilih Bulan:", options=[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12], default=[1, 2],
        format_func=lambda x: ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"][x-1]
    )

st.title("🚀 WSA Fulfillment Dashboard")
st.markdown("Otomasi pembersihan data & validasi Google Sheets.")

try:
    client = get_client()
    uploaded_file = st.file_uploader("Upload Report Excel", type=["xlsx"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        
        # --- PROSES FILTER DASAR ---
        # 1. Filter WSA/AO/PDA
        df = df[df['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)].copy()
        
        # 2. Filter Bulan
        df['Date Created DT'] = pd.to_datetime(df['Date Created'], errors='coerce')
        if selected_months:
            df = df[df['Date Created DT'].dt.month.isin(selected_months)]

        # --- FITUR FILTER STATUS ---
        if 'Status' in df.columns:
            st.sidebar.subheader("Filter Status")
            all_statuses = df['Status'].unique().tolist()
            selected_status = st.sidebar.multiselect("Pilih Status Data:", options=all_statuses, default=all_statuses)
            df = df[df['Status'].isin(selected_status)]

        # --- CEK DUPLIKAT GDOC ---
        sheet = client.open("Salinan dari NEW GDOC WSA FULFILLMENT").sheet1
        google_df = pd.DataFrame(sheet.get_all_records())
        
        if not google_df.empty and 'SC Order No/Track ID/CSRM No' in google_df.columns:
            existing = google_df['SC Order No/Track ID/CSRM No'].astype(str).unique()
            df_final = df[~df['SC Order No/Track ID/CSRM No'].astype(str).isin(existing)].copy()
        else:
            df_final = df.copy()

        # --- UI DISPLAY ---
        col1, col2, col3 = st.columns(3)
        col1.metric("Data Mentah", len(df))
        col2.metric("Data Unik Baru", len(df_final))
        col3.success("Koneksi GDoc Aman!")

        st.subheader("📋 Preview Data Hasil Filter")
        df_final['Date Created'] = df_final['Date Created'].astype(str).str.split('.').str[0]
        st.dataframe(df_final.drop(columns=['Date Created DT']), use_container_width=True)

        # --- DOWNLOAD ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.drop(columns=['Date Created DT']).to_excel(writer, index=False)
        
        st.download_button("📥 Download Excel Hasil", output.getvalue(), "wsa_clean_report.xlsx", "application/vnd.ms-excel")

except Exception as e:
    st.error(f"Error: {e}")
