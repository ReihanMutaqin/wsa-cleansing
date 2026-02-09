import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Setup credentials untuk akses Google Sheets API
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(r"C:\Users\BAYU\Downloads\project-pengolahan-data-62bff7107e8e.json", scope)
client = gspread.authorize(creds)

# Path ke file Excel input dan output
input_file = 'report (1).xlsx'  # Ganti dengan nama file Excel Anda
output_file = 'data validasi ODP PSB belum ada di gdoc.xlsx'  # Nama file untuk menyimpan hasil

# Step 1: Membaca file Excel
data = pd.read_excel(input_file)

# Filter data yang hanya mengandung 'WSA' pada kolom 'SC Order No/Track ID/CSRM No'
data = data[data['SC Order No/Track ID/CSRM No'].astype(str).str.contains('AO|PDA|WSA', na=False)]

# Menyiapkan DataFrame untuk data yang akan disimpan
processed_data = pd.DataFrame()

# 1. Menghilangkan '.0' pada kolom Date Created dan mengubahnya menjadi datetime
data['Date Created'] = data['Date Created'].astype(str).str.split('.').str[0]
processed_data['Date Created'] = pd.to_datetime(data['Date Created'], format='%Y-%m-%d %H:%M:%S', errors='coerce')

# 2. Menyalin kolom Workorder
processed_data['Workorder'] = data['Workorder']

# 3. Memisahkan kolom SC Order
processed_data['SC Order No/Track ID/CSRM No'] = data['SC Order No/Track ID/CSRM No'].apply(lambda x: x.split('_')[0] if isinstance(x, str) else x)

def safe_split_part2(x):
    if isinstance(x, str) and '-' in x:
        parts = x.split('-')
        if len(parts) > 1:
            subparts = parts[1].split('_')
            if len(subparts) > 0:
                return subparts[0]
    return None

processed_data['SC Order No/Track ID/CSRM No 2'] = data['SC Order No/Track ID/CSRM No'].apply(safe_split_part2)

# 5. Menyalin kolom Service ID / Service Number
processed_data['Service No.'] = data['Service No.']

# 6. Menyalin kolom CRM Order Type
processed_data['CRM Order Type'] = data['CRM Order Type']

# 7. Menyalin kolom Status
processed_data['Status'] = data['Status']

# 9. Menyalin kolom Service Address
processed_data['Address'] = data['Address']

# 10. Menyalin kolom Customer Name
processed_data['Customer Name'] = data['Customer Name']

# 8. Menyalin kolom Workzone
processed_data['Workzone'] = data['Workzone']

# 11. Menyalin kolom Telephon Number
processed_data['Booking Date'] = data['Booking Date'].astype(str).str.split('.').str[0]

# 11. Menyalin kolom Telephon Number
processed_data['Contact Number'] = data['Contact Number']

# Mengisi nomor kontak yang kosong berdasarkan Customer Name dari data yang sama
contact_map = processed_data.loc[processed_data['Contact Number'].notna() & (processed_data['Contact Number'] != ''), ['Customer Name', 'Contact Number']].drop_duplicates()
contact_dict = dict(zip(contact_map['Customer Name'], contact_map['Contact Number']))

def fill_contact_number(row):
    if pd.isna(row['Contact Number']) or row['Contact Number'] == '':
        return contact_dict.get(row['Customer Name'], row['Contact Number'])
    else:
        return row['Contact Number']
    
processed_data['Contact Number'] = processed_data.apply(fill_contact_number, axis=1)

# 12. Menyalin kolom CRM Order Type tapi hanya yang CREATE dan MIGRATE
processed_data = processed_data[processed_data['CRM Order Type'].isin(['CREATE', 'MIGRATE'])]

# 13. Memfilter hanya yang tanggal 'Date Created' pada bulan Oktober
processed_data = processed_data[(processed_data['Date Created'].dt.month == 1)|(processed_data['Date Created'].dt.month == 2)]

# 14. Mengatur format tanggal menjadi dd/mm/yyyy hh:mm:ss
processed_data['Date Created'] = processed_data['Date Created'].dt.strftime('%d/%m/%Y %H:%M:%S')

# 14. Mengurutkan DataFrame berdasarkan kolom Workzone dalam urutan abjad
processed_data = processed_data.sort_values(by='Workzone', ascending=True)

# 15. Mengatur urutan kolom sesuai permintaan
final_output = processed_data[['Workzone', 'Date Created', 'SC Order No/Track ID/CSRM No',
                                 'Service No.', 'Workorder', 
                                 'Customer Name', 'Address', 
                                 'Contact Number', 'CRM Order Type','Booking Date']]

# Step 2: Baca data dari Google Sheets
spreadsheet_name = "Salinan dari NEW GDOC WSA FULFILLMENT"  # Nama spreadsheet
spreadsheet = client.open(spreadsheet_name)  # Membuka spreadsheet berdasarkan nama
worksheet = spreadsheet.sheet1  # Mengakses sheet pertama

# Mengambil semua data dari Google Sheets
google_data = worksheet.get_all_records()
google_df = pd.DataFrame(google_data)

print("Kolom yang ada di DataFrame Google Sheets:")
print(google_df.columns)

# Step 3: Bandingkan dan hapus duplikat
# Mengambil nilai SC Order dari Google Sheets
if 'SC Order No/Track ID/CSRM No' in google_df.columns:
    google_sc_order = google_df['SC Order No/Track ID/CSRM No']  # Ganti dengan kolom yang tepat jika berbeda

    # Menghitung data unik dari processed_data yang tidak ada di google_sc_order
    unique_data = processed_data[~processed_data['SC Order No/Track ID/CSRM No'].isin(google_sc_order)]

    # Step 4: Menyimpan data yang tidak terduplikasi di file Excel bar  u
    unique_data.to_excel(output_file, index=False)

    print(f"Data yang tidak terduplikasi telah disimpan ke {output_file}.")
else:
    print("Kolom 'N. SC' tidak ditemukan di DataFrame Google Sheets.")