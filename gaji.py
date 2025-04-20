import json
import gspread
import streamlit as st
from difflib import get_close_matches
from datetime import datetime, timedelta
from cryptography.fernet import Fernet
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import plotly.express as px

# === Kunci Enkripsi ===
ENCRYPTION_KEY = b"o0OSj6LwOvRIiZihiTslAMdpdIIuxFvYZh70PYD_BSI="  # Ganti dengan kunci ENCRYPTION_KEY Anda

# === Fungsi Dekripsi JSON ===
def bongkar_json(input_file):
    cipher = Fernet(ENCRYPTION_KEY)
    try:
        with open(input_file, "rb") as file:
            encrypted_data = file.read()
        decrypted_data = cipher.decrypt(encrypted_data)
        return json.loads(decrypted_data.decode())
    except FileNotFoundError:
        st.error(f"❌ File tidak ditemukan: {input_file}")
        return None
    except Exception as e:
        st.error(f"❌ Gagal mendekripsi atau membaca file: {e}")
        st.error("Pastikan kunci enkripsi (ENCRYPTION_KEY) dan file kredensial sudah benar.")
        return None

# === Autentikasi Google Sheets ===
@st.cache_resource
def connect_gC(credentials_data):
    if not credentials_data:
        return None
    try:
        # Tentukan scope akses
        scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        
        # Membuat objek Credentials dari JSON
        creds = Credentials.from_service_account_info(credentials_data, scopes=scope)
        
        # Cek apakah kredensial sudah kadaluarsa
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        # Menghubungkan ke Google Sheets
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"❌ Gagal mengautentikasi Google Sheets: {e}")
        return None

# === Fungsi membuka Spreadsheet ===
def buka_spreadsheet():
    global worksheet 
    worksheet = None  # Inisialisasi worksheet
    spreadsheet = None
    sheet_names = []

    # === Dekripsi kredensial ===
    encrypted_file_path = "izingoogle_encrypted.json"
    creds_data = bongkar_json(encrypted_file_path)

    if not creds_data:
        # Pesan error sudah ditampilkan di dalam fungsi bongkar_json
        st.stop()

    # === Hubungkan ke Google Sheets ===
    gc = connect_gC(creds_data)

    if not gc:
        # Pesan error sudah ditampilkan di dalam fungsi connect_gC
        st.stop()

    st.success("✅ Berhasil terhubung ke Google Sheets!")

    # === Buka spreadsheet ===
    spreadsheet_key = "1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps"

    try:
        spreadsheet = gc.open_by_key(spreadsheet_key)
        sheet_names = [ws.title for ws in spreadsheet.worksheets()]
        st.success(f"📄 Berhasil membuka spreadsheet: {spreadsheet.title}")

        if sheet_names:
            # Default ke sheet pertama atau sheet yang terakhir dipilih jika ada di session state
            if "selected_sheet_name" not in st.session_state or st.session_state.selected_sheet_name not in sheet_names:
                st.session_state.selected_sheet_name = sheet_names[0]

            selected_sheet_name = st.selectbox(
                "Pilih Sheet:",
                sheet_names,
                key="select_sheet",
                index=sheet_names.index(st.session_state.selected_sheet_name)
            )
            st.session_state.selected_sheet_name = selected_sheet_name  # Simpan pilihan ke session state

            worksheet = spreadsheet.worksheet(selected_sheet_name)
        else:
            st.error("❌ Spreadsheet tidak memiliki sheet yang tersedia.")
            st.stop()  # Berhenti jika tidak ada sheet

    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"❌ Spreadsheet dengan key '{spreadsheet_key}' tidak ditemukan. Periksa kembali key atau hak akses.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Gagal membuka spreadsheet atau mengambil nama sheet: {e}")
        st.stop()  # Berhenti jika gagal membuka spreadsheet

    # Pastikan worksheet tersedia setelah fungsi dipanggil
    if worksheet:
        st.write(f"Worksheet yang dibuka: {worksheet.title}")
    else:
        st.error("❌ Gagal membuka worksheet.")

# --- Fungsi convert_time (Dari kode terakhir user) ---
def convert_time(waktu):
    """Konversi waktu dari string HH:MM ke objek datetime."""
    if not isinstance(waktu, str): # Tambah pengecekan tipe
        return None
    try:
        return datetime.strptime(str(waktu).strip(), "%H:%M") # Tambah strip() dan pastikan string
    except ValueError:
        return None  # Jika format tidak valid


# --- Fungsi hitung_durasi (Mengikuti LOGIKA LAMA USER untuk mencocokkan hasil gambar) ---
def hitung_durasi(jm, jk, im, ik, batas_kerja_menit):
    """
    Menghitung total jam kerja dan lembur berdasarkan batas kerja normal.
    Menggunakan logika penjumlahan interval (Ik-Jm) + (Jk-Im) saat istirahat ada,
    sesuai dengan pola hasil pada gambar.
    """
    # Mengonversi input string ke objek datetime atau None
    jm_time = convert_time(jm)
    jk_time = convert_time(jk)
    im_time = convert_time(im)
    ik_time = convert_time(ik)

    if None in [jm_time, jk_time]:
        # Mengembalikan 6 nilai sesuai ekspektasi kode Streamlit lainnya
        return (0, 0, 0, 0, 0, 0)

    # Koreksi jika Jam Keluar terjadi setelah tengah malam (hari berikutnya)
    # Asumsi Jam Masuk dan Jam Keluar mendefinisikan span utama
    if jk_time < jm_time:
        jk_time += timedelta(days=1)

    total_kerja_menit_mentah = 0 # Inisialisasi total kerja dalam menit mentah

    # Jika ada data istirahat yang valid
    if im_time and ik_time:
        # --- LOGIKA PENJUMLAHAN INTERVAL SESUAI PERMINTAAN ---
        # Koreksi Ik jika terjadi setelah tengah malam relative terhadap Jm
        # Catatan: ini bukan koreksi durasi istirahat itu sendiri (Ik vs Im)
        if ik_time < jm_time:
             ik_time += timedelta(days=1)

        # Koreksi Im jika terjadi setelah tengah malam relative terhadap Jm (Tambahkan ini jika perlu berdasarkan data)
        # if im_time < jm_time:
        #     im_time += timedelta(days=1)
        # Namun, berdasarkan logika (Jk - Im), Im biasanya di antara Jm dan Jk.
        # Jika Im > Jk (sudah dikoreksi), Jk - Im akan negatif.

        # Menghitung 'kerja awal' (Jm ke Ik) dan 'kerja akhir' (Im ke Jk) dalam menit
        # Menggunakan .total_seconds() / 60 untuk mendapatkan menit
        kerja_awal = (ik_time - jm_time).total_seconds() / 60
        kerja_akhir = (jk_time - im_time).total_seconds() / 60

        # Menjumlahkan kedua interval ini untuk mendapatkan total kerja
        # Menggunakan max(0, ...) untuk menghindari hasil negatif total kerja
        total_kerja_menit_mentah = max(0, kerja_awal + kerja_akhir)
        # --- AKHIR LOGIKA PENJUMLAHAN INTERVAL ---

    else:
        # Jika tidak ada data istirahat, hitung total jam kerja tanpa istirahat
        total_kerja_menit_mentah = (jk_time - jm_time).total_seconds() / 60

    # Hitung lembur berdasarkan batas kerja normal (dalam menit)
    lembur_menit_mentah = max(0, total_kerja_menit_mentah - batas_kerja_menit)

    # Hitung jam dan menit terpisah dari total menit mentah
    total_jam = int(total_kerja_menit_mentah // 60)
    total_menit = int(total_kerja_menit_mentah % 60)
    lembur_jam = int(lembur_menit_mentah // 60)
    lembur_menit = int(lembur_menit_mentah % 60)


    # Mengembalikan 6 nilai sesuai ekspektasi kode Streamlit lainnya
    return (
        total_jam, total_menit,
        lembur_jam, lembur_menit,
        int(round(total_kerja_menit_mentah)), # Total menit mentah (dibulatkan)
        int(round(lembur_menit_mentah))       # Lembur menit mentah (dibulatkan)
    )
# --- Akhir dari Fungsi hitung_durasi (Mengikuti Logika Lama User) ---


# --- Fungsi find_month_columns (Tetap sama) ---
def find_month_columns(df, bulan_target, bulan_list):
    """Finds the range of columns associated with a given month name in the first row."""
    if df.empty: return None, None # Handle empty dataframe
    # Ensure first row has enough elements if accessed by iloc
    if len(df.columns) == 0: return None, None
    # Try to access the first row safely
    try:
        month_row = df.iloc[0]
    except IndexError:
        return None, None # DataFrame has no rows

    start_col = None
    end_col = None
    col_count = len(df.columns)
    # Pre-process month list for case-insensitive comparison
    known_months_lower = {m.strip().lower() for m in bulan_list}

    # Find the start column
    for idx, month_cell_value in enumerate(month_row):
        # Handle potential non-string values in month row safely
        if isinstance(month_cell_value, str):
             month = month_cell_value.strip().lower()
             if month == bulan_target.lower():
                 start_col = idx
                 break # Found the start

    if start_col is None:
        return None, None # Target month not found

    # Find the end column (start of the next known month or end of DataFrame)
    for idx in range(start_col + 1, col_count):
        # Ensure index is valid for month_row before accessing iloc
        if idx < len(month_row):
             month_cell_value = month_row.iloc[idx]
        else:
             continue # Index out of bounds for this row iteration

        if pd.isna(month_cell_value) or str(month_cell_value).strip() == "": continue # Skip empty cells

        if isinstance(month_cell_value, str):
            current_cell_month = month_cell_value.strip().lower()
            # If the cell is not empty AND contains a *different* known month name
            if current_cell_month in known_months_lower and current_cell_month != bulan_target.lower():
                end_col = idx
                break # Found the start of the next month

    # If no next month was found, the range goes to the end of the DataFrame
    if end_col is None:
        end_col = col_count

    return start_col, end_col

# --- Fungsi find_category_column (Tetap sama) ---
def find_category_column(df, category_row_index, category_target, start_col, end_col):
    """Finds the column index for a specific category within a given column range."""
    # Ensure category_row_index is valid and DataFrame has enough columns/rows
    if category_row_index >= len(df) or df.empty or start_col is None or end_col is None:
         return None
    # Ensure start_col is within DataFrame bounds before trying to access columns
    if start_col is not None and start_col >= df.shape[1]: # Check if start_col is valid index
         return None


    category_row = df.iloc[category_row_index]
    # Ensure effective_end_col does not exceed the actual number of columns in the DataFrame or the row
    effective_end_col = min(end_col, len(df.columns), len(category_row)) if end_col is not None else len(df.columns)


    for col_idx in range(start_col if start_col is not None else 0, effective_end_col):
        # Ensure col_idx is within the bounds of category_row before accessing it
        if col_idx >= len(category_row):
             continue

        header = category_row.iloc[col_idx]
        if pd.isna(header) or str(header).strip() == "": continue # Skip empty header cells

        # Case-insensitive comparison
        if isinstance(header, str) and header.strip().lower() == category_target.lower():
            return col_idx
    return None # Category not found in the specified range

# --- Fungsi fetch_data_range_from_df (Memanggil hitung_durasi dengan logika lama user) ---
def fetch_data_range_from_df(df, bulan_awal, tanggal_awal, bulan_akhir, tanggal_akhir, batas_kerja_menit, bulan_list):
    """
    Fungsi untuk mengambil dan memproses data dari DataFrame Pandas dalam rentang tanggal tertentu.
    Memanggil hitung_durasi dengan logika penjumlahan interval (Ik-Jm) + (Jk-Im).
    """
    data_list = []
    # Pastikan bulan_list tidak kosong sebelum akses index
    if not bulan_list:
        st.error("Daftar bulan kosong.")
        return []
    try:
        bulan_awal_idx = bulan_list.index(bulan_awal)
        bulan_akhir_idx = bulan_list.index(bulan_akhir)
    except ValueError:
        st.error(f"Nama bulan awal ('{bulan_awal}') atau akhir ('{bulan_akhir}') tidak ditemukan dalam daftar bulan.")
        return []

    if bulan_awal_idx > bulan_akhir_idx:
        st.error("⚠️ Bulan akhir harus setelah atau sama dengan bulan awal.")
        return []

    # --- Iterate through relevant months ---
    for bulan_idx in range(bulan_awal_idx, bulan_akhir_idx + 1):
        current_bulan = bulan_list[bulan_idx]
        start_col, end_col = find_month_columns(df, current_bulan, bulan_list)

        if start_col is None:
            st.warning(f"⚠️ Kolom untuk bulan '{current_bulan}' tidak ditemukan dalam sheet.")
            continue # Bulan tidak ditemukan, lanjut ke bulan berikutnya

        # --- Find Category Columns for the current month ---
        category_row_index = 2 # Categories are in the THIRD row (index 2)
        col_map = {}

        # Find 'Tanggal' column first
        tanggal_col_idx = find_category_column(df, category_row_index, "Tanggal", start_col, end_col)

        # Backup logic if "Tanggal" header not found
        # Assume the first column index within the month range (start_col) is the date column
        # Check if start_col is a valid column index in the DataFrame
        if tanggal_col_idx is None and start_col is not None and start_col < df.shape[1]:
             tanggal_col_idx = start_col


        if tanggal_col_idx is None: # Jika kolom tanggal tetap tidak ditemukan
            st.warning(f"⚠️ Tidak dapat menemukan kolom 'Tanggal' untuk bulan '{current_bulan}'. Bulan ini dilewati.")
            continue

        col_map["Tanggal"] = tanggal_col_idx

        # Define categories based *exactly* on headers needed for later calculation + display
        # Added 'Libur' back as a category to be fetched, as it's used in the summary count
        kategori_list_needed = [
            "Jam Masuk", "Jam Kluar", "Istirahat Masuk", "Istirahat Kluar",
            "Lokasi Kerja", "Keterangan", "Paraf", "2 Regu", # General info
            "Ke Samboja", "Ke Gunung guntur", "Libur" # Specific keys used in original calc loop + summary
        ]

        for kategori in kategori_list_needed:
            cat_col_idx = find_category_column(df, category_row_index, kategori, start_col, end_col)
            col_map[kategori] = cat_col_idx # Store index or None

        # --- Iterate through Data Rows for the current month ---
        # Start from index 3 (fourth row) as rows 0, 1, 2 are headers
        for row_idx in range(3, len(df)):
            if row_idx >= len(df): break # Safety check

            row_data = df.iloc[row_idx]

            # --- Get and Validate Date (Original Logic + Safety) ---
            # Ensure the tanggal_col_idx is valid for the current row's length
            tgl_col_idx_safe = col_map.get("Tanggal")
            if tgl_col_idx_safe is None or tgl_col_idx_safe >= len(row_data):
                 continue # Skip row if Tanggal column index is invalid for this row

            try:
                tanggal_raw = row_data.iloc[tgl_col_idx_safe]
                if pd.isna(tanggal_raw) or str(tanggal_raw).strip() == "": continue

                # Clean the date string (remove dots, take first part if any)
                tanggal_str_cleaned = str(tanggal_raw).split('.')[0]
                # Filter only digits, handle potential floats read as strings
                tanggal_str = "".join(filter(str.isdigit, tanggal_str_cleaned))

                if not tanggal_str: continue # Skip if no digits found

                tanggal = int(tanggal_str)
                if not (1 <= tanggal <= 31): continue # Basic date range check

            except (ValueError, TypeError, IndexError):
                continue # Skip row on any date parsing error
            except Exception as e: # Catch unexpected errors during date processing
                # st.warning(f"Error processing date in row {row_idx+1}: {e}") # Optional debug warning
                continue


            # --- Check if Date is within User Selected Range ---
            # Check against start date
            if bulan_idx == bulan_awal_idx and tanggal < tanggal_awal:
                continue
            # Check against end date
            if bulan_idx == bulan_akhir_idx and tanggal > tanggal_akhir:
                 # If we are in the end month and the date exceeds the end date,
                 # we can actually break the inner loop for this month
                 break # optimization: stop processing rows for this month
            elif bulan_idx > bulan_akhir_idx:
                 # This case should ideally not be reached due to the outer loop range,
                 # but as a safeguard, break if somehow we are past the end month.
                 break # Safeguard break


            # --- Extract Data for Categories (Original Logic + Safety) ---
            result = {"Bulan": current_bulan, "Tanggal": tanggal}
            for kategori, col_idx in col_map.items():
                if kategori == "Tanggal": continue # Already handled

                # Ensure column index is valid for the current row's length
                if col_idx is not None and col_idx < len(row_data):
                    cell_value = row_data.iloc[col_idx]
                    # Store "-" for empty/NaN cells
                    result[kategori] = "-" if pd.isna(cell_value) or str(cell_value).strip() == "" else str(cell_value).strip()
                else:
                    # Store "-" if the category column wasn't found for this month
                    result[kategori] = "-"


            # --- Convert Times and Calculate Durations (using hitung_durasi dengan logika lama) ---
            # Get values using .get() for safety in case column wasn't mapped
            jm_val = result.get("Jam Masuk", "-")
            jk_val = result.get("Jam Kluar", "-") # Sesuai header excel
            im_val = result.get("Istirahat Masuk", "-")
            ik_val = result.get("Istirahat Kluar", "-") # Sesuai header excel

            # Memanggil fungsi hitung_durasi dengan logika penjumlahan interval
            # Note: input ke hitung_durasi adalah STRING, bukan objek datetime
            (total_jam, total_menit, lembur_jam, lembur_menit,
             total_kerja_menit_val, lembur_menit_val) = hitung_durasi(
                 jm_val, jk_val, im_val, ik_val, batas_kerja_menit # Pass string values
             )

            # Store results (Keep original formatting logic for display strings)
            # Check if total calculated work time (in raw minutes) is greater than 0 before formatting
            if total_kerja_menit_val > 0:
                # Use f-string formatting for clarity
                result["Total Jam Kerja"] = f"{total_jam} jam {total_menit} menit"
                result["Total Lembur"] = f"{lembur_jam} jam {lembur_menit} menit"
            else:
                # If total work minutes is 0, show "-"
                result["Total Jam Kerja"] = "-"
                result["Total Lembur"] = "-"

            # Store raw minutes (Original code added these, keep them for aggregation later)
            result["_total_kerja_menit"] = total_kerja_menit_val
            result["_total_lembur_menit"] = lembur_menit_val

            data_list.append(result)

    return data_list

def main():
    # === Tampilan Streamlit ===

    # Daftar bulan (pastikan sudah terdefinisi sesuai header Bulan di sheet baris 1)
    bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                "Juli", "Agustus", "September", "Oktober", "November", "Desember"]


    st.title("📊 Analisis Gaji Crew")

    # Input User (Original layout)
    col1, col2 = st.columns(2)
    with col1:
        # Set default index more robustly
        default_start_index = 0
        if "bulan_awal" in st.session_state and st.session_state.bulan_awal in bulan_list:
            default_start_index = bulan_list.index(st.session_state.bulan_awal)
        bulan_input = st.selectbox("Pilih bulan awal:", bulan_list, index=default_start_index, key="bulan_awal")

        default_tanggal_awal = 1
        if "tgl_awal" in st.session_state and 1 <= st.session_state.tgl_awal <= 31:
            default_tanggal_awal = st.session_state.tgl_awal
        tanggal_awal = st.number_input("Masukkan tanggal awal:", min_value=1, max_value=31, value=default_tanggal_awal, key="tgl_awal")

    with col2:
        # Set default index more robustly
        default_end_index = len(bulan_list) - 1 # Default ke bulan terakhir
        if "bulan_akhir" in st.session_state and st.session_state.bulan_akhir in bulan_list:
            default_end_index = bulan_list.index(st.session_state.bulan_akhir)
        # Ensure end index is not before start index by default on first run
        if "bulan_awal" in st.session_state and default_end_index < default_start_index:
            default_end_index = default_start_index

        bulan_akhir = st.selectbox("Pilih bulan akhir:", bulan_list, index=default_end_index, key="bulan_akhir")


        default_tanggal_akhir = 1
        if "tgl_akhir" in st.session_state and 1 <= st.session_state.tgl_akhir <= 31:
            default_tanggal_akhir = st.session_state.tgl_akhir
        # Set default end date to 31 if end month is the same as start month and start date is 1
        if bulan_input == bulan_akhir and tanggal_awal == 1 and default_tanggal_akhir == 1:
            default_tanggal_akhir = 31

        tanggal_akhir = st.number_input("Masukkan tanggal akhir:", min_value=1, max_value=31, value=default_tanggal_akhir, key="tgl_akhir")


    # Input batas kerja normal dalam jam
    default_batas_jam = 9.0
    if "batas_jam" in st.session_state:
        default_batas_jam = st.session_state.batas_jam

    batas_kerja_jam = st.number_input("Batas kerja normal (jam)", min_value=0.0, max_value=24.0, value=default_batas_jam, step=0.5, key="batas_jam")
    batas_kerja_menit = int(batas_kerja_jam * 60)  # Konversi ke menit (integer)

    # Inisialisasi session state jika belum ada
    if "hasil_data" not in st.session_state:
        st.session_state.hasil_data = None

    # Tombol Ambil Data
    if st.button("Ambil Data", key="fetch_button"): # Tambah key
        # Reset hasil sebelum fetch baru
        st.session_state.hasil_data = None
        if worksheet: # Pastikan worksheet sudah dipilih dan valid
            with st.spinner("Mengambil dan memproses data..."): # Beri feedback
                try:
                    # *** Get data and create DataFrame ***
                    data_values = worksheet.get_all_values()

                    if not data_values:
                        st.warning("⚠️ Sheet yang dipilih kosong.")
                        st.session_state.hasil_data = [] # Set state ke list kosong
                    else:
                        # Membuat DataFrame dari nilai yang diambil
                        # Hapus argumen 'header=None' karena tidak valid
                        df_data = pd.DataFrame(data_values) # *** PERBAIKAN DI SINI ***

                        # *** Call function with DataFrame and bulan_list ***
                        fetched_data = fetch_data_range_from_df(
                            df_data,              # Pass the DataFrame
                            bulan_input,
                            tanggal_awal,
                            bulan_akhir,
                            tanggal_akhir,
                            batas_kerja_menit,
                            bulan_list            # Pass the list of months
                        )
                        st.session_state.hasil_data = fetched_data # Simpan hasil ke session state

                except gspread.exceptions.APIError as api_err:
                    st.error(f"❌ Gagal mengambil data dari Google Sheets (API Error): {api_err}")
                    st.session_state.hasil_data = None # Set state error
                except Exception as e:
                    st.error(f"❌ Terjadi error saat mengambil atau memproses data: {e}")
                    st.exception(e) # Menampilkan traceback untuk debugging
                    st.session_state.hasil_data = None # Set state error
        else:
            # Jika worksheet belum siap (misal, spreadsheet gagal dibuka)
            st.error("❌ Tidak dapat mengambil data karena sheet belum siap/terpilih.")


    # Ambil hasil dari session_state untuk ditampilkan
    hasil_data = st.session_state.get("hasil_data", None)

    # --- Logika Tampilan Hasil ---
    if isinstance(hasil_data, list): # Hanya proses jika hasil_data adalah list
        if len(hasil_data) > 0:
            # Buat DataFrame untuk tampilan, hapus kolom internal
            df_display = pd.DataFrame(hasil_data)
            internal_cols = ['_total_kerja_menit', '_total_lembur_menit']
            # Pastikan kolom ada sebelum dihapus
            cols_to_drop = [col for col in internal_cols if col in df_display.columns]
            if cols_to_drop:
                df_display = df_display.drop(columns=cols_to_drop)

            st.write(f"📌 Menampilkan data dari **{tanggal_awal} {bulan_input}** hingga **{tanggal_akhir} {bulan_akhir}**")
            st.dataframe(df_display, use_container_width=True) # Gunakan lebar container

            # --- Perhitungan total (Menggunakan menit mentah yang disimpan) ---
            # Mengambil total menit dari kolom internal yang disimpan
            total_menit_kerja_sum = sum(item.get("_total_kerja_menit", 0) for item in hasil_data)
            total_menit_lembur_sum = sum(item.get("_total_lembur_menit", 0) for item in hasil_data)

            # Menghitung Samboja, Guntur, Off (Logika asli dari kode Anda)
            total_samboja = 0
            total_guntur = 0
            total_off = 0 # Menghitung total hari OFF

            for item in hasil_data: # Iterasi list of dict asli
                # Hitung Samboja
                samboja_val = item.get("Ke Samboja", "")
                if isinstance(samboja_val, str) and samboja_val.strip().lower() == "sendiri":
                    total_samboja += 1

                # Hitung Guntur
                guntur_val = item.get("Ke Gunung guntur", "") # Perhatikan spasi di "Gunung guntur"
                if isinstance(guntur_val, str) and guntur_val.strip().lower() == "sendiri":
                    total_guntur += 1

                # Hitung OFF (cek kolom 'Libur' atau 'Keterangan')
                libur_val = item.get("Libur", "")
                keterangan_val = item.get("Keterangan", "")

                # Cek kolom 'Libur' terlebih dahulu
                if isinstance(libur_val, str) and libur_val.strip().upper() == "OFF":
                    total_off += 1
                # Jika kolom 'Libur' kosong atau bukan 'OFF', cek kolom 'Keterangan'
                elif isinstance(keterangan_val, str) and keterangan_val.strip().upper() == "OFF":
                    total_off += 1


            # Konversi total menit ke jam + menit untuk tampilan
            def konversi_menit_display(menit):
                if not isinstance(menit, (int, float)) or pd.isna(menit) or menit < 0:
                    return "0 jam 0 menit"
                menit_int = int(round(menit)) # Bulatkan sebelum konversi
                jam = menit_int // 60
                sisa = menit_int % 60
                return f"{jam} jam {sisa} menit"

            # Tampilkan Rekap
            st.markdown("### 🔢 Rekap Data")
            st.write(f"🕒 Total Jam Kerja: **{konversi_menit_display(total_menit_kerja_sum)}**")
            st.write(f"⏱️ Total Jam Lembur: **{konversi_menit_display(total_menit_lembur_sum)}**")
            st.write(f"🏕️ Ke Gunung Guntur: **{total_guntur} hari**")
            st.write(f"🏝️ Ke Samboja: **{total_samboja} hari**")
            st.write(f"🛌 Total Libur (OFF): **{total_off} hari**")


            # === Diagram ===
            # Buat data untuk chart
            # Gunakan total menit mentah yang dijumlahkan, lalu konversi ke jam untuk chart
            chart_data = {
                "Kategori": ["Jam Kerja (jam)", "Lembur (jam)", "Ke Gunung Guntur (hari)", "Ke Samboja (hari)", "Libur (OFF) (hari)"],
                "Jumlah": [
                    round(total_menit_kerja_sum / 60, 2), 
                    round(total_menit_lembur_sum / 60, 2),
                    total_guntur,
                    total_samboja,
                    total_off
                ]
            }
            df_chart = pd.DataFrame(chart_data)

            # Filter data dengan jumlah > 0 agar chart lebih bersih
            df_chart = df_chart[df_chart['Jumlah'] > 0].reset_index(drop=True)

            if not df_chart.empty:
                st.markdown("### 📊 Visualisasi Data")
                # Gunakan key unik untuk radio button
                chart_type = st.radio("Pilih tipe grafik:", ["Diagram Batang", "Diagram Lingkaran"], index=1, key="chart_type_radio")

                if chart_type == "Diagram Batang":
                    fig = px.bar(df_chart, x="Kategori", y="Jumlah", color="Kategori",
                                text_auto=True, # Format teks otomatis
                                title="Ringkasan Kinerja Crew", height=400)
                    fig.update_traces(textposition='outside')
                else: # Diagram Lingkaran
                    fig = px.pie(df_chart, names="Kategori", values="Jumlah",
                                title="Distribusi Aktivitas Crew", height=400,
                                hole=.3) # Tambah sedikit lubang di tengah
                    # Tampilkan label, persen, dan nilai di dalam pie chart
                    # Gunakan texttemplate untuk format yang konsisten
                    fig.update_traces(textposition='inside', textinfo='percent+label',
                                    texttemplate='%{label}: %{value:.2f}') # Tampilkan label dan nilai (jam/hari)
                    fig.update_layout(uniformtext_minsize=12, uniformtext_mode='hide') # Sembunyikan teks jika terlalu kecil


                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("ℹ️ Tidak ada data agregat dengan jumlah > 0 untuk divisualisasikan.")

        elif len(hasil_data) == 0: # Jika fetch berhasil tapi tidak ada data dalam rentang tanggal
            st.info("ℹ️ Tidak ada data absensi yang ditemukan untuk rentang tanggal dan sheet yang dipilih.")

if __name__ == "__main__":
    buka_spreadsheet()
    main()