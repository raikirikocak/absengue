import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import re
from difflib import get_close_matches
from admin_page import admin_dashboard
from cryptography.fernet import Fernet


# === 1. Gunakan kunci enkripsi yang diberikan ===
ENCRYPTION_KEY = b"o0OSj6LwOvRIiZihiTslAMdpdIIuxFvYZh70PYD_BSI="

# === 2. Fungsi untuk mendekripsi file JSON ===
def decrypt_json(input_file):
    cipher = Fernet(ENCRYPTION_KEY)

    # Baca file terenkripsi
    with open(input_file, "rb") as file:
        encrypted_data = file.read()

    # Dekripsi dan ubah kembali ke JSON
    decrypted_data = cipher.decrypt(encrypted_data)
    return json.loads(decrypted_data.decode())

# Fungsi untuk autentikasi ke Google Sheets
@st.cache_resource

def authorize_google_sheets(credentials_data):
    try:
        # Autentikasi menggunakan dictionary langsung (bukan file)
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_data, scope)
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"❌ Gagal mengautentikasi Google Sheets: {e}")
        return None

# Fungsi untuk mendapatkan data dari Google Sheets
def fetch_data(worksheet):
    try:
        return worksheet.get_all_values()
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca data dari Google Sheets: {e}")
        return []
    
def logout_button():
    if st.button("Logout"):
        st.session_state.clear()  # Hapus seluruh session state
        st.success("Berhasil logout. Semua sesi dihapus. Kembali ke halaman login...")
        st.rerun()  # Refresh halaman untuk efek langsung

# === 4. Fungsi Utama ===
def login_and_access_sheet():
    # Dekripsi kredensial Google Sheets
    encrypted_file_path = "izingoogle_encrypted.json"
    try:
        creds_data = decrypt_json(encrypted_file_path)
    except Exception as e:
        st.error(f"❌ Gagal mendekripsi file: {e}")
        return
    
    # Autentikasi Google Sheets
    gc = authorize_google_sheets(creds_data)
    if not gc:
        st.stop()
    
    st.success("✅ Berhasil terhubung ke Google Sheets!")
    
    # **ADMIN LOGIN**
    if "is_admin" not in st.session_state:
        st.session_state["is_admin"] = False
    
    # **CLIENT LOGIN (INPUT ID & PILIH SHEET BERSAMAAN)**
    if "is_cliente" not in st.session_state:
        st.session_state["is_cliente"] = False
    
    # Jika sudah login sebagai admin atau klien, tampilkan dashboard yang sesuai
    if st.session_state["is_admin"]:
        admin_dashboard()
        return  # Hentikan eksekusi jika admin login
    
    if st.session_state["is_cliente"]:
        client_dashboard()
        return  # Hentikan eksekusi jika klien login
    
    # Jika belum login, tampilkan form login dan validasi ID
    admin_col, client_col = st.columns(2)
    
    with admin_col:
        st.subheader("Login Admin")
        admin_username = st.text_input("Username Admin:")
        admin_password = st.text_input("Password Admin:", type="password")
        
        if st.button("Login Admin"):
            if admin_username == "admin" and admin_password == "1234":
                st.session_state["is_admin"] = True
                st.success("Login Admin berhasil! Mengalihkan ke halaman admin...")
                st.rerun()
            else:
                st.error("Username atau password salah.")
    
    with client_col:
        st.subheader("Login Klien")

        # Coba baca file id_mapping.json
        try:
            with open("id_mapping.json", "r") as f:
                id_mapping = json.load(f)
        except json.JSONDecodeError:
            st.error("Format file 'id_mapping.json' tidak valid.")
            return
        except Exception as e:
            st.error(f"Gagal memuat id_mapping.json: {e}")
            st.stop()

        # Google Sheets API setup
        spreadsheet_key = "1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps"
        
        try:
            spreadsheet = gc.open_by_key(spreadsheet_key)
            sheet_names = [sheet.title for sheet in spreadsheet.worksheets()]
        except Exception as e:
            st.error(f"Gagal membuka Google Sheets: {e}")
            st.stop()

        input_id = st.text_input("Masukkan ID Klien:")
        selected_sheet_name = st.selectbox("Pilih Sheet:", sheet_names, key="select_sheet")

        if st.button("Akses Sheet"):
            # Periksa apakah ID cocok dengan sheet yang dipilih
            if selected_sheet_name not in id_mapping or id_mapping[selected_sheet_name] != input_id:
                st.error("ID tidak valid untuk sheet yang dipilih.")
            else:
                st.success(f"ID valid untuk Sheet '{selected_sheet_name}'.")
                st.session_state["selected_sheet_name"] = selected_sheet_name
                st.session_state["is_cliente"] = True

                try:
                    worksheet = spreadsheet.worksheet(selected_sheet_name)
                    st.session_state["worksheet"] = worksheet
                    st.success(f"Sheet '{selected_sheet_name}' berhasil diakses.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal mengakses sheet '{selected_sheet_name}': {e}")


def isi_data_page(worksheet):
    # Halaman untuk mengisi data pada worksheet
    st.subheader("Isi Data")
    if worksheet:
        # Misalkan, di sini Anda akan mengisi data ke worksheet sesuai dengan kebutuhan
        st.write(f"Sheet yang dipilih: {worksheet.title}")
    # Pilihan nama bulan
    # Pilihan nama bulan
    bulan_options = [
        "Januari", "February", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    header_text = st.selectbox("Pilih nama bulan:", bulan_options)

    # Cari header bulan dengan toleransi salah ejaan
    header_row_values = worksheet.row_values(1)  # Ambil seluruh header di baris pertama
    matches = get_close_matches(header_text, header_row_values, n=1, cutoff=0.6)

    if not matches:
        st.error(f"Bulan '{header_text}' tidak ditemukan di spreadsheet.")
        return

    header_text = matches[0]
    cell_header = worksheet.find(header_text)
    st.success(f"Bulan '{header_text}' ditemukan di baris {cell_header.row}, kolom {cell_header.col}.")

    # Pilihan kategori
    kategori_options = [
        "Jam Masuk", "Jam Kluar", "Istirahat Masuk", "Istirahat Kluar",
        "Lembur Masuk", "Lembur Kluar", "Ke Gunung Guntur", "Ke Samboja", "Holiday"
    ]
    kategori = st.selectbox("Pilih kategori:", kategori_options)

    # Inisialisasi variabel input
    tanggal = None
    nilai = None
    antar_sendiri = None
    libur = None

        # Input tanggal tetap ada di awal agar tidak terjadi error saat digunakan di kategori lain
    tanggal = st.number_input("Masukkan tanggal (angka):", min_value=1, max_value=31, step=1)

    # Pilih nilai berdasarkan kategori
    if kategori in ["Ke Gunung Guntur", "Ke Samboja"]:
        antar_sendiri = st.selectbox("Pilih mode perjalanan:", ["Antar", "Sendiri"])
        nilai = f"({antar_sendiri})"


    elif kategori == "Holiday":
        libur = st.selectbox("pilih mode libur:", ["OFF"])
        nilai = "OFF"  # Tetapkan nilai menjadi "OFF" untuk libur

    else:
        nilai = st.text_input("Masukkan nilai waktu (format HH:MM):").strip()


        # Validasi format waktu
        if nilai and not re.match(r"^([01]\d|2[0-3]):[0-5]\d$", nilai):
            st.error("Format waktu tidak valid. Gunakan format HH:MM (contoh: 08:00).")
            return

    # Pastikan semua data valid sebelum melanjutkan ke penyimpanan
    if st.button("Simpan Data"):
        try:
            # Cari kategori di kolom bulan dan 8 kolom setelahnya
            kategori_ditemukan = False
            matched_kategori = None
            row_to_update = None

            for col_offset in range(0, 10):  # Periksa hingga 8 kolom tambahan
                current_col = cell_header.col + col_offset
                kategori_header_baris = worksheet.col_values(current_col)  # Ambil seluruh kolom
                kategori_header_baris_lower = [k.lower() for k in kategori_header_baris[cell_header.row:]]  # Lowercase untuk perbandingan

                # Cek kategori di bawah header bulan (case insensitive)
                possible_matches = get_close_matches(kategori.lower(), kategori_header_baris_lower, n=1, cutoff=0.7)
                if possible_matches:
                    matched_kategori_lower = possible_matches[0]  # Ambil kategori yang paling mirip (lowercase)
                    matched_kategori = kategori_header_baris[cell_header.row + kategori_header_baris_lower.index(matched_kategori_lower)]  # Kembalikan versi asli (case sensitive)
                    row_to_update = kategori_header_baris.index(matched_kategori) + 1  # Baris kategori ditemukan
                    kategori_ditemukan = True
                    break

            if not kategori_ditemukan:
                st.error(f"Kategori '{kategori}' tidak ditemukan di bawah bulan '{header_text}' atau kolom tambahan.")
                return

            # Cari tanggal di kolom pertama
            try:
                cell_tanggal = worksheet.find(str(tanggal), in_column=1)  # Kolom tanggal di spreadsheet
            except gspread.exceptions.CellNotFound:
                st.error(f"Tanggal '{tanggal}' tidak ditemukan di spreadsheet.")
                return

            if antar_sendiri:
                nilai = antar_sendiri  # Simpan nilai langsung tanpa tambahan format


            # Perbarui nilai di kolom kategori dan baris tanggal
            worksheet.update_cell(cell_tanggal.row, current_col, nilai)
            st.success(f"Data '{nilai}' berhasil ditambahkan pada kategori '{matched_kategori}' untuk tanggal {tanggal} di bulan '{header_text}'.")
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")



def hapus_data_page(worksheet):
    st.subheader("Hapus Data")
    if worksheet:
        # Misalkan, di sini Anda akan mengisi data ke worksheet sesuai dengan kebutuhan
        st.write(f"Sheet yang dipilih: {worksheet.title}")
    bulan_options = [
        "Januari", "Februari", "Maret", "April", "Mei", "Juni",
        "Juli", "Agustus", "September", "Oktober", "November", "Desember"
    ]
    st.write(f"Anda memilih: {bulan_options}")
    header_text = st.selectbox("Pilih nama bulan:", bulan_options).strip()

    # Pilihan kategori
    kategori_options = [
        "Jam Masuk", "Jam Kluar", "Istirahat Masuk", "Istirahat Kluar",
        "Lembur Masuk", "Lembur Kluar", "Ke Gunung Guntur", "Ke Samboja", "Holiday"
    ]
    st.write(f"Anda memilih: {kategori_options}")
    kategori = st.selectbox("Pilih kategori:", kategori_options).strip()

    # Input tanggal (PASTIKAN TIDAK HILANG)
    tanggal = st.number_input("Masukkan tanggal (angka):", min_value=1, max_value=31, step=1)

    if st.button("Hapus Data"):
        try:
            # Cek apakah worksheet sudah ada di session
            if "worksheet" not in st.session_state:
                st.error("Worksheet belum diakses. Harap akses terlebih dahulu.")
                st.stop()  # Berhenti eksekusi fungsi

            worksheet = st.session_state["worksheet"]

            # Cek apakah bulan ada di sheet
            header_row_values = worksheet.row_values(1)  # Ambil seluruh header di baris pertama
            matches = get_close_matches(header_text, header_row_values, n=1, cutoff=0.6)
            if not matches:
                st.error(f"Bulan '{header_text}' tidak ditemukan di spreadsheet.")
                st.stop()  # Berhenti eksekusi fungsi

            header_text = matches[0]
            cell_header = worksheet.find(header_text)

            # Cari kategori di bawah header bulan
            kategori_ditemukan = False
            matched_kategori = None
            row_to_update = None
            col_to_update = None

            for col_offset in range(0, 9):  # Periksa hingga 8 kolom tambahan
                current_col = cell_header.col + col_offset
                kategori_header_baris = worksheet.col_values(current_col)[cell_header.row:]  # Ambil kolom
                kategori_header_baris_lower = [k.lower() for k in kategori_header_baris]  # Lowercase untuk perbandingan

                possible_matches = get_close_matches(kategori.lower(), kategori_header_baris_lower, n=1, cutoff=0.7)
                if possible_matches:
                    matched_kategori_lower = possible_matches[0]
                    matched_kategori = kategori_header_baris[kategori_header_baris_lower.index(matched_kategori_lower)]
                    row_to_update = kategori_header_baris_lower.index(matched_kategori_lower) + cell_header.row + 1
                    col_to_update = current_col
                    kategori_ditemukan = True
                    break

            if not kategori_ditemukan:
                st.error(f"Kategori '{kategori}' tidak ditemukan di bawah bulan '{header_text}' atau kolom tambahan.")
                st.stop()  # Berhenti eksekusi fungsi

            # Cari tanggal di kolom pertama
            try:
                cell_tanggal = worksheet.find(str(tanggal), in_column=1)
            except gspread.exceptions.CellNotFound:
                st.error(f"Tanggal '{tanggal}' tidak ditemukan di spreadsheet.")
                st.stop()  # Berhenti eksekusi fungsi

            # Hapus data dari sel yang sesuai
            worksheet.update_acell(f"{gspread.utils.rowcol_to_a1(cell_tanggal.row, col_to_update)}", "")  
            st.success(f"Data pada kategori '{matched_kategori}' untuk tanggal {tanggal} di bulan '{header_text}' telah dihapus.")

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")

def lihat_data_page(worksheet):
    st.title("Lihat Data")
    st.write("Halaman untuk melihat data.")
    st.write("Menampilkan data yang tersimpan di Google Sheets.")

    # Tombol untuk membuka Google Sheets
    if st.button("Buka Google Sheets"):
        st.markdown(f"[Klik di sini untuk melihat data](https://docs.google.com/spreadsheets/d/1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps/edit?usp=sharing)")
        st.info("Silakan klik link di atas untuk melihat data lengkap di Google Sheets.")

def client_dashboard():
    # Pastikan klien sudah login dan worksheet sudah dipilih
    if "is_cliente" not in st.session_state or not st.session_state["is_cliente"]:
        st.error("❌ Anda belum login atau belum memilih sheet yang valid. Silakan login terlebih dahulu.")
        return
   
    worksheet = st.session_state["worksheet"]

    # Sidebar Menu
    st.sidebar.title("Menu")
    option = st.sidebar.selectbox("Pilih Menu", ["Isi Data", "Hapus Data", "Lihat Data"], key="menu_select")

    # Tombol Logout
    if st.sidebar.button("Logout"):
        st.session_state["is_cliente"] = False
        st.session_state["selected_sheet_name"] = None
        st.session_state["worksheet"] = None
        st.success("Anda telah logout. Silakan login kembali.")
        st.rerun()

    if option == "Isi Data":
        isi_data_page(worksheet)
    elif option == "Hapus Data":
        hapus_data_page(worksheet)
    elif option == "Lihat Data":
        lihat_data_page(worksheet)


if __name__ == "__main__":
   
    login_and_access_sheet()
  
