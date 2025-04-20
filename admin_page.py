import gspread
from oauth2client.service_account import ServiceAccountCredentials
import streamlit as st
import json
import os
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os
from cryptography.fernet import Fernet
from database import sync_json_to_sqlite
from gaji import main
from gaji import buka_spreadsheet

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


def duplicate_and_add_sheet(spreadsheet_id, new_sheet_name):
    try:
        # Autentikasi menggunakan kredensial
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        # Menggunakan kredensial yang telah didekripsi
        creds_json = decrypt_json("izingoogle_encrypted.json")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
        service = build('sheets', 'v4', credentials=creds)

        # Ambil informasi spreadsheet
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        # Ambil sheet ID dari sheet pertama
        first_sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']

        # Duplikasikan sheet pertama
        duplicate_request = {
            "requests": [
                {
                    "duplicateSheet": {
                        "sourceSheetId": first_sheet_id,
                        "newSheetName": new_sheet_name
                    }
                }
            ]
        }

        response = service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body=duplicate_request
        ).execute()

        # Ambil ID sheet baru dari respons
        new_sheet_id = response['replies'][0]['duplicateSheet']['properties']['sheetId']

        # Geser sheet baru ke indeks terakhir (kanan)
        move_request = {
            "requests": [
                {
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": new_sheet_id,
                            "index": len(spreadsheet['sheets'])  # Pindah ke indeks terakhir
                        },
                        "fields": "index"
                    }
                }
            ]
        }

        service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body=move_request
        ).execute()

        # Tambahkan informasi sheet ke JSON
        sheets_json = "id_mapping.json"
        if not os.path.exists(sheets_json):
            with open(sheets_json, "w") as file:
                json.dump({}, file)

        with open(sheets_json, "r") as file:
            data = json.load(file)

        # Tambahkan atau perbarui informasi sheet
        data[new_sheet_name] = str(new_sheet_id)

        with open(sheets_json, "w") as file:
            json.dump(data, file, indent=4)

        st.success(f"Sheet '{new_sheet_name}' berhasil dibuat dan diarahkan ke kanan!")

    except HttpError as err:
        st.error(f"Terjadi kesalahan HTTP: {err}")
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

def create_user_and_duplicate_sheet_page():
    st.title("Buat dan Duplikat Sheet, Perbarui JSON")

    # Input pengguna
    spreadsheet_id = st.text_input("Masukkan ID Spreadsheet:", "1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps")
    new_sheet_name = st.text_input("Masukkan Nama Sheet Baru:", "Sheet Baru")
    new_sheet_id = st.text_input("Masukkan ID Sheet Baru (Opsional):", "")

    if st.button("Proses Buat Sheet dan Perbarui JSON"):
        try:
            sheets_json = "id_mapping.json"

            # Jika file belum ada, buat file kosong
            if not os.path.exists(sheets_json):
                with open(sheets_json, "w") as file:
                    json.dump({}, file)

            # Baca file JSON
            with open(sheets_json, "r") as file:
                data = json.load(file)

            # Cek apakah nama sheet sudah ada dalam JSON (case-sensitive)
            if new_sheet_name in data:
                st.error(f"Sheet dengan nama '{new_sheet_name}' sudah ada dalam JSON. Gunakan nama lain.")
                return  # Hentikan proses jika nama sudah ada

            # Duplikasikan sheet di spreadsheet
            duplicate_and_add_sheet(spreadsheet_id, new_sheet_name)

            # Tambahkan sheet baru ke JSON dengan ID yang diberikan (jika ada)
            if new_sheet_id:
                data[new_sheet_name] = new_sheet_id

                with open(sheets_json, "w") as file:
                    json.dump(data, file, indent=4)

                st.success(f"Sheet '{new_sheet_name}' berhasil dibuat dan ditambahkan ke JSON dengan ID '{new_sheet_id}'!")

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")


def delete_sheet(spreadsheet_id, sheet_name):
    try:
        # Lokasi file JSON
        sheets_json = "id_mapping.json"

        # Pastikan file JSON ada
        if not os.path.exists(sheets_json):
            st.error("File JSON tidak ditemukan. Tidak ada sheet yang dapat dihapus.")
            return

        # Baca data dari JSON
        with open(sheets_json, "r") as file:
            data = json.load(file)

            sheet_to_delete = next(
            (sheet for sheet in spreadsheet['sheets'] if sheet['properties']['title'].strip().lower() == sheet_name.strip().lower()),
            None
             )

            st.error(f"Sheet '{sheet_name}' tidak ditemukan di JSON.")
            return

        # Autentikasi
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        # Menggunakan kredensial yang telah didekripsi
        creds_json = decrypt_json("izingoogle_encrypted.json")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
        service = build('sheets', 'v4', credentials=creds)

        # Ambil informasi spreadsheet
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        # Cari sheet ID berdasarkan nama sheet
        sheet_to_delete = next(
            (sheet for sheet in spreadsheet['sheets'] if sheet['properties']['title'] == sheet_name),
            None
        )

        if sheet_to_delete:
            # Hapus sheet dari spreadsheet
            delete_request = {
                "requests": [
                    {
                        "deleteSheet": {
                            "sheetId": sheet_to_delete['properties']['sheetId']
                        }
                    }
                ]
            }
            service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body=delete_request
            ).execute()

            # Hapus sheet dari JSON
            del data[sheet_name]
            with open(sheets_json, "w") as file:
                json.dump(data, file, indent=4)

            st.success(f"Sheet '{sheet_name}' berhasil dihapus dari spreadsheet dan JSON!")
        else:
            st.error(f"Sheet '{sheet_name}' tidak ditemukan di spreadsheet.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Streamlit UI
def delete_sheet_ui():
    st.title("Hapus Sheet dari Spreadsheet")

    # Lokasi file JSON
    sheets_json = "id_mapping.json"

    # Pastikan file JSON ada
    if not os.path.exists(sheets_json):
        st.error("File JSON tidak ditemukan. Tidak ada sheet yang dapat dihapus.")
        return

    # Baca data dari JSON
    with open(sheets_json, "r") as file:
        data = json.load(file)

    # Input untuk spreadsheet ID
    spreadsheet_id = st.text_input("Masukkan ID Spreadsheet:", "1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps")

    # Dropdown untuk memilih sheet
    sheet_name = st.selectbox("Pilih Sheet yang Akan Dihapus:", options=list(data.keys()))

    if st.button("Hapus Sheet"):
        if spreadsheet_id and sheet_name:
            delete_sheet(spreadsheet_id, sheet_name)
        else:
            st.error("Harap masukkan ID Spreadsheet dan pilih sheet yang akan dihapus.")



# Halaman untuk melihat daftar sheet
def view_users_and_sheets():
    st.title("Lihat Daftar Sheet")

    users_json = "id_mapping.json"
    if os.path.exists(users_json):
        with open(users_json, "r") as file:
            data = json.load(file)

        st.subheader("Daftar Sheet dan ID")
        for sheet_name, sheet_id in data.items():
            st.write(f"Nama Sheet: {sheet_name}, ID Sheet: {sheet_id}")

    else:
        st.write("Belum ada sheet yang dibuat.")


def assign_new_id(spreadsheet_id):
    try:
        # Lokasi file JSON
        sheets_json = "id_mapping.json"

        # Pastikan file JSON ada
        if not os.path.exists(sheets_json):
            st.error("File JSON tidak ditemukan.")
            return

        # Baca data dari JSON
        with open(sheets_json, "r") as file:
            data = json.load(file)

        # Autentikasi
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
            "https://www.googleapis.com/auth/spreadsheets"
        ]

        # Menggunakan kredensial yang telah didekripsi
        creds_json = decrypt_json("izingoogle_encrypted.json")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_json, scope)
        service = build('sheets', 'v4', credentials=creds)

        # Ambil informasi spreadsheet
        spreadsheet = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()

        # Daftar sheet yang sudah memiliki ID dan yang belum
        sheets_with_id = []
        sheets_without_id = []

        # Loop melalui semua sheet di spreadsheet
        for sheet in spreadsheet['sheets']:
            sheet_name = sheet['properties']['title']

            if sheet_name in data:
                sheets_with_id.append((sheet_name, data[sheet_name]))  # (nama_sheet, id)
            else:
                sheets_without_id.append(sheet_name)  # nama_sheet

        # Tampilkan sheet yang sudah memiliki ID
        st.subheader("Sheet yang Sudah Memiliki ID:")
        if sheets_with_id:
            for sheet_name, sheet_id in sheets_with_id:
                st.write(f"- {sheet_name}: {sheet_id}")
        else:
            st.info("Tidak ada sheet yang sudah memiliki ID.")

        # Tampilkan sheet yang belum memiliki ID
        st.subheader("Sheet yang Belum Memiliki ID:")
        if sheets_without_id:
            # Pilih sheet yang akan diberi ID
            selected_sheet = st.selectbox(
                "Pilih Sheet yang Akan Diberi ID Baru:",
                sheets_without_id
            )

            # Input ID baru
            new_id = st.text_input(f"Masukkan ID Baru untuk {selected_sheet}:")

            if st.button(f"Tambahkan ID Baru untuk {selected_sheet}"):
                if new_id:
                    data[selected_sheet] = new_id

                    # Tulis kembali ke JSON
                    with open(sheets_json, "w") as file:
                        json.dump(data, file, indent=4)

                    st.success(f"Sheet '{selected_sheet}' diberikan ID baru: '{new_id}'")
                    st.rerun()  # Refresh untuk memperbarui tampilan
                else:
                    st.error("Harap masukkan ID baru.")
        else:
            st.info("Tidak ada sheet yang belum memiliki ID.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Fungsi untuk menghasilkan ID baru (contoh sederhana)
def generate_new_id():
    import uuid
    return str(uuid.uuid4())

# Streamlit UI untuk memberikan ID baru
def assign_new_id_ui():
    st.title("Kelola ID Sheet di Spreadsheet")

    # Input untuk spreadsheet ID (tetap diperlukan untuk autentikasi)
    spreadsheet_id = "1LrO71Y5afiKH98gHPsDN9G7p2YUiP1PPf-tIs9tSmps"  # Ganti dengan ID spreadsheet Anda

    # Langsung tampilkan sheet yang sudah memiliki ID dan yang belum
    assign_new_id(spreadsheet_id)


def update_sheet_id(sheet_name, new_id):
    try:
        # Lokasi file JSON
        sheets_json = "id_mapping.json"

        # Pastikan file JSON ada
        if not os.path.exists(sheets_json):
            st.error("File JSON tidak ditemukan.")
            return

        # Baca data dari JSON
        with open(sheets_json, "r") as file:
            data = json.load(file)

        # Periksa apakah sheet_name ada di JSON
        if sheet_name in data:
            # Update ID
            data[sheet_name] = new_id

            # Tulis kembali ke JSON
            with open(sheets_json, "w") as file:
                json.dump(data, file, indent=4)

            st.success(f"ID untuk sheet '{sheet_name}' berhasil diubah menjadi '{new_id}'.")
        else:
            st.error(f"Sheet '{sheet_name}' tidak ditemukan di JSON.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")

# Streamlit UI untuk mengganti ID
def update_sheet_id_ui():
    st.title("Update ID Sheet di JSON")

    # Lokasi file JSON
    sheets_json = "id_mapping.json"

    # Pastikan file JSON ada
    if not os.path.exists(sheets_json):
        st.error("File JSON tidak ditemukan.")
        return

    # Baca data dari JSON
    with open(sheets_json, "r") as file:
        data = json.load(file)

    # Input untuk sheet name dan new ID
    sheet_name = st.selectbox("Pilih Sheet yang Akan Diupdate:", options=list(data.keys()))
    new_id = st.text_input("Masukkan ID Baru:")

    if st.button("Update ID"):
        if sheet_name and new_id:
            update_sheet_id(sheet_name, new_id)
        else:
            st.error("Harap pilih sheet dan masukkan ID baru.")



# Fungsi utama
def admin_dashboard():
    st.sidebar.title("Menu")
    option  = st.sidebar.selectbox("Pilih Menu", ["Buat Pengguna dan Duplikat Sheet", "Lihat Pengguna dan Sheet", "Hapus Sheet","mengganti ID","Memberikan ID Kosong Pada Sheet","Anlisi Gaji Karyawan"])

    st.write(f"Anda memilih: {option}")
    if option == "Buat Pengguna dan Duplikat Sheet":
        create_user_and_duplicate_sheet_page()
    elif option == "Lihat Pengguna dan Sheet":
        view_users_and_sheets()
    elif option == "Hapus Sheet":
        delete_sheet_ui()
    elif option == "mengganti ID":
        update_sheet_id_ui()
    elif option == "Memberikan ID Kosong Pada Sheet":
        assign_new_id_ui()
    elif option == "Anlisi Gaji Karyawan":
        main()
    
    # **Tombol Logout**
    if st.button("Logout"):
        del st.session_state["is_admin"]  # Hapus status login admin
        st.success("Berhasil logout. Kembali ke halaman login...")
        st.rerun()  # Refresh halaman untuk kembali ke login

if __name__ == "__main__":
    if "is_admin" not in st.session_state or not st.session_state["is_admin"]:
        st.warning("Silakan login terlebih dahulu sebagai admin.")
    else:
        buka_spreadsheet()
        sync_json_to_sqlite()
        admin_dashboard()
