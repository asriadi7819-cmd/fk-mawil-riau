import streamlit as st
import pandas as pd
import sqlite3
import os
import base64
from datetime import datetime, date
import random
import io
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from streamlit_cookies_controller import CookieController
import cloudinary
import cloudinary.uploader
import cloudinary.api

# --- KONFIGURASI CLOUDINARY ---
cloudinary.config(
    cloud_name="dexhqltm",
    api_key="585645177582163",
    api_secret="JFQYgD_hErHRikJLQDPoVcTUKb8",
    secure=True
)

# Konfigurasi Halaman
st.set_page_config(
   page_title="FK Mawil Riau",
   page_icon="🕌",
   layout="wide"
)

# Inisialisasi Cookie Controller
cookie_manager = CookieController()

st.markdown("""
    <style>
    .stRadio div[role="radiogroup"] > label > div:first-child {
        background-color: #0E6655;
    }
    h1 {
        border-bottom: 3px solid #0E6655;
        padding-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# Buat folder penyimpanan file otomatis jika belum ada (opsional fallback lokal)
UPLOAD_DIR = "uploads_foto"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

def get_image_base64(image_path):
    if image_path:
        clean_path = str(image_path).split("|")[0].strip()
        if clean_path and os.path.exists(clean_path):
            with open(clean_path, "rb") as img_file:
                encoded = base64.b64encode(img_file.read()).decode()
                ext = clean_path.split('.')[-1].lower()
                mime = 'image/png' if ext == 'png' else 'image/jpeg'
                return f"data:{mime};base64,{encoded}"
    return None

DAFTAR_KAB_KOTA = [
    "Pekanbaru", "Dumai", "Rokan Hilir", "Bengkalis", 
    "Kampar", "Siak", "Pelalawan", "Indragiri Hulu", 
    "Indragiri Hilir", "Kuantan Singingi", "Kepulauan Meranti", "Rokan Hulu"
]

# --- INISIALISASI DATABASE SQLITE ---
def init_db():
    conn = sqlite3.connect('fk_mawil_riau.db')
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS anggota (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            sub_mawil TEXT,
            jenis_kelamin TEXT,
            alamat TEXT,
            status TEXT,
            letnan_ijazah TEXT,
            tanggal_ijazah TEXT,
            kontak TEXT,
            foto TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS struktur_pengurus (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            jabatan TEXT UNIQUE,
            nama_pejabat TEXT,
            kontak TEXT,
            foto TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS keuangan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tanggal TEXT,
            jenis TEXT,
            sub_mawil TEXT,
            jumlah REAL,
            keterangan TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS rekening_tujuan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_bank TEXT,
            nomor_rekening TEXT,
            atas_nama TEXT,
            keterangan TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cashflow_transaksi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kategori TEXT,
            pengirim TEXT,
            sub_mawil TEXT,
            tanggal TEXT,
            jumlah REAL,
            jenis_arus TEXT,
            keterangan TEXT,
            bukti_transfer TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pengumuman (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            waktu TEXT,
            judul TEXT,
            isi TEXT,
            pembuat TEXT,
            foto_pengumuman TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS galeri_umum (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            penulis TEXT,
            sub_mawil TEXT,
            waktu TEXT,
            konten TEXT,
            foto_galeri TEXT,
            tipe TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS galeri_resmi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            judul TEXT,
            kategori TEXT,
            waktu TEXT,
            foto_resmi TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reaksi_posting (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id INTEGER,
            nama_sanfk TEXT,
            reaksi TEXT,
            UNIQUE(post_id, nama_sanfk)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS reaksi_resmi (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id INTEGER,
            nama_sanfk TEXT,
            reaksi TEXT,
            UNIQUE(post_id, nama_sanfk)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS komentar_posting (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            post_id INTEGER,
            nama_sanfk TEXT,
            waktu TEXT,
            komentar TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS jadwal_rutinan (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sub_mawil TEXT,
            tanggal TEXT,
            keterangan TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS rsvp (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            sub_mawil TEXT,
            tanggal TEXT,
            status_rsvp TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS aktual_hadir (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            sub_mawil TEXT,
            tanggal TEXT,
            waktu_input TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS jadwal_kopdar_baksos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kategori TEXT,
            sub_mawil_tuan_rumah TEXT,
            tanggal TEXT,
            lokasi TEXT,
            nominal_dana REAL,
            keterangan TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS rsvp_kopdar_baksos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            sub_mawil TEXT,
            jadwal_id INTEGER,
            status_rsvp TEXT
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS aktual_hadir_kopdar_baksos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama TEXT,
            sub_mawil TEXT,
            jadwal_id INTEGER,
            waktu_input TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS token_pendaftaran (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nama_sanfk TEXT,
            token TEXT UNIQUE,
            status_pakai TEXT DEFAULT 'Belum'
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS pengaturan_captcha (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            role_pengurus TEXT,
            nama_sanfk TEXT UNIQUE,
            kode_captcha TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT,
            password TEXT,
            role TEXT,
            sub_wilayah TEXT,
            no_hp TEXT,
            nama_sanfk TEXT
        )
    ''')
    
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN no_hp TEXT")
    except sqlite3.OperationalError:
        pass 

    try:
        cursor.execute("ALTER TABLE users ADD COLUMN sub_wilayah TEXT")
    except sqlite3.OperationalError:
        pass 

    try:
        cursor.execute("ALTER TABLE users ADD COLUMN nama_sanfk TEXT")
    except sqlite3.OperationalError:
        pass 

    try:
        cursor.execute("ALTER TABLE pengaturan_captcha ADD COLUMN nama_sanfk TEXT")
    except sqlite3.OperationalError:
        pass

    conn.commit()

    jabatan_default = [
        "Ketua Mawil", 
        "Bendahara Mawil", 
        "Sekretaris Mawil", 
        "Admin Dokumentasi Mawil"
    ] + [f"Ketua Sub Mawil - {kab}" for kab in DAFTAR_KAB_KOTA]
    
    for jab in jabatan_default:
        cursor.execute("INSERT OR IGNORE INTO struktur_pengurus (jabatan, nama_pejabat, kontak, foto) VALUES (?, ?, ?, ?)", (jab, "(Belum Ditetapkan)", "-", ""))

    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        default_users = [
            ("superadmin", "super123", "Superadmin", "-", "-", "Superadmin"),
            ("admin", "admin123", "Ketua Mawil", "-", "-", "Ketua Mawil"),
            ("bendahara", "bendahara123", "Bendahara Mawil", "-", "-", "Bendahara Mawil"),
            ("sekretaris", "sekretaris123", "Sekretaris Mawil", "-", "-", "Sekretaris Mawil"),
            ("dokumentasi", "dok123", "Admin Dokumentasi Mawil", "-", "-", "Admin Dokumentasi Mawil"),
            ("ketua_sub", "sub123", "Ketua Sub Mawil", "Pekanbaru", "-", "Ketua Sub Pekanbaru"),
            ("sanfk", "sanfk123", "SanFK", "-", "081234567890", "SanFK Default")
        ]
        cursor.executemany("INSERT OR IGNORE INTO users (username, password, role, sub_wilayah, no_hp, nama_sanfk) VALUES (?, ?, ?, ?, ?, ?)", default_users)

    conn.commit()
    conn.close()

init_db()

def get_data(query, params=()):
    conn = sqlite3.connect('fk_mawil_riau.db')
    df = pd.read_sql(query, conn, params=params)
    conn.close()
    return df

def execute_query(query, params=()):
    conn = sqlite3.connect('fk_mawil_riau.db')
    cursor = conn.cursor()
    cursor.execute(query, params)
    conn.commit()
    conn.close()

# --- AMBIL COOKIES ---
cookie_logged_in = cookie_manager.get("fk_logged_in")
cookie_username = cookie_manager.get("fk_username")
cookie_role = cookie_manager.get("fk_role")
cookie_nama_sanfk = cookie_manager.get("fk_nama_sanfk")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = True if cookie_logged_in == "True" else False
if "username" not in st.session_state:
    st.session_state.username = cookie_username if cookie_username else ""
if "role" not in st.session_state:
    st.session_state.role = cookie_role if cookie_role else ""
if "nama_sanfk" not in st.session_state:
    st.session_state.nama_sanfk = cookie_nama_sanfk if cookie_nama_sanfk else ""

if "gen_captcha_code" not in st.session_state:
    st.session_state.gen_captcha_code = f"mwRIAU-{random.randint(1000, 9999)}"

role = st.session_state.role

# Tentukan sanfk_aktif_terpilih secara global atau default jika session sudah ada
sanfk_aktif_terpilih = st.session_state.nama_sanfk if "nama_sanfk" in st.session_state else ""

# --- SIDEBAR: FORM LOGIN, REGISTER ATAU MENU UTAMA ---
st.sidebar.title("🕌 FK MAWIL RIAU")
st.sidebar.markdown("**Forum Silaturrahmi SanFK**")
st.sidebar.divider()

if not st.session_state.logged_in:
    auth_mode = st.sidebar.radio("Pilih Mode", ["Login", "Daftar Akun Baru"])
    
    if auth_mode == "Login":
        st.sidebar.subheader("🔑 Silakan Login")
        login_username = st.sidebar.text_input("Username")
        login_password = st.sidebar.text_input("Password", type="password")
        
        if st.sidebar.button("Masuk", use_container_width=True):
            conn = sqlite3.connect('fk_mawil_riau.db')
            cursor = conn.cursor()
            cursor.execute("SELECT password, role, nama_sanfk FROM users WHERE username = ?", (login_username,))
            user_row = cursor.fetchone()
            conn.close()
            
            if user_row and user_row[0] == login_password:
                st.session_state.logged_in = True
                st.session_state.username = login_username
                st.session_state.role = user_row[1]
                st.session_state.nama_sanfk = user_row[2] if user_row[2] else login_username
                
                cookie_manager.set('fk_logged_in', 'True', max_age=2592000)
                cookie_manager.set('fk_username', login_username, max_age=2592000)
                cookie_manager.set('fk_role', user_row[1], max_age=2592000)
                cookie_manager.set('fk_nama_sanfk', st.session_state.nama_sanfk, max_age=2592000)
                
                st.success("Login berhasil!")
                st.rerun()
            else:
                st.sidebar.error("Username atau Password salah!")
                
    else:
        st.sidebar.subheader("📝 Pendaftaran Akun Baru")
        reg_role = st.sidebar.selectbox("Pilih Role Pendaftaran", [
            "SanFK", "Ketua Mawil", "Bendahara Mawil", 
            "Sekretaris Mawil", "Admin Dokumentasi Mawil", "Ketua Sub Mawil"
        ])
        
        reg_sub_wilayah = "-"
        if reg_role == "Ketua Sub Mawil":
            reg_sub_wilayah = st.sidebar.selectbox("Pilih Kabupaten/Kota Sub Mawil", DAFTAR_KAB_KOTA)
        
        df_anggota_reg = get_data("SELECT nama, kontak FROM anggota")
        if not df_anggota_reg.empty:
            list_nama_sanfk = df_anggota_reg['nama'].tolist()
            selected_nama_sanfk = st.sidebar.selectbox("Pilih Nama Anda (SanFK)", list_nama_sanfk)
            
            selected_row = df_anggota_reg[df_anggota_reg['nama'] == selected_nama_sanfk].iloc[0]
            db_hp = str(selected_row['kontak']) if pd.notna(selected_row['kontak']) else "-"
        else:
            st.sidebar.warning("Belum ada data SanFK di Manajemen SanFK. Harap input data anggota terlebih dahulu.")
            selected_nama_sanfk = ""
            db_hp = "-"
            
        reg_username = st.sidebar.text_input("Ketik Username (Untuk Login)")
        
        pengurus_inti_list = ["Ketua Mawil", "Bendahara Mawil", "Sekretaris Mawil", "Admin Dokumentasi Mawil"]
        reg_role_captcha_input = "-"
        if reg_role in pengurus_inti_list:
            reg_role_captcha_input = st.sidebar.text_input(f"Validasi Kode Khusus Role ({reg_role})")
        elif reg_role == "SanFK":
            reg_role_captcha_input = st.sidebar.text_input("Validasi Kode Verifikasi/Captcha SanFK")
            
        reg_password = st.sidebar.text_input("Buat Password", type="password")
            
        if st.sidebar.button("Daftar Sekarang", use_container_width=True):
            if not reg_username or not reg_password:
                st.sidebar.warning("Username dan Password wajib diisi!")
            else:
                conn = sqlite3.connect('fk_mawil_riau.db')
                cursor = conn.cursor()
                try:
                    cursor.execute("SELECT COUNT(*) FROM users WHERE username = ?", (reg_username.strip(),))
                    if cursor.fetchone()[0] > 0:
                        st.sidebar.error(f"Gagal! Username '{reg_username.strip()}' sudah digunakan.")
                        conn.close()
                        st.stop()
                    
                    if reg_role in pengurus_inti_list:
                        cursor.execute("SELECT COUNT(*) FROM users WHERE role = ?", (reg_role,))
                        if cursor.fetchone()[0] > 0:
                            st.sidebar.error(f"Gagal! Role **{reg_role}** sudah terdaftar dalam sistem dan hanya boleh ada 1 orang.")
                            conn.close()
                            st.stop()
                    
                    if reg_role == "Ketua Sub Mawil":
                        cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'Ketua Sub Mawil' AND sub_wilayah = ?", (reg_sub_wilayah,))
                        if cursor.fetchone()[0] > 0:
                            st.sidebar.error(f"Gagal! Ketua Sub Mawil untuk wilayah **{reg_sub_wilayah}** sudah terdaftar dan hanya boleh ada 1 orang.")
                            conn.close()
                            st.stop()

                    if reg_role in pengurus_inti_list:
                        cursor.execute("SELECT kode_captcha FROM pengaturan_captcha WHERE role_pengurus = ? AND nama_sanfk = ?", (reg_role, f"ROLE_{reg_role}"))
                        res_cap = cursor.fetchone()
                        saved_captcha = res_cap[0] if res_cap else ""
                        
                        if reg_role_captcha_input.strip() != saved_captcha:
                            st.sidebar.error(f"Kode Captcha untuk role **{reg_role}** tidak valid!")
                            conn.close()
                            st.stop()
                    elif reg_role == "SanFK":
                        cursor.execute("SELECT kode_captcha FROM pengaturan_captcha WHERE nama_sanfk = ?", (selected_nama_sanfk,))
                        res_cap_sanfk = cursor.fetchone()
                        saved_captcha_sanfk = res_cap_sanfk[0] if res_cap_sanfk else ""
                        
                        if saved_captcha_sanfk and reg_role_captcha_input.strip() != saved_captcha_sanfk:
                            st.sidebar.error("Kode Verifikasi / Captcha SanFK tidak valid!")
                            conn.close()
                            st.stop()
                    
                    reg_hp = db_hp
                    
                    cursor.execute("INSERT INTO users (username, password, role, sub_wilayah, no_hp, nama_sanfk) VALUES (?, ?, ?, ?, ?, ?)", 
                                   (reg_username.strip(), reg_password, reg_role, reg_sub_wilayah, reg_hp, selected_nama_sanfk))
                    conn.commit()
                    conn.close()
                    st.sidebar.success("Pendaftaran berhasil! Silakan pindah ke tab Login.")
                except sqlite3.OperationalError as e:
                    conn.close()
                    st.sidebar.error(f"Terjadi kesalahan database: {e}")
                finally:
                    if conn:
                        conn.close()
                
    st.stop()

st.sidebar.success(f"Masuk sebagai: **{st.session_state.nama_sanfk}** ({role})")
if st.sidebar.button("Keluar (Logout)", use_container_width=True):
    cookie_manager.set('fk_logged_in', 'False', max_age=0)
    cookie_manager.set('fk_username', '', max_age=0)
    cookie_manager.set('fk_role', '', max_age=0)
    cookie_manager.set('fk_nama_sanfk', '', max_age=0)
    
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.role = ""
    st.session_state.nama_sanfk = ""
    st.rerun()

st.sidebar.divider()

list_menu = [
    "Galeri & Feed Umum",
    "Beranda & Pengumuman",
    "Manajemen SanFK & KTA",
    "Agenda & Rutinan Dzikir",
    "Agenda Kopdar Mawil & Baksos",
    "Keuangan & Kotak Hijau",
    "Layanan Santunan & Kontak",
    "Galeri Resmi (Admin)",
    "Manajemen Akun & Role"
]

if role != "Superadmin":
    if role not in ["SanFK", "Sekretaris Mawil"]:
       list_menu = [m for m in list_menu if m != "Galeri & Feed Umum"]

    if role in ["Ketua Sub Mawil", "Admin Dokumentasi Mawil"]:
       list_menu = [m for m in list_menu if m != "Beranda & Pengumuman"]

    if role in ["Ketua Mawil", "Bendahara Mawil", "Ketua Sub Mawil", "Admin Dokumentasi Mawil"]:
       list_menu = [m for m in list_menu if m != "Manajemen SanFK & KTA"]

    if role in ["Ketua Mawil", "Bendahara Mawil", "Sekretaris Mawil", "Admin Dokumentasi Mawil"]:
       list_menu = [m for m in list_menu if m != "Agenda & Rutinan Dzikir"]

    if role in ["Bendahara Mawil", "Sekretaris Mawil", "Ketua Sub Mawil", "Admin Dokumentasi Mawil"]:
       list_menu = [m for m in list_menu if m != "Agenda Kopdar Mawil & Baksos"]

    if role not in ["Bendahara Mawil", "SanFK"]:
       list_menu = [m for m in list_menu if m != "Keuangan & Kotak Hijau"]

    if role != "SanFK":
       list_menu = [m for m in list_menu if m != "Layanan Santunan & Kontak"]

    if role not in ["SanFK", "Admin Dokumentasi Mawil"]:
       list_menu = [m for m in list_menu if m != "Galeri Resmi (Admin)"]
       
    list_menu = [m for m in list_menu if m != "Manajemen Akun & Role"]

if not list_menu:
    list_menu = ["Beranda & Pengumuman"]

menu = st.sidebar.radio("Navigasi Menu", list_menu)

# --- HALAMAN KHUSUS SUPERADMIN: MANAJEMEN AKUN & ROLE ---
if menu == "Manajemen Akun & Role" and role == "Superadmin":
    st.title("🛡️ Manajemen Akun & Role (Superadmin)")
    st.markdown("Kelola daftar akun pengguna, hak akses role, nomor HP verifikasi SanFK, pengaturan token pendaftaran, pengaturan captcha per SanFK/Role, serta backup dan import database.")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Daftar Akun", "Tambah Akun Baru", "Token Pendaftaran (Admin)", "Pengaturan Kode Verifikasi/Captcha", "Backup & Import Database"])
    
    with tab1:
        st.subheader("Daftar Pengguna Sistem")
        df_users = get_data("SELECT id, username, role, sub_wilayah, no_hp, nama_sanfk FROM users")
        st.dataframe(df_users, use_container_width=True)
        
        st.divider()
        st.subheader("Edit atau Hapus Akun")
        selected_user_id = st.selectbox("Pilih ID / User:", df_users['id'].tolist() if not df_users.empty else [])
        
        if not df_users.empty:
            user_row = df_users[df_users['id'] == selected_user_id].iloc[0]
            st.write(f"Mengedit User: **{user_row['username']}** (Nama SanFK: {user_row['nama_sanfk']}, Role: {user_row['role']})")
            
            edit_pass = st.text_input("Password Baru (kosongkan jika tetap)", type="password", key="edit_pass")
            edit_role = st.selectbox("Ubah Role", [
                "Superadmin", "Ketua Mawil", "Bendahara Mawil", 
                "Sekretaris Mawil", "Admin Dokumentasi Mawil", 
                "Ketua Sub Mawil", "SanFK"
            ], index=["Superadmin", "Ketua Mawil", "Bendahara Mawil", "Sekretaris Mawil", "Admin Dokumentasi Mawil", "Ketua Sub Mawil", "SanFK"].index(user_row['role']))
            
            edit_hp = st.text_input("Nomor HP / WhatsApp", value=str(user_row['no_hp'] if user_row['no_hp'] else "-"))
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Simpan Perubahan"):
                    if edit_pass:
                        execute_query("UPDATE users SET password = ?, role = ?, no_hp = ? WHERE id = ?", 
                                      (edit_pass, edit_role, edit_hp, selected_user_id))
                    else:
                        execute_query("UPDATE users SET role = ?, no_hp = ? WHERE id = ?", 
                                      (edit_role, edit_hp, selected_user_id))
                    st.success("Akun berhasil diperbarui!")
                    st.rerun()
            with col2:
                if st.button("Hapus Akun Ini", type="primary"):
                    execute_query("DELETE FROM users WHERE id = ?", (selected_user_id,))
                    st.success("Akun berhasil dihapus!")
                    st.rerun()

    with tab2:
        st.subheader("Tambah Akun Pengguna Baru oleh Superadmin")
        with st.form("form_tambah_akun"):
            new_username = st.text_input("Username Baru (Untuk Login)")
            
            df_anggota_super = get_data("SELECT nama FROM anggota")
            list_sanfk_super = df_anggota_super['nama'].tolist() if not df_anggota_super.empty else ["Belum ada data SanFK"]
            new_pilih_sanfk = st.selectbox("Pilih Nama dari Data SanFK", list_sanfk_super)
            
            new_password = st.text_input("Password", type="password")
            new_role_input = st.selectbox("Role Akses", [
                "Superadmin", "Ketua Mawil", "Bendahara Mawil", 
                "Sekretaris Mawil", "Admin Dokumentasi Mawil", 
                "Ketua Sub Mawil", "SanFK"
            ])
            new_sub_wil = st.selectbox("Sub Wilayah (Jika Ketua Sub Mawil)", ["-"] + DAFTAR_KAB_KOTA)
            new_hp_input = st.text_input("Nomor HP / WhatsApp (Misal: 0812...)")
            
            submitted = st.form_submit_button("Buat Akun")
            if submitted:
                if new_username and new_password:
                    try:
                        hp_val = new_hp_input if new_hp_input else "-"
                        sub_wil_val = new_sub_wil if new_sub_wil else "-"
                        execute_query("INSERT INTO users (username, password, role, sub_wilayah, no_hp, nama_sanfk) VALUES (?, ?, ?, ?, ?, ?)", 
                                      (new_username, new_password, new_role_input, sub_wil_val, hp_val, new_pilih_sanfk))
                        st.success(f"Akun '{new_username}' (Nama SanFK: {new_pilih_sanfk}) dengan role '{new_role_input}' berhasil ditambahkan!")
                    except sqlite3.OperationalError:
                        st.error("Terjadi kesalahan saat menambahkan akun.")
                else:
                    st.warning("Username dan Password wajib diisi!")

    with tab3:
        st.subheader("🔑 Kelola Token Pendaftaran SanFK")
        st.markdown("Buat dan bagikan kode token rahasia ke masing-masing anggota SanFK agar mereka dapat melakukan registrasi akun.")
        
        with st.form("form_buat_token"):
            df_anggota_t = get_data("SELECT nama FROM anggota")
            list_sanfk_t = df_anggota_t['nama'].tolist() if not df_anggota_t.empty else []
            pilih_sanfk_token = st.selectbox("Pilih Nama Anggota SanFK", list_sanfk_t)
            
            default_gen_token = f"FK-{random.randint(100000, 999999)}"
            input_token_val = st.text_input("Kode Token Pendaftaran", value=default_gen_token)
            
            btn_buat_token = st.form_submit_button("Simpan & Terbitkan Token")
            if btn_buat_token:
                if pilih_sanfk_token and input_token_val:
                    try:
                        execute_query("INSERT INTO token_pendaftaran (nama_sanfk, token, status_pakai) VALUES (?, ?, 'Belum')", (pilih_sanfk_token, input_token_val.strip()))
                        st.success(f"Token untuk **{pilih_sanfk_token}** berhasil dibuat: **{input_token_val.strip()}**")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("Token tersebut sudah ada di database, gunakan kode yang berbeda.")
                else:
                    st.warning("Pilih nama dan isi token terlebih dahulu!")
                    
        st.divider()
        st.markdown("### Daftar Token Pendaftaran Aktif")
        df_list_token = get_data("SELECT id, nama_sanfk AS 'Nama SanFK', token AS 'Kode Token', status_pakai AS 'Status Digunakan' FROM token_pendaftaran")
        if not df_list_token.empty:
            st.dataframe(df_list_token, use_container_width=True)
            
            token_id_hapus = st.selectbox("Pilih ID Token yang ingin dihapus:", df_list_token['id'].tolist())
            if st.button("Hapus Token Terpilih"):
                execute_query("DELETE FROM token_pendaftaran WHERE id = ?", (token_id_hapus,))
                st.success("Token berhasil dihapus!")
                st.rerun()
        else:
            st.info("Belum ada token yang diterbitkan.")

    with tab4:
        st.subheader("🔑 Pengaturan Kode Verifikasi / Captcha per SanFK & Role")
        st.markdown("Atur kode verifikasi/captcha unik untuk masing-masing anggota SanFK secara individual atau berdasarkan role pengurus.")
        
        df_anggota_cap = get_data("SELECT nama FROM anggota")
        list_sanfk_cap = df_anggota_cap['nama'].tolist() if not df_anggota_cap.empty else []
        
        pilih_tipe_atur = st.selectbox("Kategori Pengaturan", ["Berdasarkan Nama SanFK", "Berdasarkan Role Pengurus"])
        
        target_sanfk = "-"
        target_role_p = "-"
        
        if pilih_tipe_atur == "Berdasarkan Nama SanFK":
            if list_sanfk_cap:
                target_sanfk = st.selectbox("Pilih Nama SanFK", list_sanfk_cap)
            else:
                st.warning("Belum ada data anggota SanFK.")
        else:
            target_role_p = st.selectbox("Pilih Role Pengurus", [
                "Ketua Mawil", "Sekretaris Mawil", "Bendahara Mawil", "Admin Dokumentasi Mawil"
            ])
            
        col_gen1, col_gen2 = st.columns([3, 1])
        with col_gen2:
            st.write("") 
            st.write("") 
            if st.button("🎲 Generate Kode Acak", use_container_width=True):
                if pilih_tipe_atur == "Berdasarkan Role Pengurus":
                    st.session_state.gen_captcha_code = f"mwRIAU-{random.randint(1000, 9999)}"
                else:
                    st.session_state.gen_captcha_code = f"SanFK-{random.randint(1000, 9999)}"
                st.rerun()

        with st.form("form_tambah_edit_captcha"):
            input_kode_baru = st.text_input("Masukkan Kode Verifikasi / Captcha Baru", value=st.session_state.gen_captcha_code)
            
            btn_simpan_cap = st.form_submit_button("Simpan Kode Verifikasi")
            if btn_simpan_cap:
                conn = sqlite3.connect('fk_mawil_riau.db')
                cursor = conn.cursor()
                if pilih_tipe_atur == "Berdasarkan Nama SanFK" and target_sanfk != "-":
                    cursor.execute("""
                        INSERT INTO pengaturan_captcha (role_pengurus, nama_sanfk, kode_captcha) 
                        VALUES (?, ?, ?) 
                        ON CONFLICT(nama_sanfk) DO UPDATE SET kode_captcha = ?
                    """, ("SanFK", target_sanfk, input_kode_baru.strip(), input_kode_baru.strip()))
                else:
                    cursor.execute("""
                        INSERT INTO pengaturan_captcha (role_pengurus, nama_sanfk, kode_captcha) 
                        VALUES (?, ?, ?) 
                        ON CONFLICT(nama_sanfk) DO UPDATE SET kode_captcha = ?
                    """, (target_role_p, f"ROLE_{target_role_p}", input_kode_baru.strip(), input_kode_baru.strip()))
                conn.commit()
                conn.close()
                st.success("Kode verifikasi berhasil disimpan!")
                st.rerun()

        st.divider()
        st.markdown("### Daftar Kode Verifikasi Aktif")
        df_list_captcha = get_data("SELECT id, role_pengurus AS 'Role / Kategori', nama_sanfk AS 'Nama SanFK', kode_captcha AS 'Kode Verifikasi' FROM pengaturan_captcha")
        if not df_list_captcha.empty:
            st.dataframe(df_list_captcha, use_container_width=True)
            
            with st.form("form_hapus_captcha"):
                cap_id_hapus = st.selectbox("Pilih ID Baris yang ingin dihapus:", df_list_captcha['id'].tolist(), key="del_cap")
                btn_hapus_cap = st.form_submit_button("Hapus Pengaturan Terpilih", type="primary")
                if btn_hapus_cap:
                    execute_query("DELETE FROM pengaturan_captcha WHERE id = ?", (cap_id_hapus,))
                    st.success(f"Data dengan ID {cap_id_hapus} berhasil dihapus!")
                    st.rerun()
        else:
            st.info("Belum ada data pengaturan verifikasi.")

    with tab5:
        st.subheader("💾 Backup & Import Database")
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("### 📥 Backup Database")
            if os.path.exists('fk_mawil_riau.db'):
                with open('fk_mawil_riau.db', 'rb') as db_file:
                    db_bytes = db_file.read()
                    st.download_button(
                        label="Unduh File Database (.db)",
                        data=db_bytes,
                        file_name=f"fk_mawil_riau_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db",
                        mime="application/octet-stream",
                        use_container_width=True
                    )
        with col_b2:
            st.markdown("### 📤 Import / Pulihkan Database")
            uploaded_db = st.file_uploader("Pilih file database (.db)", type=["db"])
            if uploaded_db is not None:
                if st.button("Timpa & Pulihkan Database", type="primary", use_container_width=True):
                    with open("fk_mawil_riau.db", "wb") as f:
                        f.write(uploaded_db.getbuffer())
                    st.success("Database berhasil dipulihkan! Memuat ulang aplikasi...")
                    st.rerun()

df_anggota_all = get_data("SELECT * FROM anggota")
df_keuangan_all = get_data("SELECT * FROM keuangan")
df_cashflow_all = get_data("SELECT * FROM cashflow_transaksi")

# --- 1. BERANDA & PENGUMUMAN ---
if menu == "Beranda & Pengumuman":
    st.title("Selamat Datang di Sistem Informasi FK Mawil Riau")
    st.info(f"Anda sedang masuk sebagai: **{role}**" + (f" ({sanfk_aktif_terpilih})" if role == "SanFK" and sanfk_aktif_terpilih else ""))
    
    col1, col2, col3 = st.columns([1.2, 1, 1.2])
    with col1:
        st.metric("Total Sub Mawil", "12 Kab/Kota")
    with col2:
        st.metric("Total SanFK Aktif", len(df_anggota_all))
    with col3:
        total_kas = df_keuangan_all['jumlah'].sum() if not df_keuangan_all.empty else 0
        st.metric("Total Kas Tercatat", f"Rp {total_kas:,}")

    st.markdown("---")

    tab_beranda_1, tab_beranda_2 = st.tabs(["📢 Pengumuman & Maklumat Resmi", "🏛️ Profil & Struktur Pengurus"])

    with tab_beranda_1:
        st.subheader("📢 Pengumuman & Maklumat Resmi")
        
        if role in ["Ketua Mawil", "Bendahara Mawil", "Sekretaris Mawil"]:
            with st.expander("➕ Buat Pengumuman Baru"):
                jdl = st.text_input("Judul Pengumuman", key="input_jdl_p")
                isi_p = st.text_area("Isi Maklumat / Pesan", key="input_isi_p")
                
                st.markdown("---")
                st.markdown("##### 📁 Lampiran Berbagai File (Bisa Banyak)")
                metode_foto_p = st.radio("Pilih Cara Input File Pengumuman:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key="radio_p")
                
                foto_pengumuman_files = None
                cam_pengumuman = None
                if metode_foto_p == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                    foto_pengumuman_files = st.file_uploader("Pilih File (Bisa Banyak: PDF, MP3, MP4, JPG, PNG, dll)", type=None, accept_multiple_files=True, key="up_p")
                else:
                    cam_pengumuman = st.camera_input("Ambil Foto Langsung dengan Kamera", key="cam_p")
                
                file_access_settings = {}
                if foto_pengumuman_files:
                    st.markdown("##### ⚙️ Atur Hak Akses Masing-Masing File:")
                    for idx_f, f_item in enumerate(foto_pengumuman_files):
                        file_access_settings[f_item.name] = st.selectbox(
                            f"Hak Akses untuk file: **{f_item.name}**", 
                            ["Public", "Private"], 
                            index=1, 
                            key=f"akses_file_new_{idx_f}"
                        )
                elif cam_pengumuman is not None:
                    cam_akses = st.selectbox("Hak Akses untuk Foto Kamera:", ["Public", "Private"], index=1, key="akses_cam_new")
                
                st.caption("ℹ️ **Public**: Bisa didownload oleh semua role. **Private**: Hanya bisa didownload oleh pemosting.")
                
                if st.button("Publikasikan Pengumuman", key="btn_pub_p"):
                    if jdl:
                        path_list = []
                        if foto_pengumuman_files:
                            with st.spinner("Mengunggah file ke Cloudinary..."):
                                for f_item in foto_pengumuman_files:
                                    try:
                                        upload_res = cloudinary.uploader.upload(f_item, resource_type="auto")
                                        f_url = upload_res.get("secure_url")
                                        
                                        akses_dipilih = file_access_settings.get(f_item.name, "Private")
                                        path_list.append(f"{f_url}|{akses_dipilih}")
                                    except Exception as e:
                                        st.error(f"Gagal mengunggah {f_item.name}: {e}")
                            
                        if cam_pengumuman is not None:
                            with st.spinner("Mengunggah foto kamera ke Cloudinary..."):
                                try:
                                    upload_cam = cloudinary.uploader.upload(cam_pengumuman, resource_type="image")
                                    cam_url = upload_cam.get("secure_url")
                                    path_list.append(f"{cam_url}|{cam_akses}")
                                except Exception as e:
                                    st.error(f"Gagal mengunggah foto kamera: {e}")
                                
                        foto_p_str = ",".join(path_list) if path_list else ""
                        waktu_skr = datetime.now().strftime("%Y-%m-%d %H:%M")
                        execute_query("INSERT INTO pengumuman (waktu, judul, isi, pembuat, foto_pengumuman) VALUES (?, ?, ?, ?, ?)", (waktu_skr, jdl, isi_p, role, foto_p_str))
                        st.success("Pengumuman berhasil disiarkan dengan lampiran file di Cloudinary!")
                        st.rerun()
                    else:
                        st.warning("Judul pengumuman wajib diisi!")

            with st.expander("🛠️ Kelola, Edit, atau Hapus Pengumuman Anda"):
                df_kelola_p = get_data("SELECT * FROM pengumuman WHERE pembuat = ? ORDER BY id DESC", (role,))
                
                if df_kelola_p.empty:
                    st.info(f"Belum ada pengumuman yang dibuat oleh **{role}**.")
                else:
                    col_f1, col_f2 = st.columns(2)
                    with col_f1:
                        st.markdown(f"**Pembuat Terkunci:** `{role}`")
                    with col_f2:
                        df_kelola_p['Tanggal_Saja'] = df_kelola_p['waktu'].str.split().str[0]
                        filter_tgl = st.selectbox("Filter Berdasarkan Tanggal", ["Semua Tanggal"] + df_kelola_p['Tanggal_Saja'].unique().tolist())
                    
                    df_filtered = df_kelola_p.copy()
                    if filter_tgl != "Semua Tanggal":
                        df_filtered = df_filtered[df_filtered['Tanggal_Saja'] == filter_tgl]
                    
                    if df_filtered.empty:
                        st.warning("Tidak ada pengumuman yang cocok dengan tanggal tersebut.")
                    else:
                        pilihan_p_dict = {f"[{row['waktu']}] {row['judul']}": row['id'] for _, row in df_filtered.iterrows()}
                        pilih_label_p = st.selectbox("Pilih Pengumuman Anda yang Ingin Dikelola:", list(pilihan_p_dict.keys()))
                        id_p_terpilih = pilihan_p_dict[pilih_label_p]
                        
                        data_p_terpilih = get_data("SELECT * FROM pengumuman WHERE id = ?", (id_p_terpilih,)).iloc[0]
                        
                        fp_lama_str = data_p_terpilih.get('foto_pengumuman', '')
                        list_fp = [x.strip() for x in fp_lama_str.split(',') if x.strip()]

                        if list_fp:
                            st.markdown("##### 🗑️ Kelola File & Ubah Hak Akses Per File:")
                            
                            for idx_img, item_str in enumerate(list_fp):
                                if "|" in item_str:
                                    path_img, akses_file_item = item_str.split("|", 1)
                                else:
                                    path_img, akses_file_item = item_str, "Public"
                                
                                st.markdown(f"---")
                                ext_file = path_img.split('.')[-1].lower().split('?')[0]
                                
                                col_prev1, col_prev2 = st.columns([2, 2])
                                with col_prev1:
                                    st.markdown(f"**File {idx_img + 1}:** `{path_img}`")
                                    
                                    if akses_file_item == "Private":
                                        st.markdown("🔒 Status: **Tersimpan sebagai Private**")
                                    else:
                                        st.markdown("🌍 Status: **Tersimpan sebagai Public**")

                                    if ext_file in ['jpg', 'jpeg', 'png', 'webp']:
                                        st.markdown(f'<img src="{path_img}" style="width: 100%; max-height: 120px; object-fit: cover; border-radius: 4px; border: 1px solid #ccc;">', unsafe_allow_html=True)
                                    elif ext_file in ['mp4', 'mov', 'avi']:
                                        st.video(path_img)
                                    elif ext_file in ['mp3', 'wav', 'ogg', 'm4a']:
                                        st.audio(path_img)
                                    else:
                                        st.info(f"📄 Berkas / Dokumen Cloud")
                                        
                                with col_prev2:
                                    ubah_akses_item = st.selectbox(
                                        f"Ubah Akses File {idx_img + 1}", 
                                        ["Public", "Private"], 
                                        index=0 if akses_file_item == "Public" else 1, 
                                        key=f"ubah_akses_{id_p_terpilih}_{idx_img}"
                                    )
                                    
                                    if st.button(f"🗑️ Hapus File Ini", key=f"btn_del_file_single_{id_p_terpilih}_{idx_img}"):
                                        list_fp.pop(idx_img)
                                        new_fp_str = ",".join(list_fp)
                                        execute_query("UPDATE pengumuman SET foto_pengumuman = ? WHERE id = ?", (new_fp_str, id_p_terpilih))
                                        st.success(f"File {idx_img + 1} berhasil dihapus!")
                                        st.rerun()
                                        
                                    if st.button(f"💾 Simpan Akses File Ini", key=f"btn_save_akses_{id_p_terpilih}_{idx_img}"):
                                        list_fp[idx_img] = f"{path_img}|{ubah_akses_item}"
                                        new_fp_str = ",".join(list_fp)
                                        execute_query("UPDATE pengumuman SET foto_pengumuman = ? WHERE id = ?", (new_fp_str, id_p_terpilih))
                                        st.success(f"Hak akses File {idx_img + 1} berhasil diperbarui!")
                                        st.rerun()
                        else:
                            st.info("Tidak ada file yang terlampir pada pengumuman ini.")

                        st.markdown("---")
                        edit_jdl = st.text_input("Edit Judul Pengumuman", value=data_p_terpilih['judul'], key=f"ed_jdl_{id_p_terpilih}")
                        edit_isi = st.text_area("Edit Isi Maklumat / Pesan", value=data_p_terpilih['isi'], key=f"ed_isi_{id_p_terpilih}")
                        
                        st.markdown("##### 📁 Tambah File Baru")
                        edit_metode = st.radio("Pilih Cara Tambah File Baru:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key=f"ed_metode_{id_p_terpilih}")
                        edit_foto_files = None
                        edit_cam_file = None
                        if edit_metode == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                            edit_foto_files = st.file_uploader("Upload File Baru (Bisa Banyak)", type=None, accept_multiple_files=True, key=f"ed_up_{id_p_terpilih}")
                        else:
                            edit_cam_file = st.camera_input("Ambil Foto Baru via Kamera", key=f"ed_cam_{id_p_terpilih}")
                        
                        edit_file_access_settings = {}
                        if edit_foto_files:
                            for idx_ef, ef_item in enumerate(edit_foto_files):
                                edit_file_access_settings[ef_item.name] = st.selectbox(
                                    f"Hak Akses File Baru: **{ef_item.name}**", 
                                    ["Public", "Private"], 
                                    index=1, 
                                    key=f"akses_edit_new_{id_p_terpilih}_{idx_ef}"
                                )
                        elif edit_cam_file is not None:
                            edit_cam_akses = st.selectbox("Hak Akses Foto Kamera Baru:", ["Public", "Private"], index=1, key=f"akses_edit_cam_{id_p_terpilih}")
                        
                        col_eb1, col_eb2 = st.columns(2)
                        with col_eb1:
                            if st.button("💾 Simpan Perubahan Teks & Tambah File", key=f"btn_up_{id_p_terpilih}"):
                                existing_paths = list_fp
                                
                                if edit_foto_files:
                                    with st.spinner("Mengunggah file baru ke Cloudinary..."):
                                        for ef_item in edit_foto_files:
                                            try:
                                                upload_ef = cloudinary.uploader.upload(ef_item, resource_type="auto")
                                                ef_url = upload_ef.get("secure_url")
                                                
                                                akses_ef = edit_file_access_settings.get(ef_item.name, "Private")
                                                existing_paths.append(f"{ef_url}|{akses_ef}")
                                            except Exception as e:
                                                st.error(f"Gagal upload {ef_item.name}: {e}")
                                        
                                if edit_cam_file is not None:
                                    with st.spinner("Mengunggah foto kamera baru..."):
                                        try:
                                            upload_ecam = cloudinary.uploader.upload(edit_cam_file, resource_type="image")
                                            ecam_url = upload_ecam.get("secure_url")
                                            existing_paths.append(f"{ecam_url}|{edit_cam_akses}")
                                        except Exception as e:
                                            st.error(f"Gagal upload kamera: {e}")
                                        
                                foto_p_path_e = ",".join(existing_paths)
                                execute_query(
                                    "UPDATE pengumuman SET judul = ?, isi = ?, foto_pengumuman = ? WHERE id = ?",
                                    (edit_jdl, edit_isi, foto_p_path_e, id_p_terpilih)
                                )
                                st.success("Pengumuman dan file baru berhasil disimpan di Cloudinary!")
                                st.rerun()
                        with col_eb2:
                            if st.button("🗑️ Hapus Pengumuman Ini Sepenuhnya", key=f"btn_del_p_{id_p_terpilih}"):
                                execute_query("DELETE FROM pengumuman WHERE id = ?", (id_p_terpilih,))
                                st.success("Pengumuman berhasil dihapus dari sistem!")
                                st.rerun()

        df_pengumuman = get_data("SELECT * FROM pengumuman ORDER BY id DESC")
        if df_pengumuman.empty:
            st.info("Belum ada pengumuman atau maklumat resmi yang disiarkan.")
        else:
            for _, p in df_pengumuman.iterrows():
                with st.container():
                    st.markdown(f"#### 📌 {p['judul']}")
                    
                    pembuat_pengumuman = p.get('pembuat', '')
                    fp_str = p.get('foto_pengumuman', '')
                    arr_foto = [x.strip() for x in fp_str.split(',') if x.strip()]
                    
                    if arr_foto:
                        cols_img = st.columns(min(len(arr_foto), 3))
                        for i, f_item_str in enumerate(arr_foto):
                            if "|" in f_item_str:
                                f_path, akses_file_item = f_item_str.split("|", 1)
                            else:
                                f_path, akses_file_item = f_item_str, "Public"
                                
                            ext_f = f_path.split('.')[-1].lower().split('?')[0]
                            with cols_img[i % 3]:
                                if akses_file_item == "Private":
                                    st.caption(f"🔒 File {i+1}: **Tersimpan sebagai Private**")
                                else:
                                    st.caption(f"🌍 File {i+1}: **Tersimpan sebagai Public**")
                                    
                                if ext_f in ['jpg', 'jpeg', 'png', 'webp']:
                                    st.markdown(f'<img src="{f_path}" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 6px; margin-bottom: 6px; border: 1px solid #ddd;">', unsafe_allow_html=True)
                                elif ext_f in ['mp4', 'mov', 'avi']:
                                    st.video(f_path)
                                elif ext_f in ['mp3', 'wav', 'ogg', 'm4a']:
                                    st.audio(f_path)
                                else:
                                    st.info(f"📄 Berkas / Dokumen Cloud")
                                
                                can_download = (akses_file_item == "Public") or (role == pembuat_pengumuman)
                                if can_download:
                                    st.markdown(f"[📥 Download / Buka File {i+1}]({f_path})", unsafe_allow_html=True)
                                else:
                                    st.warning("🔒 File Private (Akses Dibatasi)")
                    
                    st.write(p['isi'])
                    st.caption(f"Dipublikasikan oleh: {p['pembuat']} pada {p['waktu']}")
                    st.divider()

    with tab_beranda_2:
        st.subheader("Susunan Pimpinan & Pengurus FK Mawil Riau")
        st.caption("Data pejabat pengurus dan Ketua Sub Mawil di-link langsung dari data SanFK yang sudah di-input oleh Sekretaris Mawil.")

        df_struktur = get_data("SELECT * FROM struktur_pengurus")
        if not df_struktur.empty:
            for _, row_s in df_struktur.iterrows():
                f_path_s = row_s.get('foto', '')
                b64_s = get_image_base64(f_path_s)
                
                if b64_s:
                    img_tag = f'<img src="{b64_s}" style="width: 50px; height: 60px; object-fit: cover; border-radius: 4px; border: 1px solid #0E6655;">'
                else:
                    img_tag = '<div style="width: 50px; height: 60px; background: #eee; font-size: 9px; color: #777; display: flex; align-items: center; justify-content: center; border-radius: 4px; text-align: center;">No Foto</div>'
                
                kontak_s = row_s.get('kontak', '-')
                if kontak_s and kontak_s != "-" and kontak_s.strip() != "":
                    cleaned_num = kontak_s.replace("-", "").replace(" ", "")
                    kontak_display = f'<a href="https://wa.me/{cleaned_num}" target="_blank" style="color: #0E6655; text-decoration: none;">📞 {kontak_s}</a>'
                else:
                    kontak_display = "📞 -"

                st.markdown(f"""
                <div style="display: flex; align-items: center; background: #E8F8F5; padding: 10px; border-radius: 8px; border: 1px solid #A2D9CE; margin-bottom: 8px;">
                    <div style="margin-right: 15px;">{img_tag}</div>
                    <div>
                        <h5 style="margin: 0; color: #0E6655;">{row_s['jabatan']}</h5>
                        <p style="margin: 2px 0; font-size: 15px; color: #333;"><b>{row_s['nama_pejabat']}</b></p>
                        <p style="margin: 0; font-size: 13px;">{kontak_display}</p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

        if role == "Sekretaris Mawil":
            with st.expander("🛠️ [Otoritas Sekretaris Mawil] Tentukan Pejabat dari Data SanFK"):
                st.info("Pilih jabatan, lalu pilih nama pejabat dari daftar SanFK yang terdaftar di database.")
                df_list_jabatan = get_data("SELECT id, jabatan FROM struktur_pengurus")
                
                pilihan_jab_dict = {row['jabatan']: row['id'] for _, row in df_list_jabatan.iterrows()}
                pilih_jab_label = st.selectbox("Pilih Jabatan / Posisi Pengurus:", list(pilihan_jab_dict.keys()))
                id_jab_terpilih = pilihan_jab_dict[pilih_jab_label]
                
                data_pejabat_lama = get_data("SELECT * FROM struktur_pengurus WHERE id = ?", (id_jab_terpilih,)).iloc[0]

                df_all_sanfk = get_data("SELECT * FROM anggota")
                
                with st.form("form_edit_pengurus_linked"):
                    if df_all_sanfk.empty:
                        st.warning("Belum ada data SanFK yang di-input. Silakan input data SanFK terlebih dahulu melalui menu Manajemen SanFK.")
                        list_opsi_sanfk = ["(Belum Ditetapkan)"]
                    else:
                        list_opsi_sanfk = ["(Belum Ditetapkan)"] + df_all_sanfk['nama'].tolist()
                    
                    current_name = data_pejabat_lama['nama_pejabat']
                    idx_default = list_opsi_sanfk.index(current_name) if current_name in list_opsi_sanfk else 0
                    
                    pilih_sanfk_pejabat = st.selectbox("Pilih Nama dari Data SanFK Terdaftar:", list_opsi_sanfk, index=idx_default)
                    
                    col_bp1, col_bp2 = st.columns(2)
                    with col_bp1:
                        btn_simpan_pejabat = st.form_submit_button("💾 Simpan & Sinkronkan Pejabat")
                    with col_bp2:
                        btn_reset_pejabat = st.form_submit_button("🗑️ Kosongkan Posisi Ini")
                    
                    if btn_simpan_pejabat:
                        if pilih_sanfk_pejabat == "(Belum Ditetapkan)":
                            execute_query(
                                "UPDATE struktur_pengurus SET nama_pejabat = ?, kontak = ?, foto = ? WHERE id = ?",
                                ("(Belum Ditetapkan)", "-", "", id_jab_terpilih)
                            )
                            st.success(f"Posisi **{pilih_jab_label}** dikosongkan.")
                        else:
                            row_sanfk_terpilih = df_all_sanfk[df_all_sanfk['nama'] == pilih_sanfk_pejabat].iloc[0]
                            nama_p = row_sanfk_terpilih['nama']
                            kontak_p = row_sanfk_terpilih['kontak'] if row_sanfk_terpilih['kontak'] and row_sanfk_terpilih['kontak'].strip() != "" else "-"
                            
                            foto_full_s = row_sanfk_terpilih['foto']
                            foto_p = foto_full_s.split("|")[0] if "|" in foto_full_s else foto_full_s
                            
                            execute_query(
                                "UPDATE struktur_pengurus SET nama_pejabat = ?, kontak = ?, foto = ? WHERE id = ?",
                                (nama_p, kontak_p, foto_p, id_jab_terpilih)
                            )
                            st.success(f"Pejabat untuk **{pilih_jab_label}** berhasil di-link ke SanFK **{nama_p}** (Kontak & foto tersinkron otomatis)!")
                        st.rerun()

                    if btn_reset_pejabat:
                        execute_query(
                            "UPDATE struktur_pengurus SET nama_pejabat = ?, kontak = ?, foto = ? WHERE id = ?",
                            ("(Belum Ditetapkan)", "-", "", id_jab_terpilih)
                        )
                        st.success(f"Posisi **{pilih_jab_label}** berhasil dikosongkan!")
                        st.rerun()

# --- 4. AGENDA KOPDAR MAWIL & BAKSOS ---
elif menu == "Agenda Kopdar Mawil & Baksos":
    st.title("📅 Agenda Kopdar Mawil & Baksos (Tingkat Provinsi)")
    st.info("Menu khusus untuk mengelola jadwal, konfirmasi RSVP, transparansi dana baksos, dan penguncian kehadiran aktual.")

    if role == "Ketua Mawil":
        st.success("Panel Khusus **Ketua Mawil** (Pengelola Agenda Kopdar & Baksos)")
        with st.form("form_buat_kopdar"):
            st.markdown("#### 📅 Buat Agenda Kopdar / Baksos Baru")
            kat_kopdar = st.selectbox("Kategori Agenda", ["Kopdar Mawil", "Baksos Tahunan", "Baksos Sosial / Bencana"])
            tgl_kopdar = st.date_input("Tanggal Pelaksanaan Kopdar/Baksos")
            sub_mawil_tuan_rumah = st.selectbox("Sub Mawil (Kabupaten/Kota Tuan Rumah)", DAFTAR_KAB_KOTA)
            lokasi_kopdar = st.text_input("Lokasi / Tempat Pelaksanaan")
            nominal_baksos = st.number_input("Nominal / Nilai Dana Disalurkan (Rp) [Khusus Baksos / Opsional]", min_value=0.0, step=100000.0)
            ket_kopdar = st.text_area("Keterangan / Susunan Acara")
            submit_buat_kopdar = st.form_submit_button("Publikasikan Agenda Kopdar/Baksos")
            
            if submit_buat_kopdar and lokasi_kopdar:
                execute_query(
                    "INSERT INTO jadwal_kopdar_baksos (kategori, sub_mawil_tuan_rumah, tanggal, lokasi, nominal_dana, keterangan) VALUES (?, ?, ?, ?, ?, ?)",
                    (kat_kopdar, sub_mawil_tuan_rumah, str(tgl_kopdar), lokasi_kopdar, nominal_baksos, ket_kopdar)
                )
                st.success("Agenda Kopdar / Baksos dan transparansi dana berhasil dipublikasikan!")
                st.rerun()

        st.markdown("---")
        st.markdown("#### 🛠️ Kelola / Edit / Hapus Agenda Kopdar & Baksos")
        df_kelola_kb = get_data("SELECT id, kategori, sub_mawil_tuan_rumah, tanggal, lokasi, nominal_dana FROM jadwal_kopdar_baksos")
        if df_kelola_kb.empty:
            st.info("Belum ada agenda Kopdar/Baksos yang dibuat.")
        else:
            for _, r_kb in df_kelola_kb.iterrows():
                with st.expander(f"[{r_kb['kategori']}] {r_kb['sub_mawil_tuan_rumah']} | Tanggal: {r_kb['tanggal']} - {r_kb['lokasi']} (Dana: Rp {r_kb['nominal_dana']:,.0f})"):
                    with st.form(f"form_edit_kb_{r_kb['id']}"):
                        e_kat = st.selectbox("Ubah Kategori", ["Kopdar Mawil", "Baksos Tahunan", "Baksos Sosial / Bencana"], index=["Kopdar Mawil", "Baksos Tahunan", "Baksos Sosial / Bencana"].index(r_kb['kategori']) if r_kb['kategori'] in ["Kopdar Mawil", "Baksos Tahunan", "Baksos Sosial / Bencana"] else 0, key=f"kb_kat_{r_kb['id']}")
                        e_sub = st.selectbox("Ubah Sub Mawil (Kabupaten/Kota)", DAFTAR_KAB_KOTA, index=DAFTAR_KAB_KOTA.index(r_kb['sub_mawil_tuan_rumah']) if r_kb['sub_mawil_tuan_rumah'] in DAFTAR_KAB_KOTA else 0, key=f"kb_sub_{r_kb['id']}")
                        e_tgl = st.date_input("Ubah Tanggal", value=datetime.strptime(r_kb['tanggal'], "%Y-%m-%d").date(), key=f"kb_tgl_{r_kb['id']}")
                        e_lok = st.text_input("Ubah Lokasi", value=r_kb['lokasi'], key=f"kb_lok_{r_kb['id']}")
                        e_nom = st.number_input("Ubah Nominal Dana (Rp)", value=float(r_kb['nominal_dana'] if r_kb['nominal_dana'] else 0.0), min_value=0.0, step=100000.0, key=f"kb_nom_{r_kb['id']}")
                        
                        c_b1, c_b2 = st.columns(2)
                        with c_b1:
                            btn_up_kb = st.form_submit_button("💾 Perbarui Agenda")
                        with c_b2:
                            btn_del_kb = st.form_submit_button("🗑️ Hapus Agenda")
                            
                        if btn_up_kb:
                            execute_query("UPDATE jadwal_kopdar_baksos SET kategori = ?, sub_mawil_tuan_rumah = ?, tanggal = ?, lokasi = ?, nominal_dana = ? WHERE id = ?", (e_kat, e_sub, str(e_tgl), e_lok, e_nom, r_kb['id']))
                            st.success("Agenda berhasil diperbarui!")
                            st.rerun()
                        if btn_del_kb:
                            execute_query("DELETE FROM jadwal_kopdar_baksos WHERE id = ?", (r_kb['id'],))
                            execute_query("DELETE FROM rsvp_kopdar_baksos WHERE jadwal_id = ?", (r_kb['id'],))
                            execute_query("DELETE FROM aktual_hadir_kopdar_baksos WHERE jadwal_id = ?", (r_kb['id'],))
                            st.success("Agenda berhasil dihapus!")
                            st.rerun()

    st.markdown("---")
    df_list_kb = get_data("SELECT * FROM jadwal_kopdar_baksos ORDER BY tanggal DESC")
    if df_list_kb.empty:
        st.warning("Belum ada agenda Kopdar Mawil atau Baksos yang dipublikasikan.")
    else:
        pilihan_kb_dict = {f"[{row['kategori']}] ({row['sub_mawil_tuan_rumah']}) {row['tanggal']} - {row['lokasi']}": row['id'] for _, row in df_list_kb.iterrows()}
        pilih_kb_label = st.selectbox("Pilih Agenda Kopdar / Baksos Resmi:", list(pilihan_kb_dict.keys()))
        id_kb_aktif = pilihan_kb_dict[pilih_kb_label]
        
        data_kb_pilih = get_data("SELECT * FROM jadwal_kopdar_baksos WHERE id = ?", (id_kb_aktif,)).iloc[0]
        
        nom_tampil = data_kb_pilih['nominal_dana'] if data_kb_pilih['nominal_dana'] else 0.0
        st.markdown(f"""
        <div style="background-color: #E8F8F5; padding: 15px; border-radius: 8px; border: 1px solid #0E6655; margin-bottom: 15px;">
            <h4 style="margin: 0; color: #0E6655;">📋 Detail Agenda & Transparansi Dana</h4>
            <p style="margin: 6px 0 2px 0;"><b>Kategori:</b> {data_kb_pilih['kategori']}</p>
            <p style="margin: 2px 0;"><b>Tuan Rumah (Sub Mawil):</b> {data_kb_pilih['sub_mawil_tuan_rumah']}</p>
            <p style="margin: 2px 0;"><b>Lokasi:</b> {data_kb_pilih['lokasi']}</p>
            <p style="margin: 2px 0;"><b>Tanggal:</b> {data_kb_pilih['tanggal']}</p>
            <p style="margin: 2px 0; color: #117A65; font-size: 16px;"><b>💰 Nilai Dana Disalurkan: Rp {nom_tampil:,.0f}</b></p>
            <hr style="border-color: #A2D9CE; margin: 8px 0;">
            <p style="margin: 0; font-size: 14px;"><b>Keterangan:</b> {data_kb_pilih['keterangan']}</p>
        </div>
        """, unsafe_allow_html=True)
        
        if role != "Ketua Mawil":
            st.markdown("---")
            st.subheader("Konfirmasi Kehadiran (RSVP) SanFK untuk Agenda Ini:")
            
            df_all_agt = get_data("SELECT nama, sub_mawil FROM anggota")
            if df_all_agt.empty:
                st.warning("Belum ada data SanFK terdaftar di sistem.")
            else:
                list_nama_agt_full = [f"{r['nama']} ({r['sub_mawil']})" for _, r in df_all_agt.iterrows()]
                with st.form("form_rsvp_kb"):
                    pilih_agt_kb = st.selectbox("Pilih Nama SanFK:", options=list_nama_agt_full)
                    status_rsvp_kb = st.radio("Status Kehadiran Kopdar/Baksos:", ["Inshaa Allah akan Hadir", "Tidak Bisa Hadir", "Belum Dapat Dipastikan"])
                    
                    col_rk1, col_rk2 = st.columns(2)
                    with col_rk1:
                        btn_s_rsvp = st.form_submit_button("Kirim RSVP Kopdar/Baksos")
                    with col_rk2:
                        btn_d_rsvp = st.form_submit_button("🗑️ Batalkan RSVP")
                        
                    if btn_s_rsvp and pilih_agt_kb:
                        nama_bersih = pilih_agt_kb.split(" (")[0]
                        sub_bersih = pilih_agt_kb.split(" (")[1].replace(")", "")
                        
                        cek_r_kb = get_data("SELECT * FROM rsvp_kopdar_baksos WHERE nama = ? AND jadwal_id = ?", (nama_bersih, id_kb_aktif))
                        if not cek_r_kb.empty:
                            execute_query("UPDATE rsvp_kopdar_baksos SET status_rsvp = ? WHERE nama = ? AND jadwal_id = ?", (status_rsvp_kb, nama_bersih, id_kb_aktif))
                            st.success(f"RSVP {nama_bersih} diperbarui menjadi '{status_rsvp_kb}'!")
                        else:
                            execute_query("INSERT INTO rsvp_kopdar_baksos (nama, sub_mawil, jadwal_id, status_rsvp) VALUES (?, ?, ?, ?)", (nama_bersih, sub_bersih, id_kb_aktif, status_rsvp_kb))
                            st.success(f"Terima kasih {nama_bersih}, RSVP Anda berhasil dicatat!")
                        st.rerun()
                        
                    if btn_d_rsvp and pilih_agt_kb:
                        nama_bersih = pilih_agt_kb.split(" (")[0]
                        execute_query("DELETE FROM rsvp_kopdar_baksos WHERE nama = ? AND jadwal_id = ?", (nama_bersih, id_kb_aktif))
                        st.success("RSVP berhasil dihapus!")
                        st.rerun()
        else:
            st.markdown("---")
            st.subheader("Konfirmasi Kehadiran (RSVP) SanFK untuk Agenda Ini:")
            st.info("💡 *(khusus ketua mawil login dengan role SanFK untuk konfirmasi)*")

        st.markdown("---")
        st.subheader("Daftar RSVP SanFK Sementara:")
        df_tampil_rsvp_kb = get_data("SELECT nama as 'Nama', sub_mawil as 'Sub Mawil', status_rsvp as 'Status RSVP' FROM rsvp_kopdar_baksos WHERE jadwal_id = ?", (id_kb_aktif,))
        if not df_tampil_rsvp_kb.empty:
            df_tampil_rsvp_kb.index = range(1, len(df_tampil_rsvp_kb) + 1)
            st.dataframe(df_tampil_rsvp_kb, use_container_width=True)
        else:
            st.info("Belum ada konfirmasi RSVP untuk agenda ini.")

        if role == "Ketua Mawil":
            st.markdown("---")
            st.markdown("#### 🔒 Input & Kunci Aktual Kehadiran Kopdar / Baksos (Khusus Ketua Mawil)")
            
            df_all_agt = get_data("SELECT nama, sub_mawil FROM anggota")
            list_nama_mentah = df_all_agt["nama"].tolist() if not df_all_agt.empty else []
            with st.form("form_aktual_kb_mawil"):
                hadir_kb_final = st.multiselect("Pilih SanFK yang Hadir Aktual di Acara Kopdar/Baksos:", options=list_nama_mentah)
                btn_simpan_akt_kb = st.form_submit_button("Simpan & Kunci Aktual Kehadiran Mawil")
                
                if btn_simpan_akt_kb and hadir_kb_final:
                    execute_query("DELETE FROM aktual_hadir_kopdar_baksos WHERE jadwal_id = ?", (id_kb_aktif,))
                    waktu_skr_kb = datetime.now().strftime("%Y-%m-%d %H:%M")
                    for n_hadir in hadir_kb_final:
                        row_agt_sub = df_all_agt[df_all_agt["nama"] == n_hadir].iloc[0]
                        execute_query(
                            "INSERT INTO aktual_hadir_kopdar_baksos (nama, sub_mawil, jadwal_id, waktu_input) VALUES (?, ?, ?, ?)",
                            (n_hadir, row_agt_sub['sub_mawil'], id_kb_aktif, waktu_skr_kb)
                        )
                    st.success("Aktual kehadiran Kopdar / Baksos berhasil dikunci oleh Ketua Mawil!")
                    st.rerun()

            st.markdown("---")
            st.subheader("📋 Actual List Kehadiran Kopdar / Baksos & Koreksi:")
            df_akt_kb_tampil = get_data("SELECT id, nama as 'Nama', sub_mawil as 'Sub Mawil', waktu_input as 'Waktu Input' FROM aktual_hadir_kopdar_baksos WHERE jadwal_id = ?", (id_kb_aktif,))
            if not df_akt_kb_tampil.empty:
                df_akt_kb_clean = df_akt_kb_tampil.drop(columns=['id'])
                df_akt_kb_clean.index = range(1, len(df_akt_kb_clean) + 1)
                st.dataframe(df_akt_kb_clean, use_container_width=True)
                
                with st.form("form_hapus_akt_kb"):
                    pilih_del_akt_kb = st.selectbox("Pilih Nama SanFK yang Ingin Dihapus dari Kehadiran Aktual:", options=df_akt_kb_tampil['Nama'].tolist())
                    btn_eksekusi_del_akt = st.form_submit_button("Hapus Nama dari Daftar Hadir Aktual")
                    if btn_eksekusi_del_akt and pilih_del_akt_kb:
                        execute_query("DELETE FROM aktual_hadir_kopdar_baksos WHERE jadwal_id = ? AND nama = ?", (id_kb_aktif, pilih_del_akt_kb))
                        st.success(f"Nama **{pilih_del_akt_kb}** berhasil dihapus dari kehadiran aktual!")
                        st.rerun()
            else:
                st.info("Belum ada data aktual kehadiran Kopdar/Baksos yang dikunci.")

# --- 8. GALERI RESMI ADMIN ---
elif menu == "Galeri Resmi (Admin)":
    st.title("🖼️ Galeri Resmi Organisasi (Dokumentasi Mawil)")
    st.write("Dokumentasi resmi kegiatan, kopdar, dan baksos Mawil Riau.")
    
    if role not in ["SanFK", "Admin Dokumentasi Mawil"]:
        st.warning("⚠️ Menu **Galeri Resmi (Admin)** ini hanya dapat diakses oleh role **SanFK** dan **Admin Dokumentasi Mawil** sesuai pengaturan hak akses yang telah ditetapkan.")
    else:
        if role == "Admin Dokumentasi Mawil":
            tab_gresmi_1, tab_gresmi_2 = st.tabs(["📋 Tampilan Galeri Resmi", "➕ Buat & Kelola Dokumentasi"])
            
            with tab_gresmi_1:
                st.subheader("Daftar Dokumentasi Resmi")
                df_gresmi = get_data("SELECT * FROM galeri_resmi ORDER BY id DESC")
                
                if df_gresmi.empty:
                    st.info("Belum ada dokumentasi resmi yang diunggah.")
                else:
                    for _, post in df_gresmi.iterrows():
                        with st.container():
                            post_id_r = post['id']
                            st.markdown(f"### 📌 {post['judul']}")
                            st.caption(f"Waktu: {post['waktu']}")
                            
                            fr_str = post.get('foto_resmi', '')
                            arr_foto_r = [x.strip() for x in fr_str.split(',') if x.strip()]
                            
                            if arr_foto_r:
                                cols_img_r = st.columns(min(len(arr_foto_r), 3))
                                for i, f_item_str_r in enumerate(arr_foto_r):
                                    if "|" in f_item_str_r:
                                        f_path_r, akses_file_item_r = f_item_str_r.split("|", 1)
                                    else:
                                        f_path_r, akses_file_item_r = f_item_str_r, "Public"
                                        
                                    ext_fr = f_path_r.split('.')[-1].lower().split('?')[0]
                                    with cols_img_r[i % 3]:
                                        if akses_file_item_r == "Private":
                                            st.caption(f"🔒 File {i+1}: **Tersimpan sebagai Private**")
                                        else:
                                            st.caption(f"🌍 File {i+1}: **Tersimpan sebagai Public**")
                                            
                                        if ext_fr in ['jpg', 'jpeg', 'png', 'webp']:
                                            st.markdown(f'<img src="{f_path_r}" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 6px; margin-bottom: 6px; border: 1px solid #ddd;">', unsafe_allow_html=True)
                                        elif ext_fr in ['mp4', 'mov', 'avi']:
                                            st.video(f_path_r)
                                        elif ext_fr in ['mp3', 'wav', 'ogg', 'm4a']:
                                            st.audio(f_path_r)
                                        else:
                                            st.info(f"📄 Berkas / Dokumen Cloud")
                                        
                                        can_download_r = (akses_file_item_r == "Public") or (role == "Admin Dokumentasi Mawil")
                                        if can_download_r:
                                            st.markdown(f"[📥 Download / Buka File {i+1}]({f_path_r})", unsafe_allow_html=True)
                                        else:
                                            st.warning("🔒 File Private (Akses Dibatasi)")

                            df_reaksi_resmi = get_data("SELECT reaksi, COUNT(*) as jml FROM reaksi_resmi WHERE post_id = ? GROUP BY reaksi", (post_id_r,))
                            c_like_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Like']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Like'].empty else 0
                            c_dislike_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Dislike']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Dislike'].empty else 0
                            c_love_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Love']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Love'].empty else 0

                            st.markdown(f"👍 **{c_like_r}** Likes | 👎 **{c_dislike_r}** Dislikes | ❤️ **{c_love_r}** Loves")

                            if role == "SanFK" and sanfk_aktif_terpilih:
                                cols_reaksi_r = st.columns(4)
                                with cols_reaksi_r[0]:
                                    if st.button("👍 Like", key=f"btn_rlike_{post_id_r}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Like"))
                                        st.rerun()
                                with cols_reaksi_r[1]:
                                    if st.button("👎 Dislike", key=f"btn_rdislike_{post_id_r}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Dislike"))
                                        st.rerun()
                                with cols_reaksi_r[2]:
                                    if st.button("❤️ Love", key=f"btn_rlove_{post_id_r}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Love"))
                                        st.rerun()
                                with cols_reaksi_r[3]:
                                    if st.button("❌ Batal Reaksi", key=f"btn_runreact_{post_id_r}"):
                                        execute_query("DELETE FROM reaksi_resmi WHERE post_id = ? AND nama_sanfk = ?", (post_id_r, sanfk_aktif_terpilih))
                                        st.rerun()

                            st.divider()

            with tab_gresmi_2:
                st.subheader("Unggah Dokumentasi Resmi Baru")
                judul_kegiatan = st.text_input("Tulis pesan, mutiara hikmah, atau informasi kegiatan...", key="g_judul")
                
                st.markdown("---")
                st.markdown("##### 📁 Lampiran Berbagai File (Bisa Banyak: Foto, PDF, Video, dll)")
                metode_galeri = st.radio("Pilih Cara Input File Dokumentasi:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key="radio_galeri")
                
                foto_resmi_files = None
                cam_resmi = None
                if metode_galeri == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                    foto_resmi_files = st.file_uploader("Pilih File (Bisa Banyak: PDF, MP3, MP4, JPG, PNG, dll)", type=None, accept_multiple_files=True, key="up_galeri")
                else:
                    cam_resmi = st.camera_input("Ambil Foto Dokumentasi via Kamera", key="cam_galeri")
                
                file_access_settings_r = {}
                if foto_resmi_files:
                    st.markdown("##### ⚙️ Atur Hak Akses Masing-Masing File:")
                    for idx_fr, fr_item in enumerate(foto_resmi_files):
                        file_access_settings_r[fr_item.name] = st.selectbox(
                            f"Hak Akses untuk file: **{fr_item.name}**", 
                            ["Public", "Private"], 
                            index=1, 
                            key=f"akses_file_r_{idx_fr}"
                        )
                elif cam_resmi is not None:
                    cam_akses_r = st.selectbox("Hak Akses untuk Foto Kamera:", ["Public", "Private"], index=1, key="akses_cam_r")
                
                st.caption("ℹ️ **Public**: Bisa didownload oleh semua role. **Private**: Hanya bisa didownload oleh admin.")
                
                if st.button("Publikasikan ke Galeri Resmi", key="btn_pub_galeri"):
                    if judul_kegiatan and (foto_resmi_files or cam_resmi):
                        path_list_r = []
                        if foto_resmi_files:
                            with st.spinner("Mengunggah file ke Cloudinary..."):
                                for fr_item in foto_resmi_files:
                                    try:
                                        upload_fr = cloudinary.uploader.upload(fr_item, resource_type="auto")
                                        fr_url = upload_fr.get("secure_url")
                                        
                                        akses_dipilih_r = file_access_settings_r.get(fr_item.name, "Private")
                                        path_list_r.append(f"{fr_url}|{akses_dipilih_r}")
                                    except Exception as e:
                                        st.error(f"Gagal mengunggah {fr_item.name}: {e}")
                                
                        if cam_resmi is not None:
                            with st.spinner("Mengunggah foto kamera ke Cloudinary..."):
                                try:
                                    upload_cam_r = cloudinary.uploader.upload(cam_resmi, resource_type="image")
                                    cam_url_r = upload_cam_r.get("secure_url")
                                    path_list_r.append(f"{cam_url_r}|{cam_akses_r}")
                                except Exception as e:
                                    st.error(f"Gagal mengunggah foto kamera: {e}")
                                
                        foto_r_str = ",".join(path_list_r) if path_list_r else ""
                        waktu_post_r = datetime.now().strftime("%Y-%m-%d %H:%M")
                        execute_query(
                            "INSERT INTO galeri_resmi (judul, kategori, waktu, foto_resmi) VALUES (?, ?, ?, ?)",
                            (judul_kegiatan, "-", waktu_post_r, foto_r_str)
                        )
                        st.success(f"Dokumentasi berhasil diunggah ke Galeri Resmi via Cloudinary!")
                        st.rerun()
                    else:
                        st.warning("Kolom input dan file lampiran wajib diisi!")

                st.markdown("---")
                st.subheader("🛠️ Kelola, Edit, atau Hapus Dokumentasi Resmi Anda")
                
                df_my_resmi = get_data("SELECT * FROM galeri_resmi ORDER BY id DESC")
                if df_my_resmi.empty:
                    st.info("Belum ada dokumentasi resmi yang dapat dikelola.")
                else:
                    pilihan_mr = {f"[{row['waktu']}] {row['judul']}": row['id'] for _, row in df_my_resmi.iterrows()}
                    pilih_label_mr = st.selectbox("Pilih Dokumentasi untuk Dikelola / Dihapus:", list(pilihan_mr.keys()), key="select_kelola_resmi")
                    id_mr_aktif = pilihan_mr[pilih_label_mr]
                    
                    data_resmi_terpilih = get_data("SELECT * FROM galeri_resmi WHERE id = ?", (id_mr_aktif,)).iloc[0]
                    
                    fr_lama_str = data_resmi_terpilih.get('foto_resmi', '')
                    list_fr = [x.strip() for x in fr_lama_str.split(',') if x.strip()]

                    if list_fr:
                        st.markdown("##### 🗑️ Kelola File & Ubah Hak Akses Per File:")
                        for idx_img_r, item_str_r in enumerate(list_fr):
                            if "|" in item_str_r:
                                path_img_r, akses_file_item_r = item_str_r.split("|", 1)
                            else:
                                path_img_r, akses_file_item_r = item_str_r, "Public"
                            
                            st.markdown(f"---")
                            ext_file_r = path_img_r.split('.')[-1].lower().split('?')[0]
                            
                            col_pr1, col_pr2 = st.columns([2, 2])
                            with col_pr1:
                                st.markdown(f"**File {idx_img_r + 1}:** `{path_img_r}`")
                                if akses_file_item_r == "Private":
                                    st.markdown("🔒 Status: **Tersimpan sebagai Private**")
                                else:
                                    st.markdown("🌍 Status: **Tersimpan sebagai Public**")

                                if ext_file_r in ['jpg', 'jpeg', 'png', 'webp']:
                                    st.markdown(f'<img src="{path_img_r}" style="width: 100%; max-height: 120px; object-fit: cover; border-radius: 4px; border: 1px solid #ccc;">', unsafe_allow_html=True)
                                elif ext_file_r in ['mp4', 'mov', 'avi']:
                                    st.video(path_img_r)
                                elif ext_file_r in ['mp3', 'wav', 'ogg', 'm4a']:
                                    st.audio(path_img_r)
                                else:
                                    st.info(f"📄 Berkas / Dokumen Cloud")
                                    
                            with col_pr2:
                                ubah_akses_item_r = st.selectbox(
                                    f"Ubah Akses File {idx_img_r + 1}", 
                                    ["Public", "Private"], 
                                    index=0 if akses_file_item_r == "Public" else 1, 
                                    key=f"ubah_akses_r_{id_mr_aktif}_{idx_img_r}"
                                )
                                
                                if st.button(f"🗑️ Hapus File Ini", key=f"btn_del_file_r_single_{id_mr_aktif}_{idx_img_r}"):
                                    list_fr.pop(idx_img_r)
                                    new_fr_str = ",".join(list_fr)
                                    execute_query("UPDATE galeri_resmi SET foto_resmi = ? WHERE id = ?", (new_fr_str, id_mr_aktif))
                                    st.success(f"File {idx_img_r + 1} berhasil dihapus!")
                                    st.rerun()
                                    
                                if st.button(f"💾 Simpan Akses File Ini", key=f"btn_save_akses_r_{id_mr_aktif}_{idx_img_r}"):
                                    list_fr[idx_img_r] = f"{path_img_r}|{ubah_akses_item_r}"
                                    new_fr_str = ",".join(list_fr)
                                    execute_query("UPDATE galeri_resmi SET foto_resmi = ? WHERE id = ?", (new_fr_str, id_mr_aktif))
                                    st.success(f"Hak akses File {idx_img_r + 1} berhasil diperbarui!")
                                    st.rerun()
                    else:
                        st.info("Tidak ada file yang terlampir pada dokumentasi ini.")

                    st.markdown("---")
                    edit_judul_resmi = st.text_input("Edit Informasi / Pesan", value=data_resmi_terpilih['judul'], key=f"ed_judul_resmi_{id_mr_aktif}")
                    
                    st.markdown("##### 📁 Tambah File Baru")
                    edit_metode_r = st.radio("Pilih Cara Tambah File Baru:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key=f"ed_metode_r_{id_mr_aktif}")
                    edit_foto_files_r = None
                    edit_cam_file_r = None
                    if edit_metode_r == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                        edit_foto_files_r = st.file_uploader("Upload File Baru (Bisa Banyak)", type=None, accept_multiple_files=True, key=f"ed_up_r_{id_mr_aktif}")
                    else:
                        edit_cam_file_r = st.camera_input("Ambil Foto Baru via Kamera", key=f"ed_cam_r_{id_mr_aktif}")
                    
                    edit_file_access_settings_r = {}
                    if edit_foto_files_r:
                        for idx_efr, efr_item in enumerate(edit_foto_files_r):
                            edit_file_access_settings_r[efr_item.name] = st.selectbox(
                                f"Hak Akses File Baru: **{efr_item.name}**", 
                                ["Public", "Private"], 
                                index=1, 
                                key=f"akses_edit_new_r_{id_mr_aktif}_{idx_efr}"
                            )
                    elif edit_cam_file_r is not None:
                        edit_cam_akses_r = st.selectbox("Hak Akses Foto Kamera Baru:", ["Public", "Private"], index=1, key=f"akses_edit_cam_r_{id_mr_aktif}")
                    
                    col_er1, col_er2 = st.columns(2)
                    with col_er1:
                        if st.button("💾 Simpan Perubahan & Tambah File", key=f"btn_up_resmi_{id_mr_aktif}"):
                            existing_paths_r = list_fr
                            
                            if edit_foto_files_r:
                                with st.spinner("Mengunggah file baru ke Cloudinary..."):
                                    for efr_item in edit_foto_files_r:
                                        try:
                                            upload_efr = cloudinary.uploader.upload(efr_item, resource_type="auto")
                                            efr_url = upload_efr.get("secure_url")
                                            
                                            akses_efr = edit_file_access_settings_r.get(efr_item.name, "Private")
                                            existing_paths_r.append(f"{efr_url}|{akses_efr}")
                                        except Exception as e:
                                            st.error(f"Gagal upload {efr_item.name}: {e}")
                                    
                            if edit_cam_file_r is not None:
                                with st.spinner("Mengunggah foto kamera baru..."):
                                    try:
                                        upload_ecam_r = cloudinary.uploader.upload(edit_cam_file_r, resource_type="image")
                                        ecam_url_r = upload_ecam_r.get("secure_url")
                                        existing_paths_r.append(f"{ecam_url_r}|{edit_cam_akses_r}")
                                    except Exception as e:
                                        st.error(f"Gagal upload kamera: {e}")
                                    
                            foto_r_path_e = ",".join(existing_paths_r)
                            execute_query(
                                "UPDATE galeri_resmi SET judul = ?, foto_resmi = ? WHERE id = ?",
                                (edit_judul_resmi, foto_r_path_e, id_mr_aktif)
                            )
                            st.success("Dokumentasi resmi berhasil diperbarui di Cloudinary!")
                            st.rerun()
                    with col_er2:
                        if st.button("🗑️ Hapus Dokumentasi Ini Sepenuhnya", key=f"btn_del_resmi_full_{id_mr_aktif}"):
                            execute_query("DELETE FROM galeri_resmi WHERE id = ?", (id_mr_aktif,))
                            st.success("Dokumentasi resmi berhasil dihapus dari sistem!")
                            st.rerun()
        else:
            st.subheader("Daftar Dokumentasi Resmi")
            df_gresmi = get_data("SELECT * FROM galeri_resmi ORDER BY id DESC")
            
            if df_gresmi.empty:
                st.info("Belum ada dokumentasi resmi yang diunggah.")
            else:
                for _, post in df_gresmi.iterrows():
                    with st.container():
                        post_id_r = post['id']
                        st.markdown(f"### 📌 {post['judul']}")
                        st.caption(f"Waktu: {post['waktu']}")
                        
                        fr_str = post.get('foto_resmi', '')
                        arr_foto_r = [x.strip() for x in fr_str.split(',') if x.strip()]
                        
                        if arr_foto_r:
                            cols_img_r = st.columns(min(len(arr_foto_r), 3))
                            for i, f_item_str_r in enumerate(arr_foto_r):
                                if "|" in f_item_str_r:
                                    f_path_r, akses_file_item_r = f_item_str_r.split("|", 1)
                                else:
                                    f_path_r, akses_file_item_r = f_item_str_r, "Public"
                                    
                                ext_fr = f_path_r.split('.')[-1].lower().split('?')[0]
                                with cols_img_r[i % 3]:
                                    if akses_file_item_r == "Private":
                                        st.caption(f"🔒 File {i+1}: **Tersimpan sebagai Private**")
                                    else:
                                        st.caption(f"🌍 File {i+1}: **Tersimpan sebagai Public**")
                                        
                                    if ext_fr in ['jpg', 'jpeg', 'png', 'webp']:
                                        st.markdown(f'<img src="{f_path_r}" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 6px; margin-bottom: 6px; border: 1px solid #ddd;">', unsafe_allow_html=True)
                                    elif ext_fr in ['mp4', 'mov', 'avi']:
                                        st.video(f_path_r)
                                    elif ext_fr in ['mp3', 'wav', 'ogg', 'm4a']:
                                        st.audio(f_path_r)
                                    else:
                                        st.info(f"📄 Berkas / Dokumen Cloud")
                                    
                                    can_download_r = (akses_file_item_r == "Public") or (role == "Admin Dokumentasi Mawil")
                                    if can_download_r:
                                        st.markdown(f"[📥 Download / Buka File {i+1}]({f_path_r})", unsafe_allow_html=True)
                                    else:
                                        st.warning("🔒 File Private (Akses Dibatasi)")

                        df_reaksi_resmi = get_data("SELECT reaksi, COUNT(*) as jml FROM reaksi_resmi WHERE post_id = ? GROUP BY reaksi", (post_id_r,))
                        c_like_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Like']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Like'].empty else 0
                        c_dislike_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Dislike']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Dislike'].empty else 0
                        c_love_r = int(df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Love']['jml'].values[0]) if not df_reaksi_resmi[df_reaksi_resmi['reaksi'] == 'Love'].empty else 0

                        st.markdown(f"👍 **{c_like_r}** Likes | 👎 **{c_dislike_r}** Dislikes | ❤️ **{c_love_r}** Loves")

                        if role == "SanFK" and sanfk_aktif_terpilih:
                            cols_reaksi_r = st.columns(4)
                            with cols_reaksi_r[0]:
                                if st.button("👍 Like", key=f"btn_rlike_{post_id_r}"):
                                    execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Like"))
                                    st.rerun()
                            with cols_reaksi_r[1]:
                                if st.button("👎 Dislike", key=f"btn_rdislike_{post_id_r}"):
                                    execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Dislike"))
                                    st.rerun()
                            with cols_reaksi_r[2]:
                                if st.button("❤️ Love", key=f"btn_rlove_{post_id_r}"):
                                    execute_query("INSERT OR REPLACE INTO reaksi_resmi (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id_r, sanfk_aktif_terpilih, "Love"))
                                    st.rerun()
                            with cols_reaksi_r[3]:
                                if st.button("❌ Batal Reaksi", key=f"btn_runreact_{post_id_r}"):
                                    execute_query("DELETE FROM reaksi_resmi WHERE post_id = ? AND nama_sanfk = ?", (post_id_r, sanfk_aktif_terpilih))
                                    st.rerun()

                        st.divider()

import io
import openpyxl
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- 3. AGENDA & RUTINAN DZIKIR ---
if menu == "Agenda & Rutinan Dzikir":
    st.title("📅 Agenda & Absensi Rutinan Dzikir Jahar")
    
    tab_ag1, tab_ag3 = st.tabs([
        "Rutinan Dzikir Jahar (Per Kabupaten)", 
        "Rekapitulasi Kehadiran (Rentang Tanggal)"
    ])
    
    with tab_ag1:
        st.subheader("Jadwal & Konfirmasi Kehadiran Dzikir Jahar")
        
        if role == "Ketua Sub Mawil":
            sub_pilih = sub_mawil_aktif_terpilih
        else:
            sub_pilih = st.selectbox("Pilih Sub Mawil (Kabupaten/Kota)", DAFTAR_KAB_KOTA, key="sub_pilih_tab1_sync")
        
        if role == "Ketua Sub Mawil":
            st.success(f"Panel Ketua Sub Mawil **{sub_pilih}**")
            with st.form("form_buat_jadwal"):
                st.markdown("#### 📅 Buat Jadwal Rutinan Baru")
                tgl_buat = st.date_input("Tanggal Pelaksanaan")
                ket_jadwal = st.text_input("Keterangan Jadwal (Cth: Rutinan Malam Jumat / Bulanan)")
                submit_jadwal = st.form_submit_button("Publikasikan Jadwal Rutinan")
                
                if submit_jadwal:
                    cek_duplikat = get_data(
                        "SELECT * FROM jadwal_rutinan WHERE sub_mawil = ? AND tanggal = ? AND keterangan = ?",
                        (sub_pilih, str(tgl_buat), ket_jadwal)
                    )
                    if not cek_duplikat.empty:
                        st.warning("⚠️ Jadwal dengan tanggal dan keterangan tersebut sudah pernah dibuat sebelumnya!")
                    else:
                        execute_query("INSERT INTO jadwal_rutinan (sub_mawil, tanggal, keterangan) VALUES (?, ?, ?)", (sub_pilih, str(tgl_buat), ket_jadwal))
                        st.success(f"Jadwal rutinan untuk tanggal {tgl_buat} berhasil dipublikasikan ke SanFK {sub_pilih}!")
                        st.rerun()

            st.markdown("---")
            st.markdown("#### 🛠️ Kelola Jadwal Rutinan (Edit / Hapus)")
            df_kelola_jadwal = get_data("SELECT id, tanggal, keterangan FROM jadwal_rutinan WHERE sub_mawil = ?", (sub_pilih,))
            
            if df_kelola_jadwal.empty:
                st.info("Belum ada jadwal yang dapat dikelola.")
            else:
                for _, row_j in df_kelola_jadwal.iterrows():
                    with st.expander(f"Jadwal: {row_j['tanggal']} - {row_j['keterangan']}"):
                        with st.form(f"form_edit_jadwal_{row_j['id']}"):
                            edit_tgl = st.date_input("Ubah Tanggal", value=datetime.strptime(row_j['tanggal'], "%Y-%m-%d").date(), key=f"tgl_{row_j['id']}")
                            edit_ket = st.text_input("Ubah Keterangan", value=row_j['keterangan'], key=f"ket_{row_j['id']}")
                            
                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                btn_update = st.form_submit_button("💾 Perbarui Jadwal")
                            with col_btn2:
                                btn_delete = st.form_submit_button("🗑️ Hapus Jadwal")
                                
                            if btn_update:
                                execute_query("UPDATE jadwal_rutinan SET tanggal = ?, keterangan = ? WHERE id = ?", (str(edit_tgl), edit_ket, row_j['id']))
                                st.success("Jadwal berhasil diperbarui!")
                                st.rerun()
                                
                            if btn_delete:
                                execute_query("DELETE FROM jadwal_rutinan WHERE id = ?", (row_j['id'],))
                                st.success("Jadwal berhasil dihapus!")
                                st.rerun()

        st.markdown("---")
        
        df_jadwal_sub = get_data("SELECT * FROM jadwal_rutinan WHERE sub_mawil = ?", (sub_pilih,))
        
        if df_jadwal_sub.empty:
            st.warning(f"Belum ada jadwal rutinan yang dibuat oleh Ketua Sub Mawil untuk wilayah {sub_pilih}.")
        else:
            pilihan_jadwal = ["-- Silakan Pilih Tanggal Jadwal --"] + [f"{row['tanggal']} - {row['keterangan']}" for _, row in df_jadwal_sub.iterrows()]
            tanggal_terpilih_str = st.selectbox("Pilih Jadwal Tanggal Pelaksanaan Resmi:", pilihan_jadwal, key="pilih_tgl_jadwal_resmi")
            
            if tanggal_terpilih_str == "-- Silakan Pilih Tanggal Jadwal --":
                st.info("Silakan pilih salah satu jadwal tanggal di atas untuk melihat detail presensi dan RSVP.")
            else:
                tanggal_aktif = tanggal_terpilih_str.split(" - ")[0]

                # --- BAGIAN KONFIRMASI KEHADIRAN (RSVP) ---
                if role == "SanFK":
                    st.markdown(f"### Konfirmasi Kehadiran (RSVP) untuk Rutinan **{sub_pilih}** tanggal **{tanggal_aktif}**")
                    
                    pilih_nama = sanfk_aktif_terpilih if sanfk_aktif_terpilih else st.session_state.get("sb_sanfk_aktif", "")
                    
                    if not pilih_nama:
                        st.warning("⚠️ Harap pilih profil Anda (SanFK) terlebih dahulu di sidebar.")
                    else:
                        st.info(f"Anda masuk sebagai SanFK: **{pilih_nama}**")
                        
                        with st.form("form_rsvp_sanfk_aktif"):
                            status_rsvp = st.radio("Status Kehadiran:", ["Inshaa Allah akan Hadir", "Tidak Bisa Hadir", "Belum Dapat Dipastikan"], key="status_rsvp_pilihan_sanfk")
                            
                            col_r1, col_r2 = st.columns(2)
                            with col_r1:
                                submit_rsvp = st.form_submit_button("Kirim Konfirmasi Kehadiran")
                            with col_r2:
                                submit_batal = st.form_submit_button("🗑️ Batalkan Konfirmasi")
                            
                            if submit_rsvp:
                                df_asal_rsvp = get_data("SELECT sub_mawil FROM anggota WHERE nama = ?", (pilih_nama,))
                                sub_mawil_asal_rsvp = df_asal_rsvp.iloc[0]['sub_mawil'] if not df_asal_rsvp.empty else sub_pilih

                                cek_rsvp = get_data("SELECT * FROM rsvp WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (pilih_nama, tanggal_aktif, sub_pilih))
                                if not cek_rsvp.empty:
                                    execute_query("UPDATE rsvp SET status_rsvp = ? WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (status_rsvp, pilih_nama, tanggal_aktif, sub_pilih))
                                    st.success(f"Konfirmasi kehadiran {pilih_nama} berhasil diperbarui menjadi '{status_rsvp}'!")
                                else:
                                    execute_query("INSERT INTO rsvp (nama, sub_mawil, tanggal, status_rsvp) VALUES (?, ?, ?, ?)", (pilih_nama, sub_mawil_asal_rsvp, tanggal_aktif, sub_pilih))
                                    st.success(f"Terima kasih {pilih_nama}, konfirmasi Anda ('{status_rsvp}') telah dicatat!")
                                st.rerun()
                                
                            if submit_batal:
                                execute_query("DELETE FROM rsvp WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (pilih_nama, tanggal_aktif, sub_pilih))
                                st.success(f"Konfirmasi kehadiran untuk {pilih_nama} berhasil dibatalkan/dihapus!")
                                st.rerun()

                elif role != "Ketua Sub Mawil":
                    st.markdown(f"### Konfirmasi Kehadiran (RSVP) untuk Rutinan **{sub_pilih}** tanggal **{tanggal_aktif}**")
                    
                    df_angg_all_kab = get_data("SELECT nama, sub_mawil FROM anggota")
                    if not df_angg_all_kab.empty:
                        anggota_terdaftar = [f"{row['nama']} ({row['sub_mawil']})" for _, row in df_angg_all_kab.iterrows()]
                    else:
                        anggota_terdaftar = []
                    
                    if not anggota_terdaftar:
                        st.warning("⚠️ Belum ada data SanFK terdaftar di database.")
                    else:
                        with st.form("form_rsvp_semua_role"):
                            pilih_nama_full = st.selectbox("Pilih Nama (Sesuai Data Terdaftar):", options=anggota_terdaftar, key="pilih_nama_rsvp_all")
                            pilih_nama = pilih_nama_full.split(" (")[0] if pilih_nama_full else ""
                            
                            status_rsvp = st.radio("Status Kehadiran:", ["Inshaa Allah akan Hadir", "Tidak Bisa Hadir", "Belum Dapat Dipastikan"], key="status_rsvp_pilihan")
                            
                            col_r1, col_r2 = st.columns(2)
                            with col_r1:
                                submit_rsvp = st.form_submit_button("Kirim Konfirmasi Kehadiran")
                            with col_r2:
                                submit_batal = st.form_submit_button("🗑️ Batalkan Konfirmasi")
                            
                            if submit_rsvp and pilih_nama:
                                df_asal_rsvp = get_data("SELECT sub_mawil FROM anggota WHERE nama = ?", (pilih_nama,))
                                sub_mawil_asal_rsvp = df_asal_rsvp.iloc[0]['sub_mawil'] if not df_asal_rsvp.empty else sub_pilih

                                cek_rsvp = get_data("SELECT * FROM rsvp WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (pilih_nama, tanggal_aktif, sub_pilih))
                                if not cek_rsvp.empty:
                                    execute_query("UPDATE rsvp SET status_rsvp = ? WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (status_rsvp, pilih_nama, tanggal_aktif, sub_pilih))
                                    st.success(f"Konfirmasi kehadiran {pilih_nama} berhasil diperbarui menjadi '{status_rsvp}'!")
                                else:
                                    execute_query("INSERT INTO rsvp (nama, sub_mawil, tanggal, status_rsvp) VALUES (?, ?, ?, ?)", (pilih_nama, sub_mawil_asal_rsvp, tanggal_aktif, sub_pilih))
                                    st.success(f"Terima kasih {pilih_nama}, konfirmasi Anda ('{status_rsvp}') telah dicatat!")
                                st.rerun()
                                
                            if submit_batal and pilih_nama:
                                execute_query("DELETE FROM rsvp WHERE nama = ? AND tanggal = ? AND sub_mawil = ?", (pilih_nama, tanggal_aktif, sub_pilih))
                                st.success(f"Konfirmasi kehadiran untuk {pilih_nama} berhasil dibatalkan/dihapus!")
                                st.rerun()
                else:
                    st.markdown(f"### Konfirmasi Kehadiran (RSVP) untuk Rutinan **{sub_pilih}** tanggal **{tanggal_aktif}**")
                    st.info("💡 *(ketua sub mawil dapat memantau jadwal di bawah)*")

                st.markdown("---")
                st.subheader("Daftar Konfirmasi SanFK (RSVP Sementara):")
                
                query_rsvp_tampil = """
                    SELECT 
                        r.nama as 'Nama', 
                        m.sub_mawil as 'Sub Mawil', 
                        r.tanggal as 'Tanggal', 
                        r.status_rsvp as 'Status RSVP' 
                    FROM rsvp r
                    JOIN anggota m ON r.nama = m.nama
                    WHERE r.tanggal = ?
                """
                df_rsvp_tampil = get_data(query_rsvp_tampil, (tanggal_aktif,))
                
                if not df_rsvp_tampil.empty:
                    df_rsvp_tampil.index = range(1, len(df_rsvp_tampil) + 1)
                    st.dataframe(df_rsvp_tampil, use_container_width=True)
                else:
                    st.info("Belum ada konfirmasi kehadiran dari SanFK untuk jadwal tanggal jadwal ini.")

                if role == "Ketua Sub Mawil":
                    st.markdown("---")
                    st.markdown(f"#### 🔒 Input & Revisi Aktual Kehadiran (Khusus Ketua Sub Mawil - {tanggal_aktif})")
                    
                    df_angg_all_sub = get_data("SELECT nama, sub_mawil FROM anggota")
                    if not df_angg_all_sub.empty:
                        anggota_kab = [f"{row['nama']} ({row['sub_mawil']})" for _, row in df_angg_all_sub.iterrows()]
                    else:
                        anggota_kab = []

                    with st.form("form_aktual_submawil"):
                        if not anggota_kab:
                            st.warning("Belum ada data SanFK terdaftar di database.")
                            hadir_final_full = []
                        else:
                            hadir_final_full = st.multiselect("Pilih SanFK yang Hadir Aktual di Majelis:", options=anggota_kab, key="multiselect_hadir_aktual")
                        
                        submit_aktual = st.form_submit_button("Simpan & Kunci Aktual Kehadiran")
                        if submit_aktual and hadir_final_full:
                            execute_query("DELETE FROM aktual_hadir WHERE sub_mawil = ? AND tanggal = ?", (sub_pilih, tanggal_aktif))
                            waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M")
                            
                            for item_h in hadir_final_full:
                                nama_h = item_h.split(" (")[0] if " (" in item_h else item_h
                                execute_query("INSERT INTO aktual_hadir (nama, sub_mawil, tanggal, waktu_input) VALUES (?, ?, ?, ?)", (nama_h, sub_pilih, tanggal_aktif, waktu_sekarang))
                            
                            st.success(f"Aktual kehadiran pada tanggal {tanggal_aktif} berhasil disimpan di wilayah {sub_pilih}!")
                            st.rerun()

                    st.markdown("---")
                    st.subheader(f"📋 **Actual List Kehadiran Majelis & Opsi Revisi/Hapus** ({sub_pilih} - {tanggal_aktif})")
                    
                    query_aktual_tampil = """
                        SELECT 
                            id, 
                            nama as 'Nama', 
                            sub_mawil as 'Sub Mawil Majelis', 
                            tanggal as 'Tanggal', 
                            waktu_input as 'Waktu Input' 
                        FROM aktual_hadir 
                        WHERE sub_mawil = ? AND tanggal = ?
                    """
                    df_aktual_tampil = get_data(query_aktual_tampil, (sub_pilih, tanggal_aktif))
                    
                    if not df_aktual_tampil.empty:
                        df_aktual_tampil_clean = df_aktual_tampil.drop(columns=['id'])
                        df_aktual_tampil_clean.index = range(1, len(df_aktual_tampil_clean) + 1)
                        st.dataframe(df_aktual_tampil_clean, use_container_width=True)
                        
                        st.markdown("##### 🗑️ Hapus / Koreksi Data SanFK yang Salah Masuk Hadir:")
                        with st.form("form_hapus_aktual"):
                            pilih_hapus_aktual = st.selectbox("Pilih Nama SanFK yang Ingin Dihapus dari Kehadiran Aktual:", options=df_aktual_tampil['Nama'].tolist(), key="pilih_hapus_akt_box")
                            btn_eksekusi_hapus = st.form_submit_button("Hapus Nama Terpilih dari Daftar Hadir Aktual")
                            
                            if btn_eksekusi_hapus and pilih_hapus_aktual:
                                execute_query("DELETE FROM aktual_hadir WHERE sub_mawil = ? AND tanggal = ? AND nama = ?", (sub_pilih, tanggal_aktif, pilih_hapus_aktual))
                                st.success(f"Nama **{pilih_hapus_aktual}** berhasil dihapus dari daftar kehadiran aktual!")
                                st.rerun()
                    else:
                        st.info("Belum ada data actual list kehadiran yang dikunci untuk jadwal tanggal ini.")

    with tab_ag3:
        st.subheader("📊 Rekapitulasi Aktual Kehadiran (Keseluruhan & Per Sub Mawil)")
        st.info("Fitur ini dapat diakses oleh semua role untuk memantau rekapitulasi kehadiran berdasarkan rentang tanggal dan wilayah.")

        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            filter_tipe_wilayah = st.selectbox("Cakupan Wilayah", ["Semua Wilayah (Keseluruhan)"] + DAFTAR_KAB_KOTA, key="filter_cakupan_wil_rekap")
        with col_f2:
            tgl_mulai = st.date_input("Dari Tanggal", value=date.today().replace(day=1), key="rekap_tgl_mulai")
        with col_f3:
            tgl_selesai = st.date_input("Sampai Tanggal", value=date.today(), key="rekap_tgl_selesai")

        st.markdown("---")
        
        query_rekap = """
            SELECT 
                a.id,
                a.nama as 'Nama', 
                m.sub_mawil as 'Sub Mawil Asal', 
                a.sub_mawil as 'Wilayah Majelis',
                a.tanggal as 'Tanggal', 
                a.waktu_input as 'Waktu Input' 
            FROM aktual_hadir a
            JOIN anggota m ON a.nama = m.nama
        """
        df_rekap = get_data(query_rekap)
        
        if df_rekap.empty:
            st.warning("Belum ada data actual kehadiran yang tersimpan di database.")
        else:
            df_rekap["Tanggal_Obj"] = pd.to_datetime(df_rekap["Tanggal"]).dt.date
            df_rekap = df_rekap[(df_rekap["Tanggal_Obj"] >= tgl_mulai) & (df_rekap["Tanggal_Obj"] <= tgl_selesai)]
            
            if filter_tipe_wilayah != "Semua Wilayah (Keseluruhan)":
                df_rekap = df_rekap[df_rekap["Wilayah Majelis"] == filter_tipe_wilayah]
            
            df_rekap_tampil = df_rekap[['Nama', 'Sub Mawil Asal', 'Wilayah Majelis', 'Tanggal', 'Waktu Input']].copy()

            st.markdown(f"#### Hasil Rekapitulasi: `{filter_tipe_wilayah}` *(Periode: {tgl_mulai} s.d. {tgl_selesai})*")
            st.metric("Total Kehadiran Tercatat", len(df_rekap_tampil))
            
            if not df_rekap_tampil.empty:
                df_rekap_tampil.index = range(1, len(df_rekap_tampil) + 1)
                st.dataframe(df_rekap_tampil, use_container_width=True)
                
                st.markdown("---")
                st.markdown("##### 📥 Unduh Laporan Rekapitulasi Kehadiran:")
                
                col_dl2, col_dl3 = st.columns(2)
                
                # 1. Tombol Download Excel (.xlsx) menggunakan openpyxl
                with col_dl2:
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                        df_rekap_tampil.to_excel(writer, index=False, sheet_name='Rekap Kehadiran')
                    excel_data = output_excel.getvalue()
                    
                    st.download_button(
                        label="📊 Unduh Format Excel (.xlsx)",
                        data=excel_data,
                        file_name=f"rekap_kehadiran_{filter_tipe_wilayah}_{tgl_mulai}_sd_{tgl_selesai}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_download_excel"
                    )

                # 2. Tombol Download Word (.docx) menggunakan python-docx
                with col_dl3:
                    doc = Document()
                    
                    # Judul Dokumen Word
                    p_title = doc.add_paragraph()
                    run_title = p_title.add_run("LAPORAN REKAPITULASI KEHADIRAN SANFK")
                    run_title.bold = True
                    run_title.font.size = Pt(14)
                    run_title.font.color.rgb = RGBColor(14, 102, 85) # Warna hijau tua
                    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    # Subjudul Periode & Wilayah
                    p_sub = doc.add_paragraph()
                    p_sub.add_run(f"Cakupan Wilayah: {filter_tipe_wilayah}\nPeriode: {tgl_mulai} s.d. {tgl_selesai}\nTotal Kehadiran: {len(df_rekap_tampil)} Orang")
                    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    doc.add_paragraph() # Spasi
                    
                    # Buat Tabel di Word
                    table = doc.add_table(rows=1, cols=5)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    table.style = 'Table Grid'
                    
                    hdr_cells = table.rows[0].cells
                    headers = ['No', 'Nama', 'Sub Mawil Asal', 'Wilayah Majelis', 'Tanggal']
                    for i, h_text in enumerate(headers):
                        hdr_cells[i].text = h_text
                        for paragraph in hdr_cells[i].paragraphs:
                            for run in paragraph.runs:
                                run.bold = True
                                run.font.color.rgb = RGBColor(255, 255, 255)
                        # Beri warna background header tabel (hijau tua)
                        shading_elm = parse_xml(r'<w:shd {} w:fill="0E6655"/>'.format(nsdecls('w')))
                        hdr_cells[i]._tc.get_or_add_tcPr().append(shading_elm)

                    # Masukkan data ke baris tabel Word
                    for idx, row in df_rekap_tampil.iterrows():
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(idx)
                        row_cells[1].text = str(row['Nama'])
                        row_cells[2].text = str(row['Sub Mawil Asal'])
                        row_cells[3].text = str(row['Wilayah Majelis'])
                        row_cells[4].text = str(row['Tanggal'])

                    doc.add_paragraph()
                    doc.add_paragraph("Demikian laporan rekapitulasi ini disusun untuk dipergunakan sebagaimana mestinya.\n\n\nKetua Mawil / Pengurus")

                    # Simpan ke BytesIO untuk tombol download Streamlit
                    output_word = io.BytesIO()
                    doc.save(output_word)
                    word_data = output_word.getvalue()

                    st.download_button(
                        label="📝 Unduh Format Word (.docx)",
                        data=word_data,
                        file_name=f"laporan_rekap_kehadiran_{filter_tipe_wilayah}_{tgl_mulai}_sd_{tgl_selesai}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_download_word"
                    )
            else:
                st.info("Tidak ada data kehadiran yang ditemukan pada rentang tanggal dan wilayah tersebut.")

# --- 7. GALERI & FEED UMUM ---
elif menu == "Galeri & Feed Umum":
    st.title("🌐 Galeri & Feed Umum (Silaturrahmi SanFK)")
    st.write("Ruang berbagi tulisan, foto, dan aktivitas antar-SanFK lintas daerah.")
    st.write("Catatan : Pengurus memiliki kewenangan menghapus postingan sanFK yang dinilai tidak sesuai, tanpa pemberitahuan terlebih dahulu.")
    
    if role not in ["SanFK", "Sekretaris Mawil"]:
        st.warning("⚠️ Menu **Galeri & Feed Umum** ini hanya dapat diakses oleh role **SanFK** dan **Sekretaris Mawil** sesuai pengaturan hak akses yang telah ditetapkan.")
    else:
        if role == "Sekretaris Mawil":
            st.subheader("🛠️ Panel Moderator Sekretaris Mawil (Manajemen Postingan)")
            df_galeri = get_data("SELECT * FROM galeri_umum ORDER BY id DESC")
            
            if df_galeri.empty:
                st.info("Belum ada postingan di galeri komunitas.")
            else:
                for _, post in df_galeri.iterrows():
                    with st.container():
                        post_id = post['id']
                        st.markdown(f"**{post['penulis']}** *({post['sub_mawil']})* - <small>{post['waktu']}</small>", unsafe_allow_html=True)
                        
                        if st.button(f"🗑️ [Otoritas Sekretaris] Hapus Postingan Ini", key=f"btn_sekre_del_post_{post_id}"):
                            execute_query("DELETE FROM galeri_umum WHERE id = ?", (post_id,))
                            execute_query("DELETE FROM reaksi_posting WHERE post_id = ?", (post_id,))
                            execute_query("DELETE FROM komentar_posting WHERE post_id = ?", (post_id,))
                            st.success("Postingan berhasil dihapus oleh Sekretaris Mawil!")
                            st.rerun()

                        fg_str = post.get('foto_galeri', '')
                        arr_foto_g = [x.strip() for x in fg_str.split(',') if x.strip()]
                        
                        if arr_foto_g:
                            cols_img_g = st.columns(min(len(arr_foto_g), 3))
                            for i, f_item_str_g in enumerate(arr_foto_g):
                                if "|" in f_item_str_g:
                                    f_path_g, akses_file_item_g = f_item_str_g.split("|", 1)
                                else:
                                    f_path_g, akses_file_item_g = f_item_str_g, "Public"
                                    
                                ext_fg = f_path_g.split('.')[-1].lower().split('?')[0]
                                with cols_img_g[i % 3]:
                                    if akses_file_item_g == "Private":
                                        st.caption(f"🔒 File {i+1}: **Tersimpan sebagai Private**")
                                    else:
                                        st.caption(f"🌍 File {i+1}: **Tersimpan sebagai Public**")
                                    
                                    if ext_fg in ['jpg', 'jpeg', 'png', 'webp']:
                                        st.markdown(f'<img src="{f_path_g}" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 6px; margin-bottom: 6px; border: 1px solid #ddd;">', unsafe_allow_html=True)
                                    elif ext_fg in ['mp4', 'mov', 'avi']:
                                        st.video(f_path_g)
                                    elif ext_fg in ['mp3', 'wav', 'ogg', 'm4a']:
                                        st.audio(f_path_g)
                                    else:
                                        st.info(f"📄 Berkas / Dokumen Cloud")
                                    
                                    if st.button(f"🗑️ Hapus File {i+1} Ini", key=f"btn_sekre_del_file_{post_id}_{i}"):
                                        arr_foto_g.pop(i)
                                        new_fg_str_updated = ",".join(arr_foto_g)
                                        execute_query("UPDATE galeri_umum SET foto_galeri = ? WHERE id = ?", (new_fg_str_updated, post_id))
                                        st.success(f"File {i+1} berhasil dihapus oleh Sekretaris Mawil!")
                                        st.rerun()
                        
                        if post['konten']:
                            st.write(post['konten'])

                        with st.expander("💬 Kelola Komentar Postingan Ini"):
                            df_komentar_sekre = get_data("SELECT id, nama_sanfk, waktu, komentar FROM komentar_posting WHERE post_id = ? ORDER BY id DESC", (post_id,))
                            if df_komentar_sekre.empty:
                                st.info("Belum ada komentar pada postingan ini.")
                            else:
                                for _, k_row in df_komentar_sekre.iterrows():
                                    k_id = k_row['id']
                                    st.markdown(f"**{k_row['nama_sanfk']}** <small>({k_row['waktu']})</small>", unsafe_allow_html=True)
                                    st.write(k_row['komentar'])
                                    if st.button("🗑️ Hapus Komentar Ini", key=f"btn_sekre_del_kom_{k_id}"):
                                        execute_query("DELETE FROM komentar_posting WHERE id = ?", (k_id,))
                                        st.success("Komentar berhasil dihapus oleh Sekretaris Mawil!")
                                        st.rerun()
                                    st.markdown("---")

                        st.divider()
        else:
            tab_gal_1, tab_gal_2 = st.tabs(["📋 Tampilan Postingan", "➕ Buat & Kelola Postingan"])
            
            with tab_gal_1:
                st.subheader("Timeline Komunitas")
                df_galeri = get_data("SELECT * FROM galeri_umum ORDER BY id DESC")
                
                if df_galeri.empty:
                    st.info("Belum ada postingan di galeri komunitas.")
                else:
                    for _, post in df_galeri.iterrows():
                        with st.container():
                            st.markdown(f"**{post['penulis']}** *({post['sub_mawil']})* - <small>{post['waktu']}</small>", unsafe_allow_html=True)
                            
                            fg_str = post.get('foto_galeri', '')
                            arr_foto_g = [x.strip() for x in fg_str.split(',') if x.strip()]
                            
                            if arr_foto_g:
                                cols_img_g = st.columns(min(len(arr_foto_g), 3))
                                for i, f_item_str_g in enumerate(arr_foto_g):
                                    if "|" in f_item_str_g:
                                        f_path_g, akses_file_item_g = f_item_str_g.split("|", 1)
                                    else:
                                        f_path_g, akses_file_item_g = f_item_str_g, "Public"
                                        
                                    ext_fg = f_path_g.split('.')[-1].lower().split('?')[0]
                                    with cols_img_g[i % 3]:
                                        if akses_file_item_g == "Private":
                                            st.caption(f"🔒 File {i+1}: **Tersimpan sebagai Private**")
                                        else:
                                            st.caption(f"🌍 File {i+1}: **Tersimpan sebagai Public**")
                                        
                                        if ext_fg in ['jpg', 'jpeg', 'png', 'webp']:
                                            st.markdown(f'<img src="{f_path_g}" style="width: 100%; max-height: 200px; object-fit: cover; border-radius: 6px; margin-bottom: 6px; border: 1px solid #ddd;">', unsafe_allow_html=True)
                                        elif ext_fg in ['mp4', 'mov', 'avi']:
                                            st.video(f_path_g)
                                        elif ext_fg in ['mp3', 'wav', 'ogg', 'm4a']:
                                            st.audio(f_path_g)
                                        else:
                                            st.info(f"📄 Berkas / Dokumen Cloud")
                                        
                                        can_download_g = (akses_file_item_g == "Public") or (role == "Admin Dokumentasi Mawil") or (role == post['penulis'])
                                        if can_download_g:
                                            st.markdown(f"[📥 Download / Buka File {i+1}]({f_path_g})", unsafe_allow_html=True)
                                        else:
                                            st.warning("🔒 File Private (Akses Dibatasi)")
                            
                            if post['konten']:
                                st.write(post['konten'])

                            post_id = post['id']
                            df_reaksi_post = get_data("SELECT reaksi, COUNT(*) as jml FROM reaksi_posting WHERE post_id = ? GROUP BY reaksi", (post_id,))
                            count_like = int(df_reaksi_post[df_reaksi_post['reaksi'] == 'Like']['jml'].values[0]) if not df_reaksi_post[df_reaksi_post['reaksi'] == 'Like'].empty else 0
                            count_dislike = int(df_reaksi_post[df_reaksi_post['reaksi'] == 'Dislike']['jml'].values[0]) if not df_reaksi_post[df_reaksi_post['reaksi'] == 'Dislike'].empty else 0
                            count_love = int(df_reaksi_post[df_reaksi_post['reaksi'] == 'Love']['jml'].values[0]) if not df_reaksi_post[df_reaksi_post['reaksi'] == 'Love'].empty else 0

                            st.markdown(f"👍 **{count_like}** Likes | 👎 **{count_dislike}** Dislikes | ❤️ **{count_love}** Loves")

                            if role == "SanFK" and sanfk_aktif_terpilih:
                                cols_reaksi = st.columns(4)
                                with cols_reaksi[0]:
                                    if st.button("👍 Like", key=f"btn_like_{post_id}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_posting (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id, sanfk_aktif_terpilih, "Like"))
                                        st.rerun()
                                with cols_reaksi[1]:
                                    if st.button("👎 Dislike", key=f"btn_dislike_{post_id}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_posting (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id, sanfk_aktif_terpilih, "Dislike"))
                                        st.rerun()
                                with cols_reaksi[2]:
                                    if st.button("❤️ Love", key=f"btn_love_{post_id}"):
                                        execute_query("INSERT OR REPLACE INTO reaksi_posting (post_id, nama_sanfk, reaksi) VALUES (?, ?, ?)", (post_id, sanfk_aktif_terpilih, "Love"))
                                        st.rerun()
                                with cols_reaksi[3]:
                                    if st.button("❌ Batal Reaksi", key=f"btn_unreact_{post_id}"):
                                        execute_query("DELETE FROM reaksi_posting WHERE post_id = ? AND nama_sanfk = ?", (post_id, sanfk_aktif_terpilih))
                                        st.rerun()

                            with st.expander(f"💬 Kolom Komentar (Multi Komentar)"):
                                df_komentar = get_data("SELECT id, nama_sanfk, waktu, komentar FROM komentar_posting WHERE post_id = ? ORDER BY id DESC", (post_id,))
                                
                                if role == "SanFK" and sanfk_aktif_terpilih:
                                    with st.form(f"form_tambah_kom_{post_id}", clear_on_submit=True):
                                        text_komentar_baru = st.text_area("Tulis Komentar Baru:", key=f"input_kom_baru_{post_id}")
                                        btn_kirim_kom = st.form_submit_button("Kirim Komentar")
                                        if btn_kirim_kom and text_komentar_baru:
                                            waktu_k = datetime.now().strftime("%Y-%m-%d %H:%M")
                                            execute_query("INSERT INTO komentar_posting (post_id, nama_sanfk, waktu, komentar) VALUES (?, ?, ?, ?)", (post_id, sanfk_aktif_terpilih, waktu_k, text_komentar_baru))
                                            st.success("Komentar berhasil dikirim!")
                                            st.rerun()
                                else:
                                    st.caption("🔒 *Login sebagai SanFK di sidebar untuk ikut menulis komentar.*")

                                st.markdown("---")
                                if df_komentar.empty:
                                    st.info("Belum ada komentar.")
                                else:
                                    for _, k_row in df_komentar.iterrows():
                                        k_id = k_row['id']
                                        k_nama = k_row['nama_sanfk']
                                        k_waktu = k_row['waktu']
                                        k_isi = k_row['komentar']

                                        st.markdown(f"**{k_nama}** <small>({k_waktu})</small>", unsafe_allow_html=True)
                                        st.write(k_isi)

                                        if role == "SanFK" and sanfk_aktif_terpilih == k_nama:
                                            with st.expander("⋮ Menu Opsi"):
                                                col_k1, col_k2 = st.columns(2)
                                                with col_k1:
                                                    with st.expander("✏️ Edit Komentar Ini", key=f"exp_edit_kom_{k_id}"):
                                                        with st.form(f"form_edit_kom_{k_id}"):
                                                            edit_isi_kom = st.text_area("Ubah Komentar:", value=k_isi, key=f"val_edit_kom_{k_id}")
                                                            btn_simpan_edit_kom = st.form_submit_button("Simpan Perubahan")
                                                            if btn_simpan_edit_kom:
                                                                execute_query("UPDATE komentar_posting SET komentar = ? WHERE id = ?", (edit_isi_kom, k_id))
                                                                st.success("Komentar berhasil diperbarui!")
                                                                st.rerun()
                                                with col_k2:
                                                    if st.button("🗑️ Hapus Komentar", key=f"btn_hapus_kom_{k_id}"):
                                                        execute_query("DELETE FROM komentar_posting WHERE id = ?", (k_id,))
                                                        st.success("Komentar berhasil dihapus!")
                                                        st.rerun()

                                        st.markdown("---")
                            st.divider()

            with tab_gal_2:
                if role != "SanFK":
                    st.warning("⚠️ Menu **Buat & Kelola Postingan** ini khusus diakses oleh **SanFK**.")
                else:
                    st.subheader("Buat Postingan Baru")
                    df_all_sanfk_gal = get_data("SELECT nama, sub_mawil FROM anggota")
                    
                    if df_all_sanfk_gal.empty:
                        st.warning("⚠️ Belum ada data SanFK terdaftar. Silakan input data SanFK melalui menu Manajemen SanFK terlebih dahulu.")
                    else:
                        if sanfk_aktif_terpilih:
                            penulis_terpilih = sanfk_aktif_terpilih
                            st.info(f"👤 Penulis (Terkunci ke Akun Anda): **{penulis_terpilih}**")
                            
                            row_sanfk_aktif = df_all_sanfk_gal[df_all_sanfk_gal['nama'] == penulis_terpilih].iloc[0]
                            wilayah = row_sanfk_aktif['sub_mawil']
                            st.info(f"📍 Asal Sub Mawil (Otomatis): **{wilayah}**")
                            
                            konten = st.text_area("Tulis pesan, mutiara hikmah, atau informasi kegiatan...", key="gal_konten")
                            
                            st.markdown("---")
                            st.markdown("##### 📁 Lampiran Berbagai File (Bisa Banyak: Foto, PDF, MP3, Video, dll)")
                            metode_foto_g = st.radio("Pilih Cara Input File Galeri:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key="radio_g_metode")
                            
                            foto_galeri_files = None
                            cam_galeri = None
                            if metode_foto_g == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                                foto_galeri_files = st.file_uploader("Pilih File (Bisa Banyak: PDF, MP3, MP4, JPG, PNG, dll)", type=None, accept_multiple_files=True, key="up_g")
                            else:
                                cam_galeri = st.camera_input("Ambil Foto Langsung dengan Kamera", key="cam_g")
                            
                            file_access_settings_g = {}
                            if foto_galeri_files:
                                st.markdown("##### ⚙️ Atur Hak Akses Masing-Masing File:")
                                for idx_fg, fg_item in enumerate(foto_galeri_files):
                                    file_access_settings_g[fg_item.name] = st.selectbox(
                                        f"Hak Akses untuk file: **{fg_item.name}**", 
                                        ["Public", "Private"], 
                                        index=1, 
                                        key=f"akses_file_g_{idx_fg}"
                                    )
                            elif cam_galeri is not None:
                                cam_akses_g = st.selectbox("Hak Akses untuk Foto Kamera:", ["Public", "Private"], index=1, key="akses_cam_g")
                            
                            st.caption("ℹ️ **Public**: Bisa didownload oleh semua role. **Private**: Hanya bisa didownload oleh pemosting.")
                            
                            if st.button("Posting ke Galeri", key="btn_pub_galeri_umum"):
                                if penulis_terpilih and (konten or foto_galeri_files or cam_galeri):
                                    path_list_g = []
                                    if foto_galeri_files:
                                        with st.spinner("Mengunggah file ke Cloudinary..."):
                                            for fg_item in foto_galeri_files:
                                                try:
                                                    upload_fg = cloudinary.uploader.upload(fg_item, resource_type="auto")
                                                    fg_url = upload_fg.get("secure_url")
                                                    
                                                    akses_dipilih_g = file_access_settings_g.get(fg_item.name, "Private")
                                                    path_list_g.append(f"{fg_url}|{akses_dipilih_g}")
                                                except Exception as e:
                                                    st.error(f"Gagal mengunggah {fg_item.name}: {e}")
                                            
                                    if cam_galeri is not None:
                                        with st.spinner("Mengunggah foto kamera ke Cloudinary..."):
                                            try:
                                                upload_cgal = cloudinary.uploader.upload(cam_galeri, resource_type="image")
                                                cgal_url = upload_cgal.get("secure_url")
                                                path_list_g.append(f"{cgal_url}|{cam_akses_g}")
                                            except Exception as e:
                                                st.error(f"Gagal mengunggah foto kamera: {e}")
                                        
                                    foto_g_str = ",".join(path_list_g) if path_list_g else ""
                                    waktu_post = datetime.now().strftime("%Y-%m-%d %H:%M")
                                    execute_query(
                                        "INSERT INTO galeri_umum (penulis, sub_mawil, waktu, konten, foto_galeri, tipe) VALUES (?, ?, ?, ?, ?, ?)", 
                                        (penulis_terpilih, wilayah, waktu_post, konten, foto_g_str, "Postingan")
                                    )
                                    st.success("Postingan dan lampiran file berhasil dibagikan ke Cloudinary!")
                                    st.rerun()
                                else:
                                    st.warning("Konten atau file lampiran wajib diisi!")

                    st.markdown("---")
                    st.subheader("🛠️ Kelola, Edit, atau Hapus Postingan Anda")
                    
                    df_my_post = get_data("SELECT * FROM galeri_umum WHERE penulis = ? ORDER BY id DESC", (sanfk_aktif_terpilih,))
                    st.caption(f"Menampilkan postingan milik Anda: **{sanfk_aktif_terpilih}**")

                    if df_my_post.empty:
                        st.info("Belum ada postingan yang Anda buat.")
                    else:
                        pilihan_mp = {f"[{row['waktu']}] {row['konten'][:30]}...": row['id'] for _, row in df_my_post.iterrows()}
                        pilih_label_mp = st.selectbox("Pilih Postingan Anda untuk Dikelola / Dihapus:", list(pilihan_mp.keys()), key="select_kelola_post")
                        id_mp_aktif = pilihan_mp[pilih_label_mp]
                        
                        data_post_terpilih = get_data("SELECT * FROM galeri_umum WHERE id = ?", (id_mp_aktif,)).iloc[0]
                        
                        fg_lama_str = data_post_terpilih.get('foto_galeri', '')
                        list_fg = [x.strip() for x in fg_lama_str.split(',') if x.strip()]

                        if list_fg:
                            st.markdown("##### 🗑️ Kelola File & Ubah Hak Akses Per File:")
                            for idx_img_g, item_str_g in enumerate(list_fg):
                                if "|" in item_str_g:
                                    path_img_g, akses_file_item_g = item_str_g.split("|", 1)
                                else:
                                    path_img_g, akses_file_item_g = item_str_g, "Public"
                                
                                st.markdown(f"---")
                                ext_file_g = path_img_g.split('.')[-1].lower().split('?')[0]
                                
                                col_pg1, col_pg2 = st.columns([2, 2])
                                with col_pg1:
                                    st.markdown(f"**File {idx_img_g + 1}:** `{path_img_g}`")
                                    if akses_file_item_g == "Private":
                                        st.markdown("🔒 Status: **Tersimpan sebagai Private**")
                                    else:
                                        st.markdown("🌍 Status: **Tersimpan sebagai Public**")

                                    if ext_file_g in ['jpg', 'jpeg', 'png', 'webp']:
                                        st.markdown(f'<img src="{path_img_g}" style="width: 100%; max-height: 120px; object-fit: cover; border-radius: 4px; border: 1px solid #ccc;">', unsafe_allow_html=True)
                                    elif ext_file_g in ['mp4', 'mov', 'avi']:
                                        st.video(path_img_g)
                                    elif ext_file_g in ['mp3', 'wav', 'ogg', 'm4a']:
                                        st.audio(path_img_g)
                                    else:
                                        st.info(f"📄 Berkas / Dokumen Cloud")
                                        
                                with col_pg2:
                                    ubah_akses_item_g = st.selectbox(
                                        f"Ubah Akses File {idx_img_g + 1}", 
                                        ["Public", "Private"], 
                                        index=0 if akses_file_item_g == "Public" else 1, 
                                        key=f"ubah_akses_g_{id_mp_aktif}_{idx_img_g}"
                                    )
                                    
                                    if st.button(f"🗑️ Hapus File Ini", key=f"btn_del_file_g_single_{id_mp_aktif}_{idx_img_g}"):
                                        list_fg.pop(idx_img_g)
                                        new_fg_str = ",".join(list_fg)
                                        execute_query("UPDATE galeri_umum SET foto_galeri = ? WHERE id = ?", (new_fg_str, id_mp_aktif))
                                        st.success(f"File {idx_img_g + 1} berhasil dihapus!")
                                        st.rerun()
                                        
                                    if st.button(f"💾 Simpan Akses File Ini", key=f"btn_save_akses_g_{id_mp_aktif}_{idx_img_g}"):
                                        list_fg[idx_img_g] = f"{path_img_g}|{ubah_akses_item_g}"
                                        new_fg_str = ",".join(list_fg)
                                        execute_query("UPDATE galeri_umum SET foto_galeri = ? WHERE id = ?", (new_fg_str, id_mp_aktif))
                                        st.success(f"Hak akses File {idx_img_g + 1} berhasil diperbarui!")
                                        st.rerun()
                        else:
                            st.info("Tidak ada file yang terlampir pada postingan ini.")

                        st.markdown("---")
                        edit_konten_post = st.text_area("Edit Konten Postingan", value=data_post_terpilih['konten'] if data_post_terpilih['konten'] else "", key=f"ed_konten_post_{id_mp_aktif}")
                        
                        st.markdown("##### 📁 Tambah File Baru")
                        edit_metode_g = st.radio("Pilih Cara Tambah File Baru:", ["Unggah Berbagai File (PDF, Foto, Video, dll)", "Gunakan Kamera Langsung"], horizontal=True, key=f"ed_metode_g_{id_mp_aktif}")
                        edit_foto_files_g = None
                        edit_cam_file_g = None
                        if edit_metode_g == "Unggah Berbagai File (PDF, Foto, Video, dll)":
                            edit_foto_files_g = st.file_uploader("Upload File Baru (Bisa Banyak)", type=None, accept_multiple_files=True, key=f"ed_up_g_{id_mp_aktif}")
                        else:
                            edit_cam_file_g = st.camera_input("Ambil Foto Baru via Kamera", key=f"ed_cam_g_{id_mp_aktif}")
                        
                        edit_file_access_settings_g = {}
                        if edit_foto_files_g:
                            for idx_efg, efg_item in enumerate(edit_foto_files_g):
                                edit_file_access_settings_g[efg_item.name] = st.selectbox(
                                    f"Hak Akses File Baru: **{efg_item.name}**", 
                                    ["Public", "Private"], 
                                    index=1, 
                                    key=f"akses_edit_new_g_{id_mp_aktif}_{idx_efg}"
                                )
                        elif edit_cam_file_g is not None:
                            edit_cam_akses_g = st.selectbox("Hak Akses Foto Kamera Baru:", ["Public", "Private"], index=1, key=f"akses_edit_cam_g_{id_mp_aktif}")
                        
                        col_eg1, col_eg2 = st.columns(2)
                        with col_eg1:
                            if st.button("💾 Simpan Perubahan Teks & Tambah File", key=f"btn_up_post_{id_mp_aktif}"):
                                existing_paths_g = list_fg
                                
                                if edit_foto_files_g:
                                    with st.spinner("Mengunggah file baru ke Cloudinary..."):
                                        for efg_item in edit_foto_files_g:
                                            try:
                                                upload_efg = cloudinary.uploader.upload(efg_item, resource_type="auto")
                                                efg_url = upload_efg.get("secure_url")
                                                
                                                akses_efg = edit_file_access_settings_g.get(efg_item.name, "Private")
                                                existing_paths_g.append(f"{efg_url}|{akses_efg}")
                                            except Exception as e:
                                                st.error(f"Gagal upload {efg_item.name}: {e}")
                                        
                                if edit_cam_file_g is not None:
                                    with st.spinner("Mengunggah foto kamera baru..."):
                                        try:
                                            upload_ecam_g = cloudinary.uploader.upload(edit_cam_file_g, resource_type="image")
                                            ecam_url_g = upload_ecam_g.get("secure_url")
                                            existing_paths_g.append(f"{ecam_url_g}|{edit_cam_akses_g}")
                                        except Exception as e:
                                            st.error(f"Gagal upload kamera: {e}")
                                        
                                foto_g_path_e = ",".join(existing_paths_g)
                                execute_query(
                                    "UPDATE galeri_umum SET konten = ?, foto_galeri = ? WHERE id = ?",
                                    (edit_konten_post, foto_g_path_e, id_mp_aktif)
                                )
                                st.success("Postingan dan file baru berhasil disimpan di Cloudinary!")
                                st.rerun()
                        with col_eg2:
                            if st.button("🗑️ Hapus Postingan Ini Sepenuhnya", key=f"btn_del_post_full_{id_mp_aktif}"):
                                execute_query("DELETE FROM galeri_umum WHERE id = ?", (id_mp_aktif,))
                                st.success("Postingan berhasil dihapus dari sistem!")
                                st.rerun()
                             
# --- 6. LAYANAN SANTUNAN & KONTAK ---
elif menu == "Layanan Santunan & Kontak":
    st.title("🤝 Layanan Santunan Sosial & Kontak Darurat")
    
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        st.subheader("🏥 Santunan Sakit / Kedukaan")
        st.markdown("""
        Program santunan sosial diberikan kepada SanFK atau keluarga inti yang mengalami musibah (sakit keras/rawat inap atau meninggal dunia).
        * **Ketentuan:** Berlaku bagi seluruh SanFK terdaftar.
        * **Mekanisme Usulan:** Dikoordinasikan melalui Ketua Sub Mawil setempat untuk diteruskan ke pengurus **Mawil Provinsi**.
        """)
    with col_s2:
        st.subheader("📦 Baksos Tahunan dari Pusat")
        st.markdown("""
        Informasi terkait bantuan bakti sosial skala tahunan yang anggurannya bersumber langsung dari Pusat:
        * Berdasarkan usulan proposal resmi dari Mawil.
        * Penyaluran dikoordinasikan langsung ke masing-masing titik Sub Mawil penerima manfaat.
        """)

    st.markdown("---")
    st.subheader("📞 Narahubung & Kontak Pengurus Mawil")
    
    # Ambil data pengurus dari database
    df_kontak_pengurus = get_data("SELECT jabatan, nama_pejabat, kontak FROM struktur_pengurus")
    
    def cari_pengurus(keyword):
        if df_kontak_pengurus.empty:
            return "(Belum Ditetapkan)", "-"
        for _, row in df_kontak_pengurus.iterrows():
            jabatan_db = str(row['jabatan']).lower()
            if keyword.lower() in jabatan_db:
                nama = row['nama_pejabat'] if row['nama_pejabat'] and row['nama_pejabat'].strip() != "" else "(Belum Ditetapkan)"
                kontak = row['kontak'] if row['kontak'] and row['kontak'].strip() != "" else "-"
                return nama, kontak
        return "(Belum Ditetapkan)", "-"

    ketua_nama, ketua_kontak = cari_pengurus("Ketua")
    sekre_nama, sekre_kontak = cari_pengurus("Sekretaris")
    benda_nama, benda_kontak = cari_pengurus("Bendahara")

    # Fungsi format kontak tanpa batasan huruf 'x'
    def format_kontak_html(label_jabatan, nama, kontak):
        if kontak and kontak != "-" and kontak.strip() != "":
            import re
            cleaned_digits = re.sub(r'\D', '', kontak)
            
            # Jika nomor memiliki digit yang cukup untuk WhatsApp
            if len(cleaned_digits) >= 9:
                if cleaned_digits.startswith('0'):
                    wa_number = '62' + cleaned_digits[1:]
                elif cleaned_digits.startswith('62'):
                    wa_number = cleaned_digits
                else:
                    wa_number = '62' + cleaned_digits
                    
                pesan_wa = f"Halo {nama}, saya ingin berkonsultasi mengenai Layanan Santunan FK Mawil Riau."
                wa_link = f"https://wa.me/{wa_number}?text={pesan_wa.replace(' ', '%20')}"
                
                kontak_str = f'<a href="{wa_link}" target="_blank" style="color: #0E6655; text-decoration: none; font-weight: bold;">📞 {kontak}</a>'
            else:
                # Jika format nomor tidak valid atau berupa teks biasa
                kontak_str = f'<span>{kontak}</span>'
        else:
            kontak_str = '<span><i>(Belum ada kontak)</i></span>'
            
        return f"<li><b>{label_jabatan}:</b> {nama} — {kontak_str}</li>"

    html_kontak_items = ""
    html_kontak_items += format_kontak_html("Ketua Mawil Riau", ketua_nama, ketua_kontak)
    html_kontak_items += format_kontak_html("Sekretaris Mawil (Administrasi)", sekre_nama, sekre_kontak)
    html_kontak_items += format_kontak_html("Bendahara Mawil (Keuangan & Kotak Hijau)", benda_nama, benda_kontak)

    st.markdown(f"""
    <div style="background-color: #f0f6fc; padding: 20px; border-radius: 10px; border: 1px solid #d0e1fd; margin-bottom: 15px;">
        <ul style="list-style-type: disc; margin-left: 20px; color: #1f2937; line-height: 1.8;">
            {html_kontak_items}
        </ul>
    </div>
    """, unsafe_allow_html=True)

# --- 2. MANAJEMEN SANFK & KTA ---
elif menu == "Manajemen SanFK & KTA":
    st.title("👥 Manajemen Data SanFK, Ijazah Dzikir, & KTA Digital")
    
    # --- PENGATURAN TAB BERDASARKAN ROLE ---
    if role == "Sekretaris Mawil":
        tab1, tab2, tab3 = st.tabs(["Daftar SanFK", "Tambah / Perbarui Data", "KTA Digital & QR Code"])
    else:
        tab1, tab3 = st.tabs(["Daftar SanFK", "KTA Digital & QR Code"])
        tab2 = None
    
    with tab1:
        if role == "Ketua Sub Mawil":
            filter_wilayah = sub_mawil_aktif_terpilih 
            st.info(f"Menampilkan data khusus Sub Mawil: **{filter_wilayah}**")
        else:
            filter_wilayah = st.selectbox("Filter Berdasarkan Sub Mawil", ["Semua"] + DAFTAR_KAB_KOTA)
            
        if filter_wilayah != "Semua":
            df_tampil = get_data("SELECT nama as 'Nama', sub_mawil as 'Sub Mawil', jenis_kelamin as 'Jenis Kelamin', alamat as 'Alamat', status as 'Status', letnan_ijazah as 'Letnan Ijazah', tanggal_ijazah as 'Tanggal Ijazah', kontak as 'Kontak' FROM anggota WHERE sub_mawil = ?", (filter_wilayah,))
        else:
            df_tampil = get_data("SELECT nama as 'Nama', sub_mawil as 'Sub Mawil', jenis_kelamin as 'Jenis Kelamin', alamat as 'Alamat', status as 'Status', letnan_ijazah as 'Letnan Ijazah', tanggal_ijazah as 'Tanggal Ijazah', kontak as 'Kontak' FROM anggota")
        
        if df_tampil.empty:
            st.info("Belum ada data SanFK yang terdaftar di database.")
        else:
            df_tampil.index = range(1, len(df_tampil) + 1)
            st.dataframe(df_tampil, use_container_width=True)
        
    # --- TAB 2: HANYA TAMPIL UNTUK SEKRETARIS MAWIL ---
    if role == "Sekretaris Mawil":
        with tab2:
            st.subheader("Form Input, Edit, & Hapus Data SanFK")
            
            mode_aksi = st.radio("Pilih Mode Aksi:", ["Tambah SanFK Baru", "Edit / Hapus SanFK yang Ada"], horizontal=True)
            
            if mode_aksi == "Tambah SanFK Baru":
                nama = st.text_input("Nama Lengkap", key="t_nama")
                
                if role == "Ketua Sub Mawil":
                    sub_mawil = sub_mawil_aktif_terpilih
                    st.text_input("Sub Mawil (Kabupaten/Kota)", value=sub_mawil, disabled=True, key="t_sub_locked")
                else:
                    sub_mawil = st.selectbox("Sub Mawil (Kabupaten/Kota)", DAFTAR_KAB_KOTA, key="t_sub")
                    
                jenis_kelamin = st.selectbox("Jenis Kelamin", ["Laki-laki", "Perempuan"], key="t_jk")
                alamat = st.text_area("Alamat Lengkap", key="t_alamat")
                status = st.selectbox("Status Keaktifan", ["Aktif", "Tidak Aktif", "Pindah", "Meninggal Dunia"], key="t_status")
                letnan = st.text_input("Nama Letnan / Mursyid Pemberi Ijazah Dzikir", key="t_letnan")
                tgl_ijazah = st.date_input("Tanggal Perolehan Ijazah Dzikir", value=datetime.today(), key="t_tgl")
                kontak = st.text_input("Nomor Kontak / WhatsApp", key="t_kontak")
                
                st.markdown("---")
                st.markdown("##### 📷 Pas Foto / Dokumen SanFK *(Hanya Format Gambar/PDF, Tanpa Video)*")
                
                # Pilihan 2 Metode Input File
                metode_foto_sanfk = st.radio("Pilih Cara Input File:", ["Unggah Berkas (PDF, Foto, dll)", "Gunakan Kamera Langsung"], horizontal=True, key="radio_sanfk")
                
                foto_file = None
                cam_file = None
                
                if metode_foto_sanfk == "Unggah Berkas (PDF, Foto, dll)":
                    foto_file = st.file_uploader("Unggah File Pas Foto / Dokumen SanFK (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"], key="up_sanfk")
                else:
                    cam_file = st.camera_input("Ambil Pas Foto Langsung dengan Kamera", key="cam_sanfk")
                
                st.markdown("🔒 Hak Akses File Berkas: **Private (Otomatis)**")
                akses_sanfk = "Private"
                
                if st.button("Simpan SanFK Baru", key="btn_simpan_sanfk_baru"):
                    if nama:
                        final_foto_val = ""
                        
                        # Upload ke Cloudinary jika menggunakan Unggah Berkas
                        if metode_foto_sanfk == "Unggah Berkas (PDF, Foto, dll)" and foto_file is not None:
                            with st.spinner("Mengunggah berkas ke Cloudinary..."):
                                try:
                                    upload_res = cloudinary.uploader.upload(foto_file, resource_type="auto")
                                    foto_url = upload_res.get("secure_url")
                                    final_foto_val = f"{foto_url}|{akses_sanfk}"
                                except Exception as e:
                                    st.error(f"Gagal mengunggah file: {e}")
                            
                        # Upload ke Cloudinary jika menggunakan Kamera Langsung
                        elif metode_foto_sanfk == "Gunakan Kamera Langsung" and cam_file is not None:
                            with st.spinner("Mengunggah foto kamera ke Cloudinary..."):
                                try:
                                    upload_cam = cloudinary.uploader.upload(cam_file, resource_type="image")
                                    cam_url = upload_cam.get("secure_url")
                                    final_foto_val = f"{cam_url}|{akses_sanfk}"
                                except Exception as e:
                                    st.error(f"Gagal mengunggah foto kamera: {e}")
                        
                        execute_query(
                            "INSERT INTO anggota (nama, sub_mawil, jenis_kelamin, alamat, status, letnan_ijazah, tanggal_ijazah, kontak, foto) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                            (nama, sub_mawil, jenis_kelamin, alamat, status, letnan, str(tgl_ijazah), kontak, final_foto_val)
                        )
                        
                        st.success(f"Data SanFK {nama} berhasil ditambahkan!")
                        st.rerun()
                    else:
                        st.warning("Nama lengkap wajib diisi!")
            else:
                df_list_agt = get_data("SELECT id, nama, sub_mawil FROM anggota")
                if df_list_agt.empty:
                    st.info("Belum ada data SanFK untuk diedit atau dihapus.")
                else:
                    pilihan_agt = {f"{row['nama']} ({row['sub_mawil']})": row['id'] for _, row in df_list_agt.iterrows()}
                    pilih_label = st.selectbox("Pilih SanFK yang Ingin Diedit / Dihapus:", list(pilihan_agt.keys()))
                    id_terpilih = pilihan_agt[pilih_label]
                    
                    data_terpilih = get_data("SELECT * FROM anggota WHERE id = ?", (id_terpilih,)).iloc[0]
                    
                    file_foto_lama_full = data_terpilih.get('foto', '')
                    if "|" in file_foto_lama_full:
                        file_foto_lama, akses_lama_sanfk = file_foto_lama_full.split("|", 1)
                    else:
                        file_foto_lama, akses_lama_sanfk = file_foto_lama_full, "Public"

                    nama_e = st.text_input("Nama Lengkap", value=data_terpilih['nama'], key=f"e_nama_{id_terpilih}")
                    
                    if role == "Ketua Sub Mawil":
                        sub_mawil_e = sub_mawil_aktif_terpilih
                        st.text_input("Sub Mawil (Kabupaten/Kota)", value=sub_mawil_e, disabled=True, key=f"e_sub_locked_{id_terpilih}")
                    else:
                        sub_mawil_e = st.selectbox("Sub Mawil (Kabupaten/Kota)", DAFTAR_KAB_KOTA, index=DAFTAR_KAB_KOTA.index(data_terpilih['sub_mawil']) if data_terpilih['sub_mawil'] in DAFTAR_KAB_KOTA else 0, key=f"e_sub_{id_terpilih}")
                    
                    jk_list = ["Laki-laki", "Perempuan"]
                    jk_idx = jk_list.index(data_terpilih['jenis_kelamin']) if data_terpilih['jenis_kelamin'] in jk_list else 0
                    jenis_kelamin_e = st.selectbox("Jenis Kelamin", jk_list, index=jk_idx, key=f"e_jk_{id_terpilih}")
                    
                    alamat_e = st.text_area("Alamat Lengkap", value=data_terpilih['alamat'] if data_terpilih['alamat'] else "", key=f"e_al_{id_terpilih}")
                    
                    status_list = ["Aktif", "Tidak Aktif", "Pindah", "Meninggal Dunia"]
                    status_idx = status_list.index(data_terpilih['status']) if data_terpilih['status'] in status_list else 0
                    status_e = st.selectbox("Status Keaktifan", status_list, index=status_idx, key=f"e_st_{id_terpilih}")
                    
                    letnan_e = st.text_input("Nama Letnan / Mursyid Pemberi Ijazah Dzikir", value=data_terpilih['letnan_ijazah'] if data_terpilih['letnan_ijazah'] else "", key=f"e_let_{id_terpilih}")
                    
                    try:
                        tgl_parsed = datetime.strptime(data_terpilih['tanggal_ijazah'], "%Y-%m-%d").date()
                    except:
                        tgl_parsed = datetime.today().date()
                    tgl_ijazah_e = st.date_input("Tanggal Perolehan Ijazah Dzikir", value=tgl_parsed, key=f"e_tgl_{id_terpilih}")
                    
                    kontak_e = st.text_input("Nomor Kontak / WhatsApp", value=data_terpilih['kontak'] if data_terpilih['kontak'] else "", key=f"e_kon_{id_terpilih}")
                    
                    st.markdown("🔒 Hak Akses Berkas: **Private (Otomatis)**")
                    edit_akses_sanfk = "Private"
                    
                    st.markdown("---")
                    st.markdown("##### 📷 Ganti File / Pas Foto *(Tanpa Video)*")
                    metode_ganti_sanfk = st.radio("Pilih Cara Ganti File:", ["Unggah Berkas (PDF, Foto, dll)", "Gunakan Kamera Langsung"], horizontal=True, key=f"radio_ganti_{id_terpilih}")
                    
                    foto_file_e = None
                    cam_file_e = None
                    
                    if metode_ganti_sanfk == "Unggah Berkas (PDF, Foto, dll)":
                        foto_file_e = st.file_uploader("Ganti / Upload File Baru (PDF, JPG, PNG)", type=["pdf", "jpg", "jpeg", "png"], key=f"up_e_{id_terpilih}")
                    else:
                        cam_file_e = st.camera_input("Ambil Pas Foto Baru via Kamera", key=f"cam_e_{id_terpilih}")
                    
                    col_eb1, col_eb2 = st.columns(2)
                    with col_eb1:
                        if st.button("💾 Simpan Perubahan Data", key=f"btn_up_data_{id_terpilih}"):
                            foto_path_e = file_foto_lama
                            
                            if metode_ganti_sanfk == "Unggah Berkas (PDF, Foto, dll)" and foto_file_e is not None:
                                with st.spinner("Mengunggah file baru ke Cloudinary..."):
                                    try:
                                        upload_efe = cloudinary.uploader.upload(foto_file_e, resource_type="auto")
                                        foto_path_e = upload_efe.get("secure_url")
                                    except Exception as e:
                                        st.error(f"Gagal upload file: {e}")
                                final_foto_e_val = f"{foto_path_e}|{edit_akses_sanfk}"
                            elif metode_ganti_sanfk == "Gunakan Kamera Langsung" and cam_file_e is not None:
                                with st.spinner("Mengunggah foto kamera baru..."):
                                    try:
                                        upload_came = cloudinary.uploader.upload(cam_file_e, resource_type="image")
                                        foto_path_e = upload_came.get("secure_url")
                                    except Exception as e:
                                        st.error(f"Gagal upload kamera: {e}")
                                final_foto_e_val = f"{foto_path_e}|{edit_akses_sanfk}"
                            else:
                                final_foto_e_val = f"{foto_path_e}|{edit_akses_sanfk}"
                            
                            execute_query(
                                "UPDATE anggota SET nama = ?, sub_mawil = ?, jenis_kelamin = ?, alamat = ?, status = ?, letnan_ijazah = ?, tanggal_ijazah = ?, kontak = ?, foto = ? WHERE id = ?",
                                (nama_e, sub_mawil_e, jenis_kelamin_e, alamat_e, status_e, letnan_e, str(tgl_ijazah_e), kontak_e, final_foto_e_val, id_terpilih)
                            )
                            
                            st.success(f"Data SanFK {nama_e} berhasil diperbarui di Cloudinary!")
                            st.rerun()
                    with col_eb2:
                        if st.button("🗑️ Hapus SanFK Ini", key=f"btn_del_agt_{id_terpilih}"):
                            execute_query("DELETE FROM anggota WHERE id = ?", (id_terpilih,))
                            st.success(f"Data SanFK berhasil dihapus dari database!")
                            st.rerun()

    with tab3:
        st.subheader("💳 Kartu Tanda SanFK (KTA) Digital")
        df_anggota_kta = get_data("SELECT * FROM anggota")
        if df_anggota_kta.empty:
            st.info("Belum ada data SanFK.")
        else:
            pilih_anggota = st.selectbox("Pilih SanFK untuk Cetak KTA", df_anggota_kta["nama"].tolist())
            data_a = df_anggota_kta[df_anggota_kta["nama"] == pilih_anggota].iloc[0]
            
            foto_full = data_a.get('foto', '')
            if "|" in foto_full:
                foto_path, akses_kta = foto_full.split("|", 1)
            else:
                foto_path, akses_kta = foto_full, "Public"
                
            # Jika berupa URL Cloudinary, langsung render tag img menggunakan URL tersebut
            if foto_path and (foto_path.startswith("http://") or foto_path.startswith("https://")):
                img_html = f'<img src="{foto_path}" style="width: 85px; height: 105px; object-fit: cover; border-radius: 4px; border: 1px solid #0E6655;">'
            else:
                base64_img = get_image_base64(foto_path) if foto_path else None
                if base64_img:
                    img_html = f'<img src="{base64_img}" style="width: 85px; height: 105px; object-fit: cover; border-radius: 4px; border: 1px solid #0E6655;">'
                else:
                    img_html = '<div style="font-size: 10px; color: #555; padding: 25px 0; text-align: center;">Private / No File</div>'
            
            st.markdown(f"""
            <div style="border: 2px solid #0E6655; border-radius: 10px; padding: 20px; background-color: #E8F8F5; color: #0e3d30; max-width: 500px;">
                <h3 style="margin: 0; text-align: center;">FK MAWIL RIAU</h3>
                <p style="text-align: center; font-size: 12px; margin-bottom: 15px;">Forum Silaturrahmi Majelis Dzikir Fatwa Kehidupan Mawil Riau</p>
                <hr style="border-color: #0E6655;">
                <table style="width: 100%; color: #0e3d30; border: none;">
                    <tr>
                        <td style="width: 35%; vertical-align: top; text-align: center; padding-right: 10px;">
                            <div style="width: 90px; height: 110px; border: 1px dashed #0E6655; display: flex; align-items: center; justify-content: center; background: white; margin: auto; padding: 2px;">
                                {img_html}
                            </div>
                        </td>
                        <td style="width: 65%; vertical-align: top;">
                            <p style="margin: 4px 0;"><b>Nama:</b> {data_a['nama']}</p>
                            <p style="margin: 4px 0;"><b>Sub Mawil:</b> {data_a['sub_mawil']}</p>
                            <p style="margin: 4px 0;"><b>Gender:</b> {data_a.get('jenis_kelamin', '-')}</p>
                            <p style="margin: 4px 0;"><b>Status:</b> {data_a['status']}</p>
                            <p style="margin: 4px 0;"><b>No. HP:</b> {data_a.get('kontak', '-')}</p>
                            <p style="margin: 4px 0;"><b>Ijazah Dzikir:</b><br>{data_a['letnan_ijazah']}</p>
                        </td>
                    </tr>
                </table>
                <p style="margin: 8px 0 0 0; font-size: 11px;"><b>Alamat:</b> {data_a.get('alamat', '-')}</p>
                <div style="text-align: center; margin-top: 15px; background: white; padding: 8px; border-radius: 5px;">
                    <p style="font-size: 10px; margin: 0;">[ QR CODE PRESENSI DIGITAL ]</p>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
# --- 5. KEUANGAN & KOTAK HIJAU ---
if menu == "Keuangan & Kotak Hijau":
    # Batasi akses hanya untuk SanFK dan Bendahara Mawil
    if role not in ["SanFK", "Bendahara Mawil"]:
        st.error("⛔ Akses Ditolak!")
        st.warning("Menu 'Keuangan & Kotak Hijau' khusus diperuntukkan bagi role **SanFK** dan **Bendahara Mawil**. Ketua Mawil, Sekretaris Mawil, Admin Mawil, dan Ketua Sub Mawil tidak memiliki akses ke menu ini.")
    else:
        st.title("💰 Keuangan Terpusat, Rekening Bersama, & Bukti Transfer Mandiri")
        
        # --- KETENTUAN ALUR YANG SUDAH DIRAPIKAN ---
        st.info("""
        **Ketentuan Alur Keuangan & Kotak Hijau:**

        1. **Iuran Kas SanFK & Wakaf Produktif:** Disetor oleh SanFK ke Bendahara Mawil dan Dana dikelola langsung oleh Bendahara Mawil.
        2. **Kotak Hijau:** Penyaluran saja (disetor dari pengelola kotak hijau ke Bendahara Mawil, lalu Bendahara Mawil langsung menyetor kepada bendahara pusat).
        3. **Dana dari Pusat (Baksos/Santunan):** Bendahara Mawil menerima dari pusat dan langsung menyalurkan kepada yang berhak.
        4. **Baksos Lokal:** Disetor langsung oleh SanFK ke Rekening Padepokan Fatwa Kehidupan dengan kode 3 angka nomor keanggotaan (pada jumlah setoran).
        5. **Infaq Palestina:** Disetor langsung oleh SanFK ke Rekening Padepokan Fatwa Kehidupan dengan kode unik `888` (pada jumlah setoran).
        6. **Infaq Jabung:** Disetor langsung ke Rekening Pengurus Padepokan Fatwa Kehidupan di Jabung.
        """)

        # --- PENGATURAN TAB BERDASARKAN ROLE ---
        if role == "Bendahara Mawil":
            tab_f1, tab_f3, tab_f4, tab_f5, tab_f6 = st.tabs([
                "💳 Info Rekening & Unggah",        # Index 0 (tab_f1)
                "📁 Arsip & Koreksi Bukti",          # Index 1 (tab_f3)
                "📊 Cashflow & Kategori",            # Index 2 (tab_f4)
                "🏦 Pendataan Rekening",             # Index 3 (tab_f5)
                "🛠️ Otoritas Transaksi"              # Index 4 (tab_f6)
            ])
            tab_f2 = None
        elif role == "SanFK":
            tab_f1, tab_f2, tab_f3, tab_f4 = st.tabs([
                "💳 Info Rekening & Unggah",        # Index 0 (tab_f1)
                "⏳ Menunggu Validasi",              # Index 1 (tab_f2)
                "📁 Arsip & Koreksi Bukti",          # Index 2 (tab_f3)
                "📊 Cashflow & Kategori"             # Index 3 (tab_f4)
            ])
            tab_f5 = None
            tab_f6 = None
        else:
            tab_f1, tab_f3, tab_f4 = st.tabs([
                "💳 Info Rekening & Unggah", 
                "📁 Arsip & Koreksi Bukti",
                "📊 Cashflow & Kategori"
            ])
            tab_f2 = None
            tab_f5 = None
            tab_f6 = None

        # --- TAB 1: INFO REKENING & UNGGAH ---
        with tab_f1:
            st.subheader("💳 Informasi Nomor Rekening Tujuan Transfer")
            df_rek = get_data("SELECT * FROM rekening_tujuan")
            if df_rek.empty:
                st.info("Belum ada nomor rekening tujuan yang didata oleh Bendahara Mawil.")
            else:
                for _, r_rek in df_rek.iterrows():
                    st.markdown(f"""
                    <div style="background: #E8F8F5; padding: 12px; border-radius: 8px; border: 1px solid #0E6655; margin-bottom: 10px;">
                        <h4 style="margin: 0; color: #0E6655;">🏦 {r_rek['nama_bank']}</h4>
                        <p style="margin: 4px 0; font-size: 16px;"><b>No. Rekening:</b> <code>{r_rek['nomor_rekening']}</code></p>
                        <p style="margin: 2px 0;"><b>Atas Nama:</b> {r_rek['atas_nama']}</p>
                        <p style="margin: 2px 0; font-size: 13px; color: #555;"><i>{r_rek['keterangan']}</i></p>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("---")
            st.subheader("📤 Formulir Unggah Bukti Transfer & Komentar")
            
            if "uploader_counter" not in st.session_state:
                st.session_state["uploader_counter"] = 0

            if "sukses_kirim_notif" in st.session_state:
                st.success(st.session_state["sukses_kirim_notif"])
                del st.session_state["sukses_kirim_notif"]

            if role not in ["SanFK", "Bendahara Mawil"]:
                st.warning("⚠️ Anda harus login sebagai SanFK atau Bendahara Mawil untuk mengunggah dokumen.")
            else:
                nama_pengirim_aktif = sanfk_aktif_terpilih if role == "SanFK" else "Bendahara Mawil"
                if role == "SanFK":
                    if not sanfk_aktif_terpilih:
                        st.warning("⚠️ Silakan pilih profil SanFK Anda di sidebar terlebih dahulu.")
                    else:
                        st.info(f"👤 Pengirim (SanFK): **{sanfk_aktif_terpilih}**")
                        d_sanfk_info = get_data("SELECT sub_mawil FROM anggota WHERE nama = ?", (sanfk_aktif_terpilih,))
                        sub_mw_asal = d_sanfk_info.iloc[0]['sub_mawil'] if not d_sanfk_info.empty else "Pekanbaru"
                else:
                    nama_pengirim_aktif = st.text_input("Nama Petugas / Pengirim:", value="Bendahara Mawil", key="input_nama_bendahara_f1")
                    sub_mw_asal = st.selectbox("Asal Sub Mawil / Posko:", DAFTAR_KAB_KOTA + ["Pusat"], key="select_submw_bendahara_f1")

                if role != "SanFK" or sanfk_aktif_terpilih:
                    st.markdown("📝 *Pilih metode lampiran bukti transfer:*")
                    metode_unggah = st.radio("Metode Unggah:", ["Unggah File (JPG, PNG, PDF)", "Gunakan Kamera Langsung"], horizontal=True, key="radio_metode_sanfk_f1_live")
                    
                    up_bukti_file = None
                    cam_bukti = None

                    if metode_unggah == "Unggah File (JPG, PNG, PDF)":
                        up_bukti_file = st.file_uploader("Pilih File Bukti Transfer", type=["jpg", "jpeg", "png", "pdf"], key=f"up_file_sanfk_f1_single_{st.session_state['uploader_counter']}")
                    else:
                        cam_bukti = st.camera_input("Potret Bukti Transfer dengan Kamera", key=f"cam_input_sanfk_f1_{st.session_state['uploader_counter']}")

                    with st.form("form_unggah_bukti_mandiri", clear_on_submit=True):
                        if role == "SanFK":
                            kategori_cf = "Menunggu Validasi Bendahara"
                            st.markdown("🏷️ Kategori Setoran: **Menunggu Validasi Bendahara** *(Otomatis)*")
                            
                            jumlah_tf = st.number_input("Nominal Transfer (Rp)", min_value=0.0, step=10000.0, key="num_nominal_sanfk_f1")
                            jenis_arus_tf = "Masuk (Setoran)"
                        else:
                            kategori_cf = st.selectbox(
                                "Pilih Kategori Transaksi:", 
                                [
                                    "Iuran Kas SanFK", 
                                    "Wakaf Produktif", 
                                    "Kotak Hijau", 
                                    "Dana dari Pusat (Baksos/Santunan)",
                                    "Lain-lain"
                                ],
                                key="select_kategori_bendahara_f1_live"
                            )
                            jumlah_tf = st.number_input("Nominal Transaksi (Rp)", min_value=0.0, step=10000.0, key="num_nominal_bendahara_f1_live")
                            jenis_arus_tf = st.selectbox("Jenis Arus Dana:", ["Masuk (Setoran)", "Keluar / Penyaluran"], key="select_arus_bendahara_f1_live")

                        ket_tf = st.text_area("Komentar / Catatan Transfer:", key="textarea_ket_tf_f1_live")
                        btn_kirim_dok = st.form_submit_button("Kirim Bukti Transfer & Komentar")
                        
                        if btn_kirim_dok:
                            path_bukti = ""
                            if up_bukti_file is not None:
                                with st.spinner("Mengunggah bukti transfer ke Cloudinary..."):
                                    try:
                                        upload_tf = cloudinary.uploader.upload(up_bukti_file, resource_type="auto")
                                        path_bukti = upload_tf.get("secure_url")
                                    except Exception as e:
                                        st.error(f"Gagal mengunggah file: {e}")
                            elif cam_bukti is not None:
                                with st.spinner("Mengunggah foto kamera ke Cloudinary..."):
                                    try:
                                        upload_cam_tf = cloudinary.uploader.upload(cam_bukti, resource_type="image")
                                        path_bukti = upload_cam_tf.get("secure_url")
                                    except Exception as e:
                                        st.error(f"Gagal mengunggah foto kamera: {e}")

                            if path_bukti or ket_tf:
                                final_bukti_str = f"{path_bukti}|Private" if path_bukti else ""
                                execute_query(
                                    "INSERT INTO cashflow_transaksi (kategori, pengirim, sub_mawil, tanggal, jumlah, jenis_arus, keterangan, bukti_transfer) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                    (kategori_cf, nama_pengirim_aktif, sub_mw_asal, str(date.today()), jumlah_tf, jenis_arus_tf, ket_tf, final_bukti_str)
                                )
                                
                                st.session_state["uploader_counter"] += 1
                                st.session_state["sukses_kirim_notif"] = f"✅ Berhasil! Bukti transfer dan komentar berhasil dikirim ke Cloudinary untuk kategori **{kategori_cf}** (Nominal: **Rp {jumlah_tf:,.0f}**). Form dan uploader telah dibersihkan untuk sesi baru."
                                st.rerun()
                            else:
                                st.warning("⚠️ Harap lampirkan bukti transfer atau isi komentar terlebih dahulu!")

        # --- TAB 2: DAFTAR MENUNGGU VALIDASI (Hanya untuk SanFK) ---
        if role == "SanFK" and tab_f2 is not None:
            with tab_f2:
                st.subheader("⏳ Status Setoran Anda yang Menunggu Validasi")
                if not sanfk_aktif_terpilih:
                    st.warning("⚠️ Silakan pilih profil SanFK Anda di sidebar terlebih dahulu.")
                    df_menunggu = pd.DataFrame()
                else:
                    st.info(f"💡 Tab ini menampilkan status setoran atas nama **{sanfk_aktif_terpilih}** yang sedang menunggu validasi.")
                    df_menunggu = get_data("SELECT * FROM cashflow_transaksi WHERE pengirim = ? AND kategori = 'Menunggu Validasi Bendahara' ORDER BY id DESC", (sanfk_aktif_terpilih,))
                
                if df_menunggu.empty:
                    st.success("🎉 Tidak ada setoran yang sedang menunggu validasi saat ini.")
                else:
                    df_tampil_pending = df_menunggu[['id', 'tanggal', 'pengirim', 'sub_mawil', 'jumlah', 'keterangan']].copy()
                    df_tampil_pending.columns = ['ID', 'Tanggal', 'Pengirim (SanFK)', 'Sub Mawil', 'Jumlah (Rp)', 'Komentar/Catatan']
                    df_tampil_pending.index = range(1, len(df_tampil_pending) + 1)
                    st.dataframe(df_tampil_pending, use_container_width=True)

        # --- TAB 3: ARSIP & KOREKSI BUKTI ---
        with tab_f3:
            st.subheader("📁 Arsip & Koreksi Bukti Transfer")
            st.info("🔒 Tab ini menampilkan arsip bukti transfer. Bendahara Mawil dapat melihat dan membuka bukti yang dikirimkan SanFK, sedangkan penghapusan arsip hanya dapat dilakukan oleh pengirim (SanFK) atau pembuat data.")

            if role not in ["SanFK", "Bendahara Mawil"]:
                st.warning("⚠️ Akses dibatasi.")
            else:
                if role == "SanFK":
                    if not sanfk_aktif_terpilih:
                        st.warning("Silakan pilih profil SanFK di sidebar.")
                        df_arsip = pd.DataFrame()
                    else:
                        df_arsip = get_data("SELECT * FROM cashflow_transaksi WHERE pengirim = ? ORDER BY id DESC", (sanfk_aktif_terpilih,))
                else:
                    df_arsip = get_data("SELECT * FROM cashflow_transaksi ORDER BY id DESC")

                if df_arsip.empty:
                    st.info("Belum ada arsip bukti transfer atau komentar yang tersimpan.")
                else:
                    for _, r_arsip in df_arsip.iterrows():
                        with st.container():
                            st.markdown(f"""
                            <div style="background: #F4F6F6; padding: 14px; border-radius: 8px; border: 1px solid #BDC3C7; margin-bottom: 12px;">
                                <p style="margin: 0; font-size: 14px; color: #7F8C8D;">📅 Tanggal: {r_arsip['tanggal']} | 👤 Pengirim: <b>{r_arsip['pengirim']}</b> ({r_arsip['sub_mawil']}) ID Transaksi: [{r_arsip['id']}]</p>
                                <p style="margin: 4px 0;">🏷️ Kategori: <b>{r_arsip['kategori']}</b> | 💰 Nominal: <b>Rp {r_arsip['jumlah']:,.0f}</b></p>
                                <p style="margin: 4px 0; background: #fff; padding: 8px; border-radius: 4px;">💬 <b>Komentar/Catatan:</b> {r_arsip['keterangan'] if r_arsip['keterangan'] else '-'}</p>
                            </div>
                            """, unsafe_allow_html=True)

                            b_str = r_arsip.get('bukti_transfer', '')
                            b_path = b_str.split("|")[0].strip() if "|" in b_str else b_str.strip()
                            
                            if b_path:
                                ext_file = b_path.split('.')[-1].lower().split('?')[0]
                                if ext_file in ['jpg', 'jpeg', 'png', 'webp']:
                                    st.markdown(f'<img src="{b_path}" style="max-width: 250px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 8px;">', unsafe_allow_html=True)
                                else:
                                    st.info("📄 Berkas / Dokumen Cloud Terlampir")
                                
                                st.markdown(f"[📥 Download / Buka Bukti (ID: {r_arsip['id']})]({b_path})", unsafe_allow_html=True)
                            
                            boleh_hapus = True
                            if role == "Bendahara Mawil" and r_arsip['pengirim'] != "Bendahara Mawil":
                                boleh_hapus = False

                            if boleh_hapus:
                                col_del1, col_del2 = st.columns([1, 4])
                                with col_del1:
                                    if st.button("🗑️ Hapus Data Ini", key=f"btn_hapus_arsip_{r_arsip['id']}", type="primary"):
                                        execute_query("DELETE FROM cashflow_transaksi WHERE id = ?", (r_arsip['id'],))
                                        st.success(f"Data transaksi ID [{r_arsip['id']}] berhasil dihapus!")
                                        st.rerun()
                                with col_del2:
                                    st.caption("Jika foto/dokumen salah, klik tombol hapus di samping, lalu unggah kembali melalui Tab 1.")
                            else:
                                st.info("🔒 Arsip dari SanFK ini hanya dapat dilihat oleh Bendahara Mawil (penghapusan arsip wewenang SanFK pengirim).")

                            st.divider()

        # --- TAB 4: LAPORAN CASHFLOW & KATEGORI ---
        with tab_f4:
            st.subheader("📊 Laporan Cashflow Terstruktur")
            
            df_cf_all = get_data("SELECT * FROM cashflow_transaksi WHERE kategori != 'Menunggu Validasi Bendahara' ORDER BY tanggal ASC, id ASC")
            
            if df_cf_all.empty:
                st.info("Belum ada data cashflow sah tercatat.")
            else:
                df_cf_all['tanggal_dt'] = pd.to_datetime(df_cf_all['tanggal'])
                min_date = df_cf_all['tanggal_dt'].min().date()
                max_date = df_cf_all['tanggal_dt'].max().date()

                st.markdown("📅 **Filter Rentang Tanggal Laporan:**")
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    tgl_mulai = st.date_input("Dari Tanggal:", value=min_date, key="filter_tgl_mulai_cf")
                with col_d2:
                    tgl_selesai = st.date_input("Sampai Tanggal:", value=max_date, key="filter_tgl_selesai_cf")

                cat_filter = st.selectbox(
                    "Filter Kategori:", 
                    [
                        "Semua Kategori", 
                        "Iuran Kas SanFK", 
                        "Wakaf Produktif", 
                        "Kotak Hijau", 
                        "Dana dari Pusat (Baksos/Santunan)",
                        "Lain-lain"
                    ],
                    key="filter_kategori_cf_tab_laporan"
                )

                mask_sebelum = df_cf_all['tanggal_dt'].dt.date < tgl_mulai
                df_sebelum = df_cf_all[mask_sebelum]
                if cat_filter != "Semua Kategori":
                    df_sebelum = df_sebelum[df_sebelum['kategori'] == cat_filter]
                
                masuk_sebelum = df_sebelum[df_sebelum['jenis_arus'] == 'Masuk (Setoran)']['jumlah'].sum()
                keluar_sebelum = df_sebelum[df_sebelum['jenis_arus'] == 'Keluar / Penyaluran']['jumlah'].sum()
                saldo_awal = masuk_sebelum - keluar_sebelum

                mask_rentang = (df_cf_all['tanggal_dt'].dt.date >= tgl_mulai) & (df_cf_all['tanggal_dt'].dt.date <= tgl_selesai)
                df_cf_f = df_cf_all[mask_rentang]
                if cat_filter != "Semua Kategori":
                    df_cf_f = df_cf_f[df_cf_f['kategori'] == cat_filter]

                total_masuk = df_cf_f[df_cf_f['jenis_arus'] == 'Masuk (Setoran)']['jumlah'].sum()
                total_keluar = df_cf_f[df_cf_f['jenis_arus'] == 'Keluar / Penyaluran']['jumlah'].sum()
                saldo_akhir = saldo_awal + total_masuk - total_keluar

                st.markdown("---")
                
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                with col_m1:
                    st.metric("Saldo Awal", f"Rp {saldo_awal:,.0f}")
                with col_m2:
                    st.metric("Total Masuk", f"Rp {total_masuk:,.0f}")
                with col_m3:
                    st.metric("Total Keluar", f"Rp {total_keluar:,.0f}")
                with col_m4:
                    st.metric("Saldo Akhir", f"Rp {saldo_akhir:,.0f}")

                st.markdown("---")
                
                # Tampilkan Tabel Data
                df_t_show = df_cf_f[['id', 'tanggal', 'kategori', 'pengirim', 'sub_mawil', 'jumlah', 'jenis_arus', 'keterangan']].copy()
                df_t_show.columns = ['ID', 'Tanggal', 'Kategori', 'Pengirim', 'Sub Mawil', 'Jumlah (Rp)', 'Arus', 'Keterangan']
                st.dataframe(df_t_show, use_container_width=True, hide_index=True)

                # --- FITUR EDIT & HAPUS KHUSUS BENDAHARA MAWIL PADA TAB LAPORAN ---
                if role == "Bendahara Mawil":
                    st.markdown("---")
                    st.markdown("### 🛠️ Kelola Transaksi (Edit & Hapus oleh Bendahara Mawil)")
                    
                    if df_cf_f.empty:
                        st.info("Tidak ada transaksi pada filter ini untuk dikelola.")
                    else:
                        dict_transaksi_pilihan = {
                            f"ID [{r['id']}] - {r['tanggal']} | {r['kategori']} | Rp {r['jumlah']:,.0f} ({r['pengirim']})": r['id']
                            for _, r in df_cf_f.iterrows()
                        }
                        
                        pilih_trx_label = st.selectbox(
                            "Pilih Transaksi yang Ingin Diedit atau Dihapus:",
                            list(dict_transaksi_pilihan.keys()),
                            key="select_trx_kelola_bendahara"
                        )
                        id_trx_pilih = dict_transaksi_pilihan[pilih_trx_label]
                        
                        trx_detail_row = get_data("SELECT * FROM cashflow_transaksi WHERE id = ?", (id_trx_pilih,)).iloc[0]

                        with st.form(f"form_edit_hapus_trx_{id_trx_pilih}", clear_on_submit=True):
                            st.markdown(f"**Edit Data Transaksi [ID: {id_trx_pilih}]**")
                            
                            kategori_opsi = [
                                "Iuran Kas SanFK", 
                                "Wakaf Produktif", 
                                "Kotak Hijau", 
                                "Dana dari Pusat (Baksos/Santunan)",
                                "Lain-lain"
                            ]
                            kat_lama = trx_detail_row['kategori']
                            if kat_lama not in kategori_opsi:
                                kategori_opsi.insert(0, kat_lama)
                            idx_kat = kategori_opsi.index(kat_lama)

                            e_kat = st.selectbox("Kategori Transaksi:", kategori_opsi, index=idx_kat, key=f"edit_kat_{id_trx_pilih}")
                            e_jumlah = st.number_input("Nominal (Rp):", min_value=0.0, step=10000.0, value=float(trx_detail_row['jumlah']), key=f"edit_jumlah_{id_trx_pilih}")
                            
                            arus_opsi = ["Masuk (Setoran)", "Keluar / Penyaluran"]
                            idx_arus = arus_opsi.index(trx_detail_row['jenis_arus']) if trx_detail_row['jenis_arus'] in arus_opsi else 0
                            e_arus = st.selectbox("Jenis Arus:", arus_opsi, index=idx_arus, key=f"edit_arus_{id_trx_pilih}")
                            
                            e_ket = st.text_input("Keterangan / Catatan:", value=str(trx_detail_row['keterangan']), key=f"edit_ket_{id_trx_pilih}")

                            col_btn1, col_btn2 = st.columns(2)
                            with col_btn1:
                                btn_simpan_edit = st.form_submit_button("💾 Simpan Perubahan")
                            with col_btn2:
                                btn_hapus_trx = st.form_submit_button("🗑️ Hapus Transaksi Ini")

                            if btn_simpan_edit:
                                execute_query(
                                    "UPDATE cashflow_transaksi SET kategori = ?, jumlah = ?, jenis_arus = ?, keterangan = ? WHERE id = ?",
                                    (e_kat, e_jumlah, e_arus, e_ket, id_trx_pilih)
                                )
                                st.success(f"✅ Transaksi ID [{id_trx_pilih}] berhasil diperbarui!")
                                st.rerun()

                            if btn_hapus_trx:
                                execute_query("DELETE FROM cashflow_transaksi WHERE id = ?", (id_trx_pilih,))
                                st.success(f"🗑️ Transaksi ID [{id_trx_pilih}] berhasil dihapus dari tabel cashflow!")
                                st.rerun()

                # --- TOMBOL DOWNLOAD EXCEL & WORD ---
                st.markdown("---")
                st.markdown("📥 **Unduh Laporan Keuangan:**")
                col_dl1, col_dl2 = st.columns(2)

                # 1. GENERATE EXCEL (.xlsx)
                with col_dl1:
                    output_excel = io.BytesIO()
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Laporan Cashflow"
                    
                    ws.append(["LAPORAN CASHFLOW"])
                    ws.append(["KEUANGAN FK MAWIL RIAU"])
                    ws.append([f"Periode: {tgl_mulai} s.d. {tgl_selesai} | Kategori: {cat_filter}"])
                    ws.append([])
                    
                    ws.append(["Saldo Awal", saldo_awal])
                    ws.append(["Total Masuk", total_masuk])
                    ws.append(["Total Keluar", total_keluar])
                    ws.append(["Saldo Akhir", saldo_akhir])
                    ws.append([])
                    
                    ws.append(list(df_t_show.columns))
                    
                    for _, row in df_t_show.iterrows():
                        ws.append(list(row))
                        
                    wb.save(output_excel)
                    output_excel.seek(0)
                    
                    st.download_button(
                        label="📥 Download Laporan (Excel .xlsx)",
                        data=output_excel,
                        file_name=f"laporan_cashflow_{tgl_mulai}_sd_{tgl_selesai}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_dl_excel_cf"
                    )

                # 2. GENERATE WORD (.docx)
                with col_dl2:
                    output_word = io.BytesIO()
                    doc = Document()
                    
                    p_title1 = doc.add_paragraph()
                    run_title1 = p_title1.add_run("LAPORAN CASHFLOW")
                    run_title1.bold = True
                    run_title1.font.size = Pt(16)
                    run_title1.font.color.rgb = RGBColor(14, 102, 85)
                    p_title1.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    p_title2 = doc.add_paragraph()
                    run_title2 = p_title2.add_run("KEUANGAN FK MAWIL RIAU")
                    run_title2.bold = True
                    run_title2.font.size = Pt(14)
                    run_title2.font.color.rgb = RGBColor(14, 102, 85)
                    p_title2.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    p_sub = doc.add_paragraph(f"Periode: {tgl_mulai} s.d. {tgl_selesai} | Kategori: {cat_filter}")
                    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    doc.add_paragraph()

                    doc.add_heading("Ringkasan Saldo", level=2)
                    p_sum = doc.add_paragraph()
                    p_sum.add_run(f"• Saldo Awal : Rp {saldo_awal:,.0f}\n")
                    p_sum.add_run(f"• Total Masuk : Rp {total_masuk:,.0f}\n")
                    p_sum.add_run(f"• Total Keluar : Rp {total_keluar:,.0f}\n")
                    p_sum.add_run(f"• Saldo Akhir : Rp {saldo_akhir:,.0f}\n")
                    
                    doc.add_heading("Detail Transaksi", level=2)
                    
                    if not df_t_show.empty:
                        table = doc.add_table(rows=1, cols=len(df_t_show.columns))
                        table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        table.style = 'Table Grid'
                        
                        hdr_cells = table.rows[0].cells
                        for i, col_name in enumerate(df_t_show.columns):
                            hdr_cells[i].text = str(col_name)
                            for paragraph in hdr_cells[i].paragraphs:
                                for run in paragraph.runs:
                                    run.bold = True
                                    
                        for _, row in df_t_show.iterrows():
                            row_cells = table.add_row().cells
                            for i, val in enumerate(row):
                                row_cells[i].text = str(val)
                                
                    doc.save(output_word)
                    output_word.seek(0)
                    
                    st.download_button(
                        label="📥 Download Laporan (Word .docx)",
                        data=output_word,
                        file_name=f"laporan_cashflow_{tgl_mulai}_sd_{tgl_selesai}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_dl_word_cf"
                    )

        # --- TAB 5 & 6 KHUSUS BENDAHARA MAWIL ---
        if role == "Bendahara Mawil":
            with tab_f5:
                st.subheader("🏦 Tab Khusus Pendataan Nomor Rekening (Bendahara Mawil)")
                with st.form("form_tambah_rekening_baru", clear_on_submit=True):
                    st.markdown("##### ➕ Tambah Rekening Tujuan Baru")
                    n_bank = st.text_input("Nama Bank (Cth: BSI, BCA, Mandiri)", key="add_bank_f4")
                    n_rek = st.text_input("Nomor Rekening", key="add_norek_f4")
                    n_an = st.text_input("Atas Nama Pemilik Rekening", key="add_an_f4")
                    n_ket = st.text_input("Keterangan Rekening (Cth: Padepokan Fatwa Kehidupan / Jabung)", key="add_ket_f4")
                    
                    btn_simpan_rek = st.form_submit_button("Simpan Rekening Baru")
                    if btn_simpan_rek and n_rek:
                        execute_query("INSERT INTO rekening_tujuan (nama_bank, nomor_rekening, atas_nama, keterangan) VALUES (?, ?, ?, ?)", (n_bank, n_rek, n_an, n_ket))
                        st.success("Nomor rekening berhasil ditambahkan!")
                        st.rerun()

                st.markdown("---")
                st.markdown("##### 🛠️ Kelola, Edit, atau Hapus Nomor Rekening Terdaftar")
                
                df_rek_manage = get_data("SELECT * FROM rekening_tujuan")
                if df_rek_manage.empty:
                    st.info("Belum ada nomor rekening yang terdaftar di database.")
                else:
                    pilihan_rek_dict = {f"ID [{r['id']}]: {r['nama_bank']} - {r['nomor_rekening']} ({r['atas_nama']})": r['id'] for _, r in df_rek_manage.iterrows()}
                    pilih_label_rek = st.selectbox("Pilih Nomor Rekening untuk Dikelola:", list(pilihan_rek_dict.keys()), key="select_rek_manage_box_f4")
                    id_rek_aktif = pilihan_rek_dict[pilih_label_rek]

                    df_cek_aktif = get_data("SELECT * FROM rekening_tujuan WHERE id = ?", (id_rek_aktif,))
                    if not df_cek_aktif.empty:
                        data_rek_pilih = df_cek_aktif.iloc[0]

                        col_h1, col_h2 = st.columns([3, 1])
                        with col_h2:
                            if st.button("🗑️ Hapus Rekening Ini", key=f"btn_del_rek_outside_{id_rek_aktif}", type="primary"):
                                execute_query("DELETE FROM rekening_tujuan WHERE id = ?", (id_rek_aktif,))
                                st.success("Nomor rekening berhasil dihapus dari database!")
                                st.rerun()

                        with st.form(f"form_edit_rek_{id_rek_aktif}"):
                            e_bank = st.text_input("Ubah Nama Bank", value=data_rek_pilih['nama_bank'], key=f"ebank_{id_rek_aktif}")
                            e_norek = st.text_input("Ubah Nomor Rekening", value=data_rek_pilih['nomor_rekening'], key=f"enorek_{id_rek_aktif}")
                            e_an = st.text_input("Ubah Atas Nama", value=data_rek_pilih['atas_nama'], key=f"ean_{id_rek_aktif}")
                            e_ket = st.text_input("Ubah Keterangan", value=data_rek_pilih['keterangan'], key=f"eket_{id_rek_aktif}")

                            btn_s_erek = st.form_submit_button("💾 Simpan Perubahan Rekening")
                            if btn_s_erek:
                                execute_query("UPDATE rekening_tujuan SET nama_bank = ?, nomor_rekening = ?, atas_nama = ?, keterangan = ? WHERE id = ?", (e_bank, e_norek, e_an, e_ket, id_rek_aktif))
                                st.success("Data rekening berhasil diperbarui!")
                                st.rerun()

            with tab_f6:
                st.subheader("🛠️ Otoritas & Validasi Transaksi")
                st.info("💡 Menu ini menampilkan daftar transaksi yang memerlukan validasi atau koreksi dari Bendahara Mawil.")

                df_all_tf = get_data("SELECT * FROM cashflow_transaksi WHERE kategori = 'Menunggu Validasi Bendahara' ORDER BY id DESC")
                
                if df_all_tf.empty:
                    st.success("🎉 Semua transaksi sudah tervalidasi! Tidak ada data yang menunggu validasi saat ini.")
                else:
                    pilihan_all_dict = {
                        f"[{r['tanggal']}] {r['pengirim']} ({r['sub_mawil']}) - Rp {r['jumlah']:,.0f} [ID: {r['id']}]": r['id']
                        for _, r in df_all_tf.iterrows()
                    }
                    
                    pilih_label_all = st.selectbox(
                        "Pilih Transaksi untuk Divalidasi:",
                        list(pilihan_all_dict.keys()),
                        key="select_all_transaksi_tab6"
                    )
                    id_all_pilih = pilihan_all_dict[pilih_label_all]

                    data_all_detail = get_data(
                        "SELECT * FROM cashflow_transaksi WHERE id = ?", (id_all_pilih,)
                    ).iloc[0]

                    with st.form(f"form_otoritas_umum_{id_all_pilih}", clear_on_submit=True):
                        st.markdown("##### 📝 Validasi / Koreksi Transaksi")
                        
                        kategori_tersedia_otoritas = [
                            "Iuran Kas SanFK", 
                            "Wakaf Produktif", 
                            "Kotak Hijau", 
                            "Dana dari Pusat (Baksos/Santunan)",
                            "Lain-lain"
                        ]

                        e_kat_umum = st.selectbox(
                            "Pilih Kategori Sah Transaksi:", 
                            kategori_tersedia_otoritas, 
                            index=0,
                            key=f"ekat_umum_{id_all_pilih}"
                        )
                        
                        st.markdown(f"**Nominal Tetap:** Rp {data_all_detail['jumlah']:,.0f}")
                        e_arus_umum = st.selectbox("Jenis Arus Dana:", ["Masuk (Setoran)", "Keluar / Penyaluran"], index=0, key=f"earus_umum_{id_all_pilih}")
                        e_ket_umum = st.text_input("Keterangan", value=data_all_detail['keterangan'], key=f"eket_umum_{id_all_pilih}")
                        
                        col_o1, col_o2 = st.columns(2)
                        with col_o1:
                            btn_s_oum = st.form_submit_button("💾 Validasi")
                        with col_o2:
                            btn_h_oum = st.form_submit_button("🗑️ Hapus Transaksi")

                        if btn_s_oum:
                            execute_query(
                                "UPDATE cashflow_transaksi SET kategori = ?, jenis_arus = ?, keterangan = ? WHERE id = ?", 
                                (e_kat_umum, e_arus_umum, e_ket_umum, id_all_pilih)
                            )
                            st.success(f"✅ Transaksi ID [{id_all_pilih}] berhasil divalidasi ke kategori **{e_kat_umum}**! Form dibersihkan.")
                            st.rerun()
                            
                        if btn_h_oum:
                            execute_query("DELETE FROM cashflow_transaksi WHERE id = ?", (id_all_pilih,))
                            st.success(f"🗑️ Transaksi ID [{id_all_pilih}] berhasil dihapus! Form dibersihkan.")
                            st.rerun()
