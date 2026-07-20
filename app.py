import io
import json
import os
import sqlite3
import sys
import uuid
from datetime import datetime, date

from flask import Flask, render_template, request, redirect, url_for, flash, g, send_file, jsonify
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

DB_PATH = None
IMPORT_TMP_DIR = None

# ---------------- Konfigurasi Path untuk PyInstaller Exe ----------------
if getattr(sys, "frozen", False):
    # Jika berjalan sebagai .exe hasil build PyInstaller
    BASE_DIR = sys._MEIPASS
    # Letakkan database di folder yang sama dengan file .exe utama
    DB_PATH = os.path.join(os.path.dirname(sys.executable), "sidap.db")
else:
    # Jika berjalan normal via python app.py
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DB_PATH = os.path.join(BASE_DIR, "sidap.db")

IMPORT_TMP_DIR = os.path.join(BASE_DIR, "import_tmp")

# Beritahu Flask lokasi folder templates dan static yang berada di dalam folder ekstraksi .exe
app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)
app.secret_key = "sidap-dev-secret"


def is_valid_iso_date(s):
    """Cek strict format YYYY-MM-DD DAN kewajaran tahunnya. Jangan percaya
    <input type=date> di browser bakal selalu ngirim format yang masuk akal
    -- interaksi keyboard yang aneh di date picker bisa hasilin tahun kayak
    '4423' yang formatnya sah tapi jelas typo."""
    try:
        parsed = datetime.strptime(s, "%Y-%m-%d")
    except (ValueError, TypeError):
        return False
    tahun_sekarang = datetime.now().year
    return 1990 <= parsed.year <= tahun_sekarang + 2

# ---------------- Koneksi DB ----------------
def get_db():
    if "db" not in g:
        if not db_has_schema():
            init_db()
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    schema_path = os.path.join(BASE_DIR, "schema.sql")
    if not os.path.exists(schema_path):
        raise FileNotFoundError(f"schema.sql not found in {BASE_DIR}. Please restore schema.sql before running SIDAP.")

    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    with open(schema_path, "r", encoding="utf-8") as f:
        conn.executescript(f.read())
    conn.commit()
    conn.close()


def db_has_schema():
    if not os.path.exists(DB_PATH):
        return False
    try:
        conn = sqlite3.connect(DB_PATH)
        required_tables = ["pelanggan", "instalasi", "permohonan"]
        existing = {
            row[0] for row in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()
        }
        conn.close()
        return all(name in existing for name in required_tables)
    except sqlite3.DatabaseError:
        return False


def ensure_data_paths():
    os.makedirs(IMPORT_TMP_DIR, exist_ok=True)
    if not db_has_schema():
        init_db()


def daftar_kecamatan(db):
    rows = db.execute(
        "SELECT DISTINCT kecamatan FROM pelanggan ORDER BY kecamatan"
    ).fetchall()
    return [r["kecamatan"] for r in rows]


def daftar_pelanggan_pilihan(db):
    rows = db.execute("SELECT id, nama, kecamatan FROM pelanggan ORDER BY nama").fetchall()
    return [{"id": r["id"], "label": f"{r['nama']} \u2014 {r['kecamatan']}"} for r in rows]


# ---------------- Import Excel: template + parsing ----------------
HEADER_MAP = {
    "no": "no",
    "nama pelanggan": "nama",
    "alamat pelanggan": "alamat",
    "kelurahan": "kelurahan",
    "kecamatan": "kecamatan",
    "nomor instalasi": "nomor_instalasi",
    "tanggal pasang": "tanggal_pasang",
    "diameter pipa": "diameter_pipa",
    "tekanan air": "tekanan_air",
    "status (sib/bk)": "status",
    "petugas": "petugas",
    "keterangan": "keterangan",
}
WAJIB_DIISI = ["nama", "kecamatan", "nomor_instalasi", "tanggal_pasang", "status"]
VALID_STATUSES = ("SIB", "BK")

PERMOHONAN_HEADER_MAP = {
    "no": "no",
    "nama pelanggan": "nama_pelanggan",
    "lokasi": "lokasi",
    "kelurahan": "kelurahan",
    "kecamatan": "kecamatan",
    "jenis": "jenis",
    "tanggal permohonan": "tanggal_permohonan",
    "tanggal survey": "tanggal_survey",
    "petugas survey": "petugas_survey",
    "ditindaklanjuti": "ditindaklanjuti",
    "jenis pipa": "jenis_pipa",
    "tanggal dikirim hublang": "tanggal_dikirim_hublang",
    "tanggal kembali hublang": "tanggal_kembali_hublang",
    "keterangan": "keterangan",
}
PERMOHONAN_REQUIRED = ["nama_pelanggan", "kecamatan", "jenis", "tanggal_permohonan"]
PERMOHONAN_PIPA_TYPES = ("P.Dinas", "P.Distribusi")
PERMOHONAN_DITINDAKLANJUTI = {
    "": None,
    "belum": None,
    "ya": 1,
    "tidak": 0,
    "0": 0,
    "1": 1,
}


def generate_template_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Pelanggan"

    headers = [
        "No", "Nama Pelanggan", "Alamat Pelanggan", "Kelurahan", "Kecamatan",
        "Nomor Instalasi", "Tanggal Pasang", "Diameter Pipa", "Tekanan Air",
        "Status (SIB/BK)", "Petugas", "Keterangan",
    ]
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="14203A", end_color="14203A")
    body_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D7DAE0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 30

    widths = [6, 22, 26, 16, 16, 16, 14, 12, 12, 16, 14, 22]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    example = [1, "Contoh Nama Pelanggan", "Jl. Contoh No.1", "Contoh Kelurahan",
               "Contoh Kecamatan", "83.001", "2026-01-15", "\u00d82", 0.4, "SIB", "", ""]
    for col, val in enumerate(example, start=1):
        c = ws.cell(row=2, column=col, value=val)
        c.font = body_font
        c.border = border

    ws.cell(row=2, column=6).number_format = "@"
    for r in range(2, 202):
        ws.cell(row=r, column=6).number_format = "@"
        ws.cell(row=r, column=7).number_format = "DD-MM-YYYY"
        for col in range(1, 13):
            ws.cell(row=r, column=col).border = border
            ws.cell(row=r, column=col).font = body_font

    dv = DataValidation(type="list", formula1='"SIB,BK"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("J2:J201")
    ws.freeze_panes = "A2"

    notes = wb.create_sheet("Catatan Pengisian")
    rows = [
        ("Kolom", "Catatan"),
        ("Nomor Instalasi", "Ketik apa adanya termasuk titik (contoh: 83.236). Jangan biarkan Excel mengubahnya jadi angka."),
        ("Nama Pelanggan", "Kalau satu pelanggan punya lebih dari satu instalasi, tulis nama persis sama di tiap baris -- itu bukan duplikat."),
        ("Status (SIB/BK)", "Pilih dari dropdown. SIB = Sambungan Instalasi Baru, BK = Buka Kembali."),
        ("Petugas", "Boleh dikosongkan."),
    ]
    for r, (a, b) in enumerate(rows, start=1):
        notes.cell(row=r, column=1, value=a).font = Font(bold=(r == 1))
        cb = notes.cell(row=r, column=2, value=b)
        cb.font = Font(bold=(r == 1))
        cb.alignment = Alignment(wrap_text=True, vertical="top")
    notes.column_dimensions["A"].width = 20
    notes.column_dimensions["B"].width = 75

    return wb


def parse_tanggal(val):
    if val is None or val == "":
        return None, None
    if isinstance(val, (datetime, date)):
        parsed = val if isinstance(val, datetime) else datetime.combine(val, datetime.min.time())
        if not (1990 <= parsed.year <= datetime.now().year + 2):
            return None, f'Tahun tanggal "{parsed.year}" tidak wajar'
        return parsed.strftime("%Y-%m-%d"), None
    raw = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y"):
        try:
            parsed = datetime.strptime(raw, fmt)
            if not (1990 <= parsed.year <= datetime.now().year + 2):
                return None, f'Tahun tanggal "{parsed.year}" tidak wajar'
            return parsed.strftime("%Y-%m-%d"), None
        except ValueError:
            continue
    return None, f'Format tanggal "{raw}" tidak dikenali, pakai format tanggal (bukan teks bebas)'


def validate_permohonan_data(data):
    errors = []

    if not data.get("nama_pelanggan", "").strip():
        errors.append("Nama pelanggan kosong")
    if not data.get("kecamatan", "").strip():
        errors.append("Kecamatan kosong")
    if data.get("jenis") not in VALID_STATUSES:
        errors.append("Jenis permohonan tidak valid")
    if not data.get("tanggal_permohonan", "").strip():
        errors.append("Tanggal permohonan kosong")
    elif not is_valid_iso_date(data.get("tanggal_permohonan")):
        errors.append("Tanggal permohonan tidak valid")

    tanggal_survey = data.get("tanggal_survey", "").strip()
    if tanggal_survey and not is_valid_iso_date(tanggal_survey):
        errors.append("Tanggal survey tidak valid")

    tanggal_dikirim = data.get("tanggal_dikirim_hublang", "").strip()
    if tanggal_dikirim and not is_valid_iso_date(tanggal_dikirim):
        errors.append("Tanggal dikirim ke Hublang tidak valid")

    tanggal_kembali = data.get("tanggal_kembali_hublang", "").strip()
    if tanggal_kembali and not is_valid_iso_date(tanggal_kembali):
        errors.append("Tanggal kembali ke Hublang tidak valid")

    return errors


def parse_upload(fileobj, db):
    wb = load_workbook(fileobj, data_only=True)
    if not wb.worksheets:
        raise ValueError("File Excel ini kosong, tidak ada sheet sama sekali.")
    # Pakai sheet PERTAMA di file, apapun namanya -- jangan maksa nama sheet
    # persis "Data Pelanggan". Kalau orang rename sheet, save-as, atau
    # duplikat tab pas ngedit manual, import tetap harus jalan selama
    # struktur kolomnya benar.
    ws = wb.worksheets[0]

    header_row = [c.value for c in ws[1]]
    col_index = {}
    for idx, h in enumerate(header_row):
        if h is None:
            continue
        key = HEADER_MAP.get(str(h).strip().lower())
        if key:
            col_index[key] = idx

    missing = [k for k in ["nama", "kecamatan", "nomor_instalasi", "tanggal_pasang", "status"] if k not in col_index]
    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan di header: {', '.join(missing)}")

    existing_nomor = {
        r["nomor_instalasi"] for r in db.execute("SELECT nomor_instalasi FROM instalasi").fetchall()
    }
    existing_pelanggan = {}
    for r in db.execute("SELECT id, nama, kecamatan FROM pelanggan").fetchall():
        existing_pelanggan[(r["nama"].strip().lower(), r["kecamatan"].strip().lower())] = r["id"]

    rows = []
    batch_pelanggan = {}
    next_batch_no = 1

    for row_cells in ws.iter_rows(min_row=2, values_only=False):
        values = {}
        for key, idx in col_index.items():
            values[key] = row_cells[idx].value if idx < len(row_cells) else None

        nama = str(values.get("nama") or "").strip()
        kecamatan = str(values.get("kecamatan") or "").strip()
        nomor_instalasi = str(values.get("nomor_instalasi") or "").strip()
        tanggal_pasang, tanggal_error = parse_tanggal(values.get("tanggal_pasang"))
        status = str(values.get("status") or "").strip().upper()

        if not any([nama, kecamatan, nomor_instalasi]):
            continue

        row_errors = []
        if not nama:
            row_errors.append("Nama kosong")
        if not kecamatan:
            row_errors.append("Kecamatan kosong")
        if not nomor_instalasi:
            row_errors.append("Nomor instalasi kosong")
        if tanggal_error:
            row_errors.append(tanggal_error)
        elif not tanggal_pasang:
            row_errors.append("Tanggal pasang kosong")
        if status not in ("SIB", "BK"):
            row_errors.append(f'Status "{status}" tidak valid (harus SIB atau BK)')
        if nomor_instalasi in existing_nomor:
            row_errors.append(f'Nomor instalasi "{nomor_instalasi}" sudah ada di database')

        key = (nama.lower(), kecamatan.lower())
        if key in existing_pelanggan:
            pelanggan_status = f"gabung ke pelanggan lama (id {existing_pelanggan[key]})"
        elif key in batch_pelanggan:
            pelanggan_status = f"gabung ke baris baru lain di file ini (#{batch_pelanggan[key]})"
        else:
            batch_pelanggan[key] = next_batch_no
            pelanggan_status = f"pelanggan baru (#{next_batch_no})"
            next_batch_no += 1

        rows.append({
            "nama": nama, "alamat": str(values.get("alamat") or ""),
            "kelurahan": str(values.get("kelurahan") or ""), "kecamatan": kecamatan,
            "nomor_instalasi": nomor_instalasi, "tanggal_pasang": tanggal_pasang or "",
            "diameter_pipa": str(values.get("diameter_pipa") or ""),
            "tekanan_air": values.get("tekanan_air"),
            "status": status, "petugas": str(values.get("petugas") or ""),
            "keterangan": str(values.get("keterangan") or ""),
            "pelanggan_status": pelanggan_status,
            "errors": row_errors,
        })
        existing_nomor.add(nomor_instalasi)

    return rows


def generate_permohonan_template_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data Permohonan"

    headers = [
        "No", "Nama Pelanggan", "Lokasi", "Kelurahan", "Kecamatan",
        "Jenis", "Tanggal Permohonan", "Tanggal Survey", "Petugas Survey",
        "Ditindaklanjuti", "Jenis Pipa", "Tanggal Dikirim Hublang",
        "Tanggal Kembali Hublang", "Keterangan",
    ]

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_fill = PatternFill("solid", start_color="14203A", end_color="14203A")
    body_font = Font(name="Arial", size=10)
    thin = Side(style="thin", color="D7DAE0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[1].height = 30

    widths = [6, 22, 24, 16, 16, 10, 14, 14, 16, 14, 14, 16, 16, 24]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    example = [
        1, "Contoh Nama Pelanggan", "Jl. Contoh No.1", "Contoh Kelurahan",
        "Contoh Kecamatan", "SIB", "2026-01-15", "2026-01-18", "Petugas A",
        "Belum", "P.Dinas", "2026-01-20", "2026-01-25",
        "Catatan khusus jika ada",
    ]
    for col, val in enumerate(example, start=1):
        c = ws.cell(row=2, column=col, value=val)
        c.font = body_font
        c.border = border

    for r in range(2, 202):
        for col in range(1, len(headers) + 1):
            ws.cell(row=r, column=col).border = border
            ws.cell(row=r, column=col).font = body_font
    for r in range(2, 202):
        ws.cell(row=r, column=7).number_format = "DD-MM-YYYY"
        ws.cell(row=r, column=8).number_format = "DD-MM-YYYY"
        ws.cell(row=r, column=12).number_format = "DD-MM-YYYY"
        ws.cell(row=r, column=13).number_format = "DD-MM-YYYY"

    dv_jenis = DataValidation(type="list", formula1='"SIB,BK"', allow_blank=True)
    dv_ditindaklanjuti = DataValidation(type="list", formula1='"Belum,Ya,Tidak"', allow_blank=True)
    dv_pipa = DataValidation(type="list", formula1='"P.Dinas,P.Distribusi"', allow_blank=True)
    ws.add_data_validation(dv_jenis)
    ws.add_data_validation(dv_ditindaklanjuti)
    ws.add_data_validation(dv_pipa)
    dv_jenis.add("F2:F201")
    dv_ditindaklanjuti.add("J2:J201")
    dv_pipa.add("K2:K201")
    ws.freeze_panes = "A2"

    notes = wb.create_sheet("Catatan Pengisian")
    rows = [
        ("Kolom", "Catatan"),
        ("Nama Pelanggan", "Nama lengkap pelanggan atau pemohon permohonan."),
        ("Kecamatan", "Kecamatan wajib diisi dan harus sesuai data wilayah."),
        ("Jenis", "Pilih SIB atau BK."),
        ("Tanggal Permohonan", "Gunakan format tanggal. Jika cell berubah format jadi teks, edit ulang.") ,
        ("Ditindaklanjuti", "Isi Belum / Ya / Tidak. Biarkan kosong untuk Belum."),
        ("Jenis Pipa", "P.Dinas atau P.Distribusi jika diketahui. Boleh dikosongkan."),
    ]
    for r, (a, b) in enumerate(rows, start=1):
        notes.cell(row=r, column=1, value=a).font = Font(bold=(r == 1))
        cb = notes.cell(row=r, column=2, value=b)
        cb.font = Font(bold=(r == 1))
        cb.alignment = Alignment(wrap_text=True, vertical="top")
    notes.column_dimensions["A"].width = 20
    notes.column_dimensions["B"].width = 80

    return wb


def parse_permohonan_upload(fileobj, db):
    wb = load_workbook(fileobj, data_only=True)
    if not wb.worksheets:
        raise ValueError("File Excel ini kosong, tidak ada sheet sama sekali.")
    ws = wb.worksheets[0]

    header_row = [c.value for c in ws[1]]
    col_index = {}
    for idx, h in enumerate(header_row):
        if h is None:
            continue
        key = PERMOHONAN_HEADER_MAP.get(str(h).strip().lower())
        if key:
            col_index[key] = idx

    missing = [k for k in PERMOHONAN_REQUIRED if k not in col_index]
    if missing:
        raise ValueError(f"Kolom wajib tidak ditemukan di header: {', '.join(missing)}")

    rows = []
    for row_cells in ws.iter_rows(min_row=2, values_only=False):
        values = {}
        for key, idx in col_index.items():
            values[key] = row_cells[idx].value if idx < len(row_cells) else None

        nama_pelanggan = str(values.get("nama_pelanggan") or "").strip()
        lokasi = str(values.get("lokasi") or "").strip()
        kelurahan = str(values.get("kelurahan") or "").strip()
        kecamatan = str(values.get("kecamatan") or "").strip()
        jenis = str(values.get("jenis") or "").strip().upper()
        tanggal_permohonan, tanggal_permohonan_error = parse_tanggal(values.get("tanggal_permohonan"))
        tanggal_survey, tanggal_survey_error = parse_tanggal(values.get("tanggal_survey"))
        petugas_survey = str(values.get("petugas_survey") or "").strip()
        ditindaklanjuti_raw = str(values.get("ditindaklanjuti") or "").strip()
        jenis_pipa = str(values.get("jenis_pipa") or "").strip()
        tanggal_dikirim_hublang, tanggal_dikirim_error = parse_tanggal(values.get("tanggal_dikirim_hublang"))
        tanggal_kembali_hublang, tanggal_kembali_error = parse_tanggal(values.get("tanggal_kembali_hublang"))
        keterangan = str(values.get("keterangan") or "").strip()

        if not any([nama_pelanggan, lokasi, kelurahan, kecamatan, jenis, values.get("tanggal_permohonan")]):
            continue

        row_errors = []
        if not nama_pelanggan:
            row_errors.append("Nama pelanggan kosong")
        if not kecamatan:
            row_errors.append("Kecamatan kosong")
        if jenis not in VALID_STATUSES:
            row_errors.append(f"Jenis '{jenis}' tidak valid (harus SIB atau BK)")
        if tanggal_permohonan_error:
            row_errors.append(tanggal_permohonan_error)
        elif not tanggal_permohonan:
            row_errors.append("Tanggal permohonan kosong")
        if tanggal_survey_error:
            row_errors.append(tanggal_survey_error)
        if tanggal_dikirim_error:
            row_errors.append(tanggal_dikirim_error)
        if tanggal_kembali_error:
            row_errors.append(tanggal_kembali_error)
        if jenis_pipa and jenis_pipa not in PERMOHONAN_PIPA_TYPES:
            row_errors.append(f"Jenis pipa '{jenis_pipa}' tidak valid")

        ditindaklanjuti_key = ditindaklanjuti_raw.lower()
        valid_ditindaklanjuti_keys = set(PERMOHONAN_DITINDAKLANJUTI)
        if ditindaklanjuti_raw and ditindaklanjuti_key not in valid_ditindaklanjuti_keys:
            row_errors.append(f"Ditindaklanjuti '{ditindaklanjuti_raw}' tidak valid (Belum / Ya / Tidak)")
        ditindaklanjuti = PERMOHONAN_DITINDAKLANJUTI.get(ditindaklanjuti_key)

        rows.append({
            "nama_pelanggan": nama_pelanggan,
            "lokasi": lokasi,
            "kelurahan": kelurahan,
            "kecamatan": kecamatan,
            "jenis": jenis,
            "tanggal_permohonan": tanggal_permohonan or "",
            "tanggal_survey": tanggal_survey or "",
            "petugas_survey": petugas_survey,
            "ditindaklanjuti": ditindaklanjuti,
            "ditindaklanjuti_display": ditindaklanjuti_raw or "Belum",
            "jenis_pipa": jenis_pipa,
            "tanggal_dikirim_hublang": tanggal_dikirim_hublang or "",
            "tanggal_kembali_hublang": tanggal_kembali_hublang or "",
            "keterangan": keterangan,
            "errors": row_errors,
        })

    return rows


def permohonan_ditindaklanjuti_label(value):
    if value == 1:
        return "Ya"
    if value == 0:
        return "Tidak"
    return "Belum"


@app.context_processor
def inject_helpers():

    db = get_db()
    return dict(
        semua_kecamatan=daftar_kecamatan(db),
        today=datetime.now().strftime("%d %B %Y"),
    )


# ---------------- Halaman Utama Dashboard (Client-Sided) ----------------
@app.route("/")
def dashboard():
    return render_template("dashboard.html")


# ---------------- Endpoint API Data Dashboard (REST API) ----------------
@app.route("/api/dashboard-stats")
def api_dashboard_stats():
    db = get_db()

    total_pelanggan = db.execute("SELECT COUNT(*) c FROM pelanggan").fetchone()["c"]
    total_kecamatan = db.execute("SELECT COUNT(DISTINCT kecamatan) c FROM pelanggan").fetchone()["c"]
    pasang_bulan_ini = db.execute(
        "SELECT COUNT(*) c FROM instalasi WHERE strftime('%Y-%m', tanggal_pasang) = strftime('%Y-%m','now')"
    ).fetchone()["c"]
    jumlah_sib = db.execute("SELECT COUNT(*) c FROM instalasi WHERE status='SIB'").fetchone()["c"]
    jumlah_bk = db.execute("SELECT COUNT(*) c FROM instalasi WHERE status='BK'").fetchone()["c"]

    per_kecamatan_rows = db.execute(
        "SELECT kecamatan, COUNT(*) c FROM pelanggan GROUP BY kecamatan ORDER BY c DESC"
    ).fetchall()
    per_kecamatan = [{"kecamatan": r["kecamatan"], "c": r["c"]} for r in per_kecamatan_rows]

    per_bulan_rows = db.execute(
        """SELECT strftime('%Y-%m', tanggal_pasang) bulan, COUNT(*) c
           FROM instalasi
           GROUP BY bulan ORDER BY bulan ASC LIMIT 6"""
    ).fetchall()
    per_bulan = [{"bulan": r["bulan"], "c": r["c"]} for r in per_bulan_rows]

    recent_rows = db.execute(
        """SELECT p.nama, p.kecamatan, i.nomor_instalasi, i.tanggal_pasang, i.status
           FROM instalasi i JOIN pelanggan p ON p.id = i.pelanggan_id
           ORDER BY i.id DESC LIMIT 5"""
    ).fetchall()
    recent = [{
        "nama": r["nama"],
        "kecamatan": r["kecamatan"],
        "nomor_instalasi": r["nomor_instalasi"],
        "tanggal_pasang": r["tanggal_pasang"],
        "status": r["status"]
    } for r in recent_rows]

    total_permohonan = db.execute("SELECT COUNT(*) c FROM permohonan").fetchone()["c"]
    permohonan_menunggu = db.execute("SELECT COUNT(*) c FROM permohonan WHERE ditindaklanjuti IS NULL").fetchone()["c"]
    permohonan_selisih = db.execute("SELECT COUNT(*) c FROM permohonan WHERE ditindaklanjuti = 0").fetchone()["c"]
    permohonan_selesai = db.execute("SELECT COUNT(*) c FROM permohonan WHERE instalasi_id IS NOT NULL").fetchone()["c"]
    permohonan_bulan_ini = db.execute(
        "SELECT COUNT(*) c FROM permohonan WHERE strftime('%Y-%m', tanggal_permohonan) = strftime('%Y-%m','now')"
    ).fetchone()["c"]

    recent_permohonan_rows = db.execute(
        "SELECT nama_pelanggan, kecamatan, jenis, tanggal_permohonan FROM permohonan ORDER BY id DESC LIMIT 5"
    ).fetchall()
    recent_permohonan = [{
        "nama_pelanggan": r["nama_pelanggan"], "kecamatan": r["kecamatan"],
        "jenis": r["jenis"], "tanggal_permohonan": r["tanggal_permohonan"],
    } for r in recent_permohonan_rows]

    return jsonify({
        "total_pelanggan": total_pelanggan,
        "total_kecamatan": total_kecamatan,
        "pasang_bulan_ini": pasang_bulan_ini,
        "sib_bk": f"{jumlah_sib} / {jumlah_bk}",
        "per_kecamatan": per_kecamatan,
        "per_bulan": per_bulan,
        "recent": recent,
        "total_permohonan": total_permohonan,
        "permohonan_menunggu": permohonan_menunggu,
        "permohonan_selisih": permohonan_selisih,
        "permohonan_selesai": permohonan_selesai,
        "permohonan_bulan_ini": permohonan_bulan_ini,
        "recent_permohonan": recent_permohonan,
    })


# ---------------- Semua Pelanggan (list + cari + filter kecamatan) ----------------
@app.route("/pelanggan/template")
def pelanggan_template():
    wb = generate_template_workbook()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="template_data_pelanggan_sidap.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/pelanggan/impor", methods=["GET", "POST"])
def pelanggan_impor():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Pilih file Excel dulu.", "error")
            return redirect(url_for("pelanggan_impor"))
        try:
            rows = parse_upload(file, get_db())
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("pelanggan_impor"))
        except Exception:
            flash("File tidak bisa dibaca. Pastikan formatnya .xlsx dan pakai template yang disediakan.", "error")
            return redirect(url_for("pelanggan_impor"))

        if not rows:
            flash("Tidak ada baris data yang terbaca dari file ini.", "error")
            return redirect(url_for("pelanggan_impor"))

        os.makedirs(IMPORT_TMP_DIR, exist_ok=True)
        token = uuid.uuid4().hex
        with open(os.path.join(IMPORT_TMP_DIR, f"{token}.json"), "w") as f:
            json.dump(rows, f)

        jumlah_error = sum(1 for r in rows if r["errors"])
        jumlah_baru = len({r["pelanggan_status"] for r in rows if "pelanggan baru" in r["pelanggan_status"]})
        return render_template(
            "import_preview.html", rows=rows, token=token,
            jumlah_error=jumlah_error, jumlah_baru=jumlah_baru, total=len(rows),
        )

    return render_template("import_upload.html")


@app.route("/pelanggan/impor/konfirmasi", methods=["POST"])
def pelanggan_impor_konfirmasi():
    token = request.form.get("token", "")
    path = os.path.join(IMPORT_TMP_DIR, f"{token}.json")
    if not os.path.exists(path):
        flash("Sesi impor sudah kedaluwarsa, silakan upload ulang.", "error")
        return redirect(url_for("pelanggan_impor"))

    with open(path) as f:
        rows = json.load(f)
    os.remove(path)

    valid_rows = [r for r in rows if not r["errors"]]
    if not valid_rows:
        flash("Tidak ada baris valid yang bisa disimpan -- semua baris punya error.", "error")
        return redirect(url_for("pelanggan_impor"))

    db = get_db()
    pelanggan_cache = {}
    for r in db.execute("SELECT id, nama, kecamatan FROM pelanggan").fetchall():
        pelanggan_cache[(r["nama"].strip().lower(), r["kecamatan"].strip().lower())] = r["id"]

    tersimpan = 0
    for r in valid_rows:
        key = (r["nama"].strip().lower(), r["kecamatan"].strip().lower())
        if key not in pelanggan_cache:
            cur = db.execute(
                "INSERT INTO pelanggan (nama, alamat, kelurahan, kecamatan) VALUES (?,?,?,?)",
                (r["nama"], r["alamat"], r["kelurahan"], r["kecamatan"]),
            )
            pelanggan_cache[key] = cur.lastrowid

        db.execute(
            """INSERT INTO instalasi
               (pelanggan_id, nomor_instalasi, tanggal_pasang, diameter_pipa,
                tekanan_air, status, petugas, keterangan)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                pelanggan_cache[key], r["nomor_instalasi"], r["tanggal_pasang"],
                r["diameter_pipa"], r["tekanan_air"] or None, r["status"],
                r["petugas"] or None, r["keterangan"],
            ),
        )
        tersimpan += 1

    db.commit()
    dilewati = len(rows) - tersimpan
    pesan = f"{tersimpan} baris berhasil diimpor."
    if dilewati:
        pesan += f" {dilewati} baris dilewati karena ada error."
    flash(pesan)
    return redirect(url_for("pelanggan_list"))


@app.route("/permohonan/template")
def permohonan_template():
    wb = generate_permohonan_template_workbook()
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="template_data_permohonan_sidap.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/permohonan/impor", methods=["GET", "POST"])
def permohonan_impor():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Pilih file Excel dulu.", "error")
            return redirect(url_for("permohonan_impor"))
        try:
            rows = parse_permohonan_upload(file, get_db())
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("permohonan_impor"))
        except Exception:
            flash("File tidak bisa dibaca. Pastikan formatnya .xlsx dan pakai template yang disediakan.", "error")
            return redirect(url_for("permohonan_impor"))

        if not rows:
            flash("Tidak ada baris data yang terbaca dari file ini.", "error")
            return redirect(url_for("permohonan_impor"))

        os.makedirs(IMPORT_TMP_DIR, exist_ok=True)
        token = uuid.uuid4().hex
        with open(os.path.join(IMPORT_TMP_DIR, f"{token}.json"), "w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False)

        jumlah_error = sum(1 for r in rows if r["errors"])
        jumlah_valid = len(rows) - jumlah_error
        return render_template(
            "permohonan_import_preview.html", rows=rows, token=token,
            jumlah_error=jumlah_error, jumlah_valid=jumlah_valid, total=len(rows),
        )

    return render_template("permohonan_import_upload.html")


@app.route("/permohonan/impor/konfirmasi", methods=["POST"])
def permohonan_impor_konfirmasi():
    token = request.form.get("token", "")
    path = os.path.join(IMPORT_TMP_DIR, f"{token}.json")
    if not os.path.exists(path):
        flash("Sesi impor sudah kedaluwarsa, silakan upload ulang.", "error")
        return redirect(url_for("permohonan_impor"))

    with open(path, encoding="utf-8") as f:
        rows = json.load(f)
    os.remove(path)

    valid_rows = [r for r in rows if not r["errors"]]
    if not valid_rows:
        flash("Tidak ada baris valid yang bisa disimpan -- semua baris punya error.", "error")
        return redirect(url_for("permohonan_impor"))

    db = get_db()
    tersimpan = 0
    for r in valid_rows:
        db.execute(
            """INSERT INTO permohonan
               (nama_pelanggan, lokasi, kelurahan, kecamatan, jenis, tanggal_permohonan,
                tanggal_survey, petugas_survey, ditindaklanjuti, jenis_pipa,
                tanggal_dikirim_hublang, tanggal_kembali_hublang, keterangan)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                r["nama_pelanggan"], r["lokasi"], r["kelurahan"], r["kecamatan"], r["jenis"],
                r["tanggal_permohonan"], r["tanggal_survey"] or None,
                r["petugas_survey"] or None, r["ditindaklanjuti"], r["jenis_pipa"] or None,
                r["tanggal_dikirim_hublang"] or None, r["tanggal_kembali_hublang"] or None,
                r["keterangan"],
            ),
        )
        tersimpan += 1
    db.commit()

    dilewati = len(rows) - tersimpan
    pesan = f"{tersimpan} baris permohonan berhasil diimpor."
    if dilewati:
        pesan += f" {dilewati} baris dilewati karena ada error."
    flash(pesan)
    return redirect(url_for("permohonan_list"))


# ---------------- Cari Pelanggan ----------------
@app.route("/cari")
def cari_pelanggan():

    # Server hanya melempar halaman HTML kosong tanpa memproses database
    return render_template("cari_pelanggan.html")


@app.route("/api/cari")
def api_cari_pelanggan():
    db = get_db()
    q = request.args.get("q", "").strip()
    scope = request.args.get("scope", "semua")
    rows = []

    if q:
        query = """SELECT p.id pelanggan_id, p.nama, p.kecamatan,
                          i.id instalasi_id, i.nomor_instalasi, i.tanggal_pasang, i.status
                   FROM instalasi i JOIN pelanggan p ON p.id = i.pelanggan_id
                   WHERE 1=1"""
        params = []
        if scope == "nama":
            query += " AND p.nama LIKE ?"
            params.append(f"%{q}%")
        elif scope == "instalasi":
            query += " AND i.nomor_instalasi LIKE ?"
            params.append(f"%{q}%")
        elif scope == "kecamatan":
            query += " AND p.kecamatan LIKE ?"
            params.append(f"%{q}%")
        else:
            query += " AND (p.nama LIKE ? OR i.nomor_instalasi LIKE ? OR p.kecamatan LIKE ?)"
            params += [f"%{q}%", f"%{q}%", f"%{q}%"]
        query += " ORDER BY p.nama, i.tanggal_pasang"

        found = db.execute(query, params).fetchall()
        seq_counter = {}
        for r in found:
            seq_counter[r["pelanggan_id"]] = seq_counter.get(r["pelanggan_id"], 0) + 1
            d = dict(r)
            d["urutan_instalasi"] = seq_counter[r["pelanggan_id"]]
            rows.append(d)

    # Mengembalikan hasil pencarian dalam bentuk JSON mentah
    return jsonify({
        "q": q,
        "scope": scope,
        "results": rows
    })

# 1. Rute untuk menampilkan halaman (Ini yang membuat halaman Semua Pelanggan terbuka)
@app.route("/pelanggan")
def pelanggan_list():
    return render_template("pelanggan_list.html")


# 2. Rute API untuk menyuplai data (INI YANG HILANG ATAU BELUM TERSIMPAN)
@app.route("/api/pelanggan")
def api_pelanggan_list():
    db = get_db()
    q = request.args.get("q", "").strip()
    kecamatan = request.args.get("kecamatan", "semua")
    status = request.args.get("status", "semua")
    tahun = request.args.get("tahun", "").strip()
    bulan = request.args.get("bulan", "").strip()
    page = max(1, request.args.get("page", 1, type=int))
    per_page = 30

    conditions = []
    params = []
    if q:
        conditions.append(
            "(p.nama LIKE ? OR EXISTS (SELECT 1 FROM instalasi WHERE pelanggan_id=p.id AND nomor_instalasi LIKE ?))"
        )
        params += [f"%{q}%", f"%{q}%"]
    if kecamatan != "semua":
        conditions.append("p.kecamatan = ?")
        params.append(kecamatan)

    inst_conditions = []
    inst_params = []
    if status != "semua":
        inst_conditions.append("status = ?")
        inst_params.append(status)
    if tahun:
        inst_conditions.append("strftime('%Y', tanggal_pasang) = ?")
        inst_params.append(tahun)
    if bulan:
        inst_conditions.append("strftime('%m', tanggal_pasang) = ?")
        inst_params.append(bulan.zfill(2))
    if inst_conditions:
        conditions.append(
            f"EXISTS (SELECT 1 FROM instalasi WHERE pelanggan_id=p.id AND {' AND '.join(inst_conditions)})"
        )
        params += inst_params

    # ... (bagian parameter query string tetap sama) ...

    where_clause = " AND ".join(conditions) if conditions else "1=1"
    
    # Tambahkan kata DISTINCT di dalam group_concat agar status yang kembar disaring
    query = f"""SELECT p.id pelanggan_id, p.nama, p.alamat, p.kelurahan, p.kecamatan,
                       (SELECT COUNT(*) FROM instalasi WHERE pelanggan_id=p.id) jumlah_instalasi,
                       (SELECT group_concat(DISTINCT status) FROM instalasi WHERE pelanggan_id=p.id) semua_status,
                       (SELECT MAX(tanggal_pasang) FROM instalasi WHERE pelanggan_id=p.id) tanggal_terbaru
                FROM pelanggan p
                WHERE {where_clause}
                ORDER BY p.nama"""

    all_rows = db.execute(query, params).fetchall()
    total = len(all_rows)
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = min(page, total_pages)
    
    sliced_rows = all_rows[(page - 1) * per_page: page * per_page]

    # Masukkan data semua_status ke dalam JSON response
    rows = [{
        "pelanggan_id": r["pelanggan_id"],
        "nama": r["nama"],
        "alamat": r["alamat"],
        "kelurahan": r["kelurahan"],
        "kecamatan": r["kecamatan"],
        "jumlah_instalasi": r["jumlah_instalasi"],
        "semua_status": r["semua_status"] if r["semua_status"] else "",
        "tanggal_terbaru": r["tanggal_terbaru"] if r["tanggal_terbaru"] else ""
    } for r in sliced_rows]

    kecamatan_list = daftar_kecamatan(db)

    return jsonify({
        "rows": rows,
        "total": total,
        "page": page,
        "total_pages": total_pages,
        "semua_kecamatan": kecamatan_list
    })


@app.route("/pelanggan/export")
def pelanggan_export():
    db = get_db()
    q = request.args.get("q", "").strip()
    kecamatan = request.args.get("kecamatan", "semua")
    status = request.args.get("status", "semua")
    tahun = request.args.get("tahun", "").strip()
    bulan = request.args.get("bulan", "").strip()

    conditions = []
    params = []
    if q:
        conditions.append(
            "(p.nama LIKE ? OR EXISTS (SELECT 1 FROM instalasi WHERE pelanggan_id=p.id AND nomor_instalasi LIKE ?))"
        )
        params += [f"%{q}%", f"%{q}%"]
    if kecamatan != "semua":
        conditions.append("p.kecamatan = ?")
        params.append(kecamatan)

    inst_conditions = []
    inst_params = []
    if status != "semua":
        inst_conditions.append("status = ?")
        inst_params.append(status)
    if tahun:
        inst_conditions.append("strftime('%Y', tanggal_pasang) = ?")
        inst_params.append(tahun)
    if bulan:
        inst_conditions.append("strftime('%m', tanggal_pasang) = ?")
        inst_params.append(bulan.zfill(2))
    if inst_conditions:
        conditions.append(
            f"EXISTS (SELECT 1 FROM instalasi WHERE pelanggan_id=p.id AND {' AND '.join(inst_conditions)})"
        )
        params += inst_params

    where_clause = " AND ".join(conditions) if conditions else "1=1"
    rows = db.execute(
        f"""SELECT p.nama, p.alamat, p.kelurahan, p.kecamatan,
                          (SELECT group_concat(DISTINCT status) FROM instalasi WHERE pelanggan_id=p.id) semua_status,
                          (SELECT MAX(tanggal_pasang) FROM instalasi WHERE pelanggan_id=p.id) tanggal_terbaru,
                          (SELECT COUNT(*) FROM instalasi WHERE pelanggan_id=p.id) jumlah_instalasi
                   FROM pelanggan p
                   WHERE {where_clause}
                   ORDER BY p.nama""",
        params,
    ).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Semua Pelanggan"
    ws.append([
        "Nama", "Alamat", "Kelurahan", "Kecamatan",
        "Status", "Tanggal Pasang Terakhir", "Jumlah Instalasi"
    ])
    for r in rows:
        ws.append([
            r["nama"], r["alamat"], r["kelurahan"], r["kecamatan"],
            r["semua_status"] or "", r["tanggal_terbaru"] or "", r["jumlah_instalasi"],
        ])

    for idx, width in enumerate([25, 35, 20, 20, 15, 18, 16], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"data_pelanggan_{tahun or 'semua'}_{bulan or 'semua'}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------- Detail satu pelanggan ----------------
@app.route("/pelanggan/<int:pid>")
def pelanggan_detail(pid):
    db = get_db()
    pelanggan = db.execute("SELECT * FROM pelanggan WHERE id=?", (pid,)).fetchone()
    instalasi = db.execute(
        "SELECT * FROM instalasi WHERE pelanggan_id=? ORDER BY tanggal_pasang", (pid,)
    ).fetchall()
    return render_template("pelanggan_detail.html", pelanggan=pelanggan, instalasi=instalasi)


# ---------------- Tambah pelanggan + instalasi ----------------
@app.route("/pelanggan/tambah", methods=["GET", "POST"])
def pelanggan_tambah():
    db = get_db()
    if request.method == "POST":
        nomor_instalasi = request.form["nomor_instalasi"].strip()
        pelanggan_existing_id = request.form.get("pelanggan_existing_id", "").strip()

        dup = db.execute(
            "SELECT 1 FROM instalasi WHERE nomor_instalasi=?", (nomor_instalasi,)
        ).fetchone()
        if dup:
            flash(f"Nomor instalasi \"{nomor_instalasi}\" sudah dipakai instalasi lain. Cek kembali sebelum simpan.", "error")
            return render_template("pelanggan_form.html", pelanggan=None, form_data=request.form,
                                    semua_pelanggan_pilihan=daftar_pelanggan_pilihan(db))

        if not is_valid_iso_date(request.form.get("tanggal_pasang", "")):
            flash("Tanggal pasang tidak valid. Pakai date picker, jangan ketik manual.", "error")
            return render_template("pelanggan_form.html", pelanggan=None, form_data=request.form,
                                    semua_pelanggan_pilihan=daftar_pelanggan_pilihan(db))

        if pelanggan_existing_id:
            existing = db.execute("SELECT id, nama FROM pelanggan WHERE id=?", (pelanggan_existing_id,)).fetchone()
            if not existing:
                flash("Pelanggan yang dipilih tidak ditemukan. Pilih dari daftar, jangan ketik manual.", "error")
                return render_template("pelanggan_form.html", pelanggan=None, form_data=request.form,
                                        semua_pelanggan_pilihan=daftar_pelanggan_pilihan(db))
            pelanggan_id = existing["id"]
            nama_untuk_pesan = existing["nama"]
        else:
            cur = db.execute(
                "INSERT INTO pelanggan (nama, alamat, kelurahan, kecamatan) VALUES (?,?,?,?)",
                (
                    request.form["nama"],
                    request.form.get("alamat", ""),
                    request.form.get("kelurahan", ""),
                    request.form["kecamatan"],
                ),
            )
            pelanggan_id = cur.lastrowid
            nama_untuk_pesan = request.form["nama"]

        db.execute(
            """INSERT INTO instalasi
               (pelanggan_id, nomor_instalasi, tanggal_pasang, diameter_pipa,
                tekanan_air, status, petugas, keterangan)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                pelanggan_id,
                nomor_instalasi,
                request.form["tanggal_pasang"],
                request.form.get("diameter_pipa", ""),
                request.form.get("tekanan_air") or None,
                request.form["status"],
                request.form.get("petugas") or None,
                request.form.get("keterangan", ""),
            ),
        )
        db.commit()
        flash(f"Instalasi {nomor_instalasi} tersimpan untuk {nama_untuk_pesan}.")
        return redirect(url_for("pelanggan_list"))

    return render_template("pelanggan_form.html", pelanggan=None, form_data=None,
                            semua_pelanggan_pilihan=daftar_pelanggan_pilihan(db))


# ---------------- Tambah instalasi baru ----------------
@app.route("/pelanggan/<int:pid>/instalasi/tambah", methods=["GET", "POST"])
def instalasi_tambah(pid):
    db = get_db()
    pelanggan = db.execute("SELECT * FROM pelanggan WHERE id=?", (pid,)).fetchone()
    if request.method == "POST":
        nomor_instalasi = request.form["nomor_instalasi"].strip()
        dup = db.execute(
            "SELECT 1 FROM instalasi WHERE nomor_instalasi=?", (nomor_instalasi,)
        ).fetchone()
        if dup:
            flash(f"Nomor instalasi \"{nomor_instalasi}\" sudah dipakai instalasi lain. Cek kembali sebelum simpan.", "error")
            return render_template("instalasi_form.html", pelanggan=pelanggan, form_data=request.form)

        if not is_valid_iso_date(request.form.get("tanggal_pasang", "")):
            flash("Tanggal pasang tidak valid. Pakai date picker, jangan ketik manual.", "error")
            return render_template("instalasi_form.html", pelanggan=pelanggan, form_data=request.form)

        db.execute(
            """INSERT INTO instalasi
               (pelanggan_id, nomor_instalasi, tanggal_pasang, diameter_pipa,
                tekanan_air, status, petugas, keterangan)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                pid,
                nomor_instalasi,
                request.form["tanggal_pasang"],
                request.form.get("diameter_pipa", ""),
                request.form.get("tekanan_air") or None,
                request.form["status"],
                request.form.get("petugas") or None,
                request.form.get("keterangan", ""),
            ),
        )
        db.commit()
        flash(f"Instalasi baru untuk {pelanggan['nama']} tersimpan.")
        return redirect(url_for("pelanggan_detail", pid=pid))

    return render_template("instalasi_form.html", pelanggan=pelanggan, form_data=None)


# ---------------- Edit / hapus instalasi ----------------
@app.route("/instalasi/<int:iid>/edit", methods=["GET", "POST"])
def instalasi_edit(iid):
    db = get_db()
    inst = db.execute("SELECT * FROM instalasi WHERE id=?", (iid,)).fetchone()
    pelanggan = db.execute("SELECT * FROM pelanggan WHERE id=?", (inst["pelanggan_id"],)).fetchone()

    if request.method == "POST":
        nomor_instalasi = request.form["nomor_instalasi"].strip()
        dup = db.execute(
            "SELECT 1 FROM instalasi WHERE nomor_instalasi=? AND id!=?", (nomor_instalasi, iid)
        ).fetchone()
        if dup:
            flash(f"Nomor instalasi \"{nomor_instalasi}\" sudah dipakai instalasi lain. Cek kembali sebelum simpan.", "error")
            return render_template("instalasi_edit.html", instalasi=inst, pelanggan=pelanggan)

        if not is_valid_iso_date(request.form.get("tanggal_pasang", "")):
            flash("Tanggal pasang tidak valid. Pakai date picker, jangan ketik manual.", "error")
            return render_template("instalasi_edit.html", instalasi=inst, pelanggan=pelanggan)

        db.execute(
            """UPDATE instalasi SET nomor_instalasi=?, tanggal_pasang=?, diameter_pipa=?,
               tekanan_air=?, status=?, petugas=?, keterangan=? WHERE id=?""",
            (
                nomor_instalasi,
                request.form["tanggal_pasang"],
                request.form.get("diameter_pipa", ""),
                request.form.get("tekanan_air") or None,
                request.form["status"],
                request.form.get("petugas") or None,
                request.form.get("keterangan", ""),
                iid,
            ),
        )
        db.commit()
        flash("Instalasi diperbarui.")
        return redirect(url_for("pelanggan_detail", pid=inst["pelanggan_id"]))

    return render_template("instalasi_edit.html", instalasi=inst, pelanggan=pelanggan)


@app.route("/instalasi/<int:iid>/hapus", methods=["POST"])
def instalasi_hapus(iid):
    db = get_db()
    inst = db.execute("SELECT pelanggan_id FROM instalasi WHERE id=?", (iid,)).fetchone()
    db.execute("DELETE FROM instalasi WHERE id=?", (iid,))
    db.commit()
    flash("Instalasi dihapus.")
    return redirect(url_for("pelanggan_detail", pid=inst["pelanggan_id"]))


# ---------------- Edit / hapus pelanggan ----------------
@app.route("/pelanggan/<int:pid>/edit", methods=["GET", "POST"])
def pelanggan_edit(pid):
    db = get_db()
    if request.method == "POST":
        db.execute(
            "UPDATE pelanggan SET nama=?, alamat=?, kelurahan=?, kecamatan=? WHERE id=?",
            (
                request.form["nama"],
                request.form.get("alamat", ""),
                request.form.get("kelurahan", ""),
                request.form["kecamatan"],
                pid,
            ),
        )
        db.commit()
        flash("Data pelanggan diperbarui.")
        return redirect(url_for("pelanggan_detail", pid=pid))

    pelanggan = db.execute("SELECT * FROM pelanggan WHERE id=?", (pid,)).fetchone()
    return render_template("pelanggan_form.html", pelanggan=pelanggan)


@app.route("/pelanggan/<int:pid>/hapus", methods=["POST"])
def pelanggan_hapus(pid):
    db = get_db()
    db.execute("DELETE FROM instalasi WHERE pelanggan_id=?", (pid,))
    db.execute("DELETE FROM pelanggan WHERE id=?", (pid,))
    db.commit()
    flash("Pelanggan dan seluruh instalasinya dihapus.")
    return redirect(url_for("pelanggan_list"))


# ---------------- Menu Kecamatan ----------------
@app.route("/kecamatan")
def kecamatan_list():
    db = get_db()
    rows = db.execute(
        """SELECT p.kecamatan,
                  COUNT(DISTINCT p.id) jumlah_pelanggan,
                  SUM(CASE WHEN i.status='SIB' THEN 1 ELSE 0 END) jumlah_sib,
                  SUM(CASE WHEN i.status='BK' THEN 1 ELSE 0 END) jumlah_bk
           FROM pelanggan p LEFT JOIN instalasi i ON i.pelanggan_id = p.id
           GROUP BY p.kecamatan ORDER BY p.kecamatan"""
    ).fetchall()
    return render_template("kecamatan_list.html", rows=rows)

@app.route("/api/kecamatan-stats")
def api_kecamatan_stats():
    db = get_db()
    rows = db.execute(
        """SELECT p.kecamatan,
                  COUNT(DISTINCT p.id) jumlah_pelanggan,
                  SUM(CASE WHEN i.status='SIB' THEN 1 ELSE 0 END) jumlah_sib,
                  SUM(CASE WHEN i.status='BK' THEN 1 ELSE 0 END) jumlah_bk
           FROM pelanggan p LEFT JOIN instalasi i ON i.pelanggan_id = p.id
           GROUP BY p.kecamatan ORDER BY p.kecamatan"""
    ).fetchall()
    
    data = [{
        "kecamatan": r["kecamatan"],
        "jumlah_pelanggan": r["jumlah_pelanggan"],
        "jumlah_sib": r["jumlah_sib"] or 0,
        "jumlah_bk": r["jumlah_bk"] or 0
    } for r in rows]
    
    return jsonify(data)

# ---------------- Menu Tahun & Bulan ----------------
# ---------------- Menu Tahun & Bulan (REST API & Client-Sided) ----------------
@app.route("/periode")
def periode_list():
    return render_template("periode_list.html")


@app.route("/api/periode-stats")
def api_periode_stats():
    db = get_db()
    
    # 1. Ambil daftar tahun unik untuk dropdown
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_pasang) t FROM instalasi
               WHERE tanggal_pasang IS NOT NULL AND strftime('%Y', tanggal_pasang) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    
    # Tentukan tahun terpilih (default tahun terbaru)
    tahun_terpilih = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    if tahun_list and tahun_terpilih not in tahun_list:
        tahun_terpilih = tahun_list[0]

    # 2. Ambil data pemasangan per bulan pada tahun tersebut
    per_bulan = db.execute(
        """SELECT strftime('%m', tanggal_pasang) bulan, COUNT(*) jumlah
           FROM instalasi
           WHERE strftime('%Y', tanggal_pasang) = ?
           GROUP BY bulan""",
        (tahun_terpilih,),
    ).fetchall()
    
    bulan_map = {r["bulan"]: r["jumlah"] for r in per_bulan}
    nama_bulan = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
    
    ringkasan = [
        {"nomor": f"{i + 1:02d}", "nama": nama, "jumlah": bulan_map.get(f"{i + 1:02d}", 0)}
        for i, nama in enumerate(nama_bulan)
    ]
    max_jumlah = max([r["jumlah"] for r in ringkasan], default=1) or 1

    # 3. Hitung jumlah data yang tanggal pasangnya rusak/kosong
    rusak = db.execute(
        """SELECT COUNT(*) c FROM instalasi
           WHERE tanggal_pasang IS NULL OR strftime('%Y', tanggal_pasang) IS NULL"""
    ).fetchone()["c"]

    return jsonify({
        "tahun_list": tahun_list,
        "tahun_terpilih": tahun_terpilih,
        "ringkasan": ringkasan,
        "max_jumlah": max_jumlah,
        "rusak": rusak
    })


# ---------------- Rekapitulasi Laporan Bulanan ----------------
NAMA_BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
              "Juli", "Agustus", "September", "Oktober", "November", "Desember"]


def hitung_rekap_bulanan(db, tahun, bulan):
    semua_kec = daftar_kecamatan(db)
    data = db.execute(
        """SELECT p.kecamatan,
                  SUM(CASE WHEN i.status='SIB' THEN 1 ELSE 0 END) sib,
                  SUM(CASE WHEN i.status='BK' THEN 1 ELSE 0 END) bk
           FROM instalasi i JOIN pelanggan p ON p.id = i.pelanggan_id
           WHERE strftime('%Y', i.tanggal_pasang) = ? AND strftime('%m', i.tanggal_pasang) = ?
           GROUP BY p.kecamatan""",
        (tahun, bulan),
    ).fetchall()
    peta = {r["kecamatan"]: {"sib": r["sib"] or 0, "bk": r["bk"] or 0} for r in data}

    rows = []
    total_sib = total_bk = 0
    for kec in semua_kec:
        sib = peta.get(kec, {}).get("sib", 0)
        bk = peta.get(kec, {}).get("bk", 0)
        rows.append({"kecamatan": kec, "sib": sib, "bk": bk, "jumlah": sib + bk})
        total_sib += sib
        total_bk += bk

    return rows, total_sib, total_bk


@app.route("/laporan")
def laporan_bulanan():
    # Hanya melempar halaman kosong agar browser yang menyusun tabelnya
    return render_template("laporan_bulanan.html")

@app.route("/laporan/unduh")
def laporan_unduh_xlsx():
    import io
    import openpyxl
    from flask import send_file

    db = get_db()
    tahun = request.args.get("tahun", "").strip()
    bulan = request.args.get("bulan", "").strip()
    
    # 1. Ambil data rekap dari fungsi helper bawaan
    rows, total_sib, total_bk = hitung_rekap_bulanan(db, tahun, bulan)
    nama_bulan = NAMA_BULAN[int(bulan) - 1] if bulan else ""

    # 2. Membuat file Excel baru menggunakan openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Laporan Bulanan"

    # Baris Judul Laporan
    ws.append(["LAPORAN BULANAN"])
    ws.append([f"BULAN: {nama_bulan.upper()} {tahun}"])
    ws.append([])  # Baris kosong untuk jarak

    # Header Tabel Excel
    ws.append(["NO", "KECAMATAN", "SIB", "BK", "JUMLAH"])

    # Isi Data Kecamatan
    for idx, r in enumerate(rows, start=1):
        ws.append([
            idx,
            r["kecamatan"].upper(),
            r["sib"] if r["sib"] > 0 else "-",
            r["bk"] if r["bk"] > 0 else "-",
            r["jumlah"]
        ])

    # Baris Total Paling Bawah
    ws.append([
        "JUMLAH",
        "",
        total_sib,
        total_bk,
        total_sib + total_bk
    ])

    # 3. Simpan file Excel ke dalam memori buffer (BytesIO) agar tidak mengotori storage
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    # 4. Alirkan file Excel langsung ke browser client
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=f"Laporan_Bulanan_{tahun}_{bulan}.xlsx"
    )

@app.route("/api/laporan-stats")
def api_laporan_stats():
    db = get_db()
    
    # 1. Ambil daftar tahun unik untuk dropdown filter
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_pasang) t FROM instalasi
               WHERE tanggal_pasang IS NOT NULL AND strftime('%Y', tanggal_pasang) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    tahun = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    bulan = request.args.get("bulan", "") or datetime.now().strftime("%m")

    # 2. Hitung rekap bulanan menggunakan fungsi helper bawaan
    rows, total_sib, total_bk = hitung_rekap_bulanan(db, tahun, bulan)
    nama_bulan = NAMA_BULAN[int(bulan) - 1]

    return jsonify({
        "tahun_list": tahun_list,
        "tahun_terpilih": tahun,
        "bulan_terpilih": bulan,
        "nama_bulan": nama_bulan,
        "nama_bulan_list": list(enumerate(NAMA_BULAN, start=1)),
        "rows": rows,
        "total_sib": total_sib,
        "total_bk": total_bk,
        "total_jumlah": total_sib + total_bk
    })


# ---------------- Eksekusi Aplikasi Utama ----------------
KETERANGAN_PERMOHONAN = [
    "Berkas / Diproses",
    "Permohonan Double",
    "Meter Air Sudah Terpasang",
    "Rumah Tidak Ditemukan",
    "Butuh Pipa Distribusi",
    "Tidak ada Akses Jalan Pipa",
    "Alamat tidak dapat ditemukan",
    "Rumah Tanah PJKA",
    "Bekas Pemutusan / Buka Kembali",
]


def panjang_pipa_label(jenis_pipa):
    """Panjang pipa itu bukan field terpisah -- diturunkan dari jenis_pipa.
    P.Dinas selalu <=10 meter, P.Distribusi selalu >10 meter (sesuai
    konfirmasi kabag)."""
    if jenis_pipa == "P.Dinas":
        return "\u226410 meter"
    if jenis_pipa == "P.Distribusi":
        return ">10 meter (P.Distribusi)"
    return None


def cari_atau_buat_pelanggan(db, nama, alamat, kelurahan, kecamatan):
    """Cocokin ke pelanggan existing (nama+ALAMAT sama persis, bukan cuma
    nama+kecamatan) -- biar gak salah gabung. Kasus nyata dari laporan
    Hublang: satu nama bisa muncul berkali-kali di kecamatan yang sama tapi
    alamat beda-beda (developer ngajuin permohonan buat banyak unit rumah
    sekaligus) -- itu HARUS dianggap pelanggan/properti berbeda, bukan satu
    pelanggan yang kebetulan namanya sama."""
    existing = db.execute(
        "SELECT id FROM pelanggan WHERE LOWER(nama)=LOWER(?) AND LOWER(alamat)=LOWER(?) AND LOWER(kecamatan)=LOWER(?)",
        (nama.strip(), alamat.strip(), kecamatan.strip()),
    ).fetchone()
    if existing:
        return existing["id"]
    cur = db.execute(
        "INSERT INTO pelanggan (nama, alamat, kelurahan, kecamatan) VALUES (?,?,?,?)",
        (nama, alamat, kelurahan, kecamatan),
    )
    return cur.lastrowid


# ---------------- Permohonan (dari Hublang) ----------------
@app.route("/permohonan")
def permohonan_list():
    db = get_db()
    jenis = request.args.get("jenis", "semua")
    status = request.args.get("status", "semua")  # semua/menunggu/ditindaklanjuti/selisih/selesai
    kecamatan = request.args.get("kecamatan", "semua")
    tahun = request.args.get("tahun", "").strip()
    bulan = request.args.get("bulan", "").strip()
    q = request.args.get("q", "").strip()
    page = max(1, request.args.get("page", 1, type=int))
    per_page = 30

    conditions = []
    params = []
    if jenis != "semua":
        conditions.append("jenis = ?")
        params.append(jenis)
    if kecamatan != "semua":
        conditions.append("kecamatan = ?")
        params.append(kecamatan)
    if status == "menunggu":
        conditions.append("ditindaklanjuti IS NULL")
    elif status == "ditindaklanjuti":
        conditions.append("ditindaklanjuti = 1 AND instalasi_id IS NULL")
    elif status == "selisih":
        conditions.append("ditindaklanjuti = 0")
    elif status == "selesai":
        conditions.append("instalasi_id IS NOT NULL")
    if tahun:
        conditions.append("strftime('%Y', tanggal_permohonan) = ?")
        params.append(tahun)
    if bulan:
        conditions.append("strftime('%m', tanggal_permohonan) = ?")
        params.append(bulan.zfill(2))
    if q:
        conditions.append("(nama_pelanggan LIKE ? OR lokasi LIKE ? OR kecamatan LIKE ?)")
        params += [f"%{q}%", f"%{q}%", f"%{q}%"]

    where_clause = " AND ".join(conditions) if conditions else "1=1"
    rows_all = db.execute(
        f"SELECT * FROM permohonan WHERE {where_clause} ORDER BY tanggal_permohonan DESC, id DESC",
        params,
    ).fetchall()

    total = len(rows_all)
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = min(page, total_pages)
    rows = rows_all[(page - 1) * per_page: page * per_page]

    ringkas = {
        "menunggu": db.execute("SELECT COUNT(*) c FROM permohonan WHERE ditindaklanjuti IS NULL").fetchone()["c"],
        "selisih": db.execute("SELECT COUNT(*) c FROM permohonan WHERE ditindaklanjuti = 0").fetchone()["c"],
        "selesai": db.execute("SELECT COUNT(*) c FROM permohonan WHERE instalasi_id IS NOT NULL").fetchone()["c"],
    }

    return render_template(
        "permohonan_list.html", rows=rows, jenis=jenis, status=status, kecamatan=kecamatan,
        tahun=tahun, bulan=bulan, q=q,
        page=page, total_pages=total_pages, total=total, ringkas=ringkas,
    )


@app.route("/permohonan/export")
def permohonan_export():
    db = get_db()
    jenis = request.args.get("jenis", "semua")
    status = request.args.get("status", "semua")
    kecamatan = request.args.get("kecamatan", "semua")
    tahun = request.args.get("tahun", "").strip()
    bulan = request.args.get("bulan", "").strip()
    q = request.args.get("q", "").strip()

    conditions = []
    params = []
    if jenis != "semua":
        conditions.append("jenis = ?")
        params.append(jenis)
    if kecamatan != "semua":
        conditions.append("kecamatan = ?")
        params.append(kecamatan)
    if status == "menunggu":
        conditions.append("ditindaklanjuti IS NULL")
    elif status == "ditindaklanjuti":
        conditions.append("ditindaklanjuti = 1 AND instalasi_id IS NULL")
    elif status == "selisih":
        conditions.append("ditindaklanjuti = 0")
    elif status == "selesai":
        conditions.append("instalasi_id IS NOT NULL")
    if tahun:
        conditions.append("strftime('%Y', tanggal_permohonan) = ?")
        params.append(tahun)
    if bulan:
        conditions.append("strftime('%m', tanggal_permohonan) = ?")
        params.append(bulan.zfill(2))
    if q:
        conditions.append("(nama_pelanggan LIKE ? OR lokasi LIKE ? OR kecamatan LIKE ?)")
        params += [f"%{q}%", f"%{q}%", f"%{q}%"]

    where_clause = " AND ".join(conditions) if conditions else "1=1"
    rows = db.execute(
        f"""SELECT nama_pelanggan, lokasi, kelurahan, kecamatan, jenis,
                          tanggal_permohonan, tanggal_survey, ditindaklanjuti,
                          jenis_pipa, tanggal_dikirim_hublang, tanggal_kembali_hublang,
                          petugas_survey, keterangan
                   FROM permohonan
                   WHERE {where_clause}
                   ORDER BY tanggal_permohonan DESC, id DESC""",
        params,
    ).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Permohonan"
    ws.append([
        "Nama Pelanggan", "Lokasi", "Kelurahan", "Kecamatan", "Jenis",
        "Tanggal Permohonan", "Tanggal Survey", "Ditindaklanjuti", "Jenis Pipa",
        "Tanggal Dikirim Hublang", "Tanggal Kembali Hublang", "Petugas Survey", "Keterangan"
    ])
    for r in rows:
        ws.append([
            r["nama_pelanggan"], r["lokasi"], r["kelurahan"], r["kecamatan"], r["jenis"],
            r["tanggal_permohonan"] or "", r["tanggal_survey"] or "",
            permohonan_ditindaklanjuti_label(r["ditindaklanjuti"]), r["jenis_pipa"] or "",
            r["tanggal_dikirim_hublang"] or "", r["tanggal_kembali_hublang"] or "",
            r["petugas_survey"] or "", r["keterangan"] or ""
        ])

    for idx, width in enumerate([25, 30, 18, 18, 10, 16, 16, 14, 18, 18, 18, 18, 30], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    filename = f"data_permohonan_{tahun or 'semua'}_{bulan or 'semua'}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/permohonan/kecamatan")
def permohonan_kecamatan():
    db = get_db()
    rows = db.execute(
        """SELECT kecamatan,
                  COUNT(*) total,
                  SUM(CASE WHEN jenis='SIB' THEN 1 ELSE 0 END) sib,
                  SUM(CASE WHEN jenis='BK' THEN 1 ELSE 0 END) bk,
                  SUM(CASE WHEN ditindaklanjuti IS NULL THEN 1 ELSE 0 END) menunggu,
                  SUM(CASE WHEN ditindaklanjuti = 0 THEN 1 ELSE 0 END) selisih,
                  SUM(CASE WHEN instalasi_id IS NOT NULL THEN 1 ELSE 0 END) selesai
           FROM permohonan GROUP BY kecamatan ORDER BY kecamatan"""
    ).fetchall()
    return render_template("permohonan_kecamatan.html", rows=rows)


@app.route("/permohonan/periode")
def permohonan_periode():
    db = get_db()
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_permohonan) t FROM permohonan
               WHERE tanggal_permohonan IS NOT NULL AND strftime('%Y', tanggal_permohonan) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    tahun_terpilih = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    if tahun_list and tahun_terpilih not in tahun_list:
        tahun_terpilih = tahun_list[0]

    per_bulan = db.execute(
        """SELECT strftime('%m', tanggal_permohonan) bulan,
                  COUNT(*) jumlah,
                  SUM(CASE WHEN jenis='SIB' THEN 1 ELSE 0 END) sib,
                  SUM(CASE WHEN jenis='BK' THEN 1 ELSE 0 END) bk
           FROM permohonan
           WHERE strftime('%Y', tanggal_permohonan) = ?
           GROUP BY bulan""",
        (tahun_terpilih,),
    ).fetchall()
    bulan_map = {r["bulan"]: dict(r) for r in per_bulan}
    ringkasan = [
        {
            "nomor": f"{i + 1:02d}", "nama": nama,
            "jumlah": bulan_map.get(f"{i + 1:02d}", {}).get("jumlah", 0),
            "sib": bulan_map.get(f"{i + 1:02d}", {}).get("sib", 0),
            "bk": bulan_map.get(f"{i + 1:02d}", {}).get("bk", 0),
        }
        for i, nama in enumerate(NAMA_BULAN)
    ]
    max_jumlah = max([r["jumlah"] for r in ringkasan], default=1) or 1

    return render_template(
        "permohonan_periode.html", tahun_list=tahun_list, tahun_terpilih=tahun_terpilih,
        ringkasan=ringkasan, max_jumlah=max_jumlah,
    )


@app.route("/permohonan/cari")
def permohonan_cari():
    db = get_db()
    q = request.args.get("q", "").strip()
    scope = request.args.get("scope", "semua")
    rows = []

    if q:
        query = "SELECT * FROM permohonan WHERE 1=1"
        params = []
        if scope == "nama":
            query += " AND nama_pelanggan LIKE ?"
            params.append(f"%{q}%")
        elif scope == "kecamatan":
            query += " AND kecamatan LIKE ?"
            params.append(f"%{q}%")
        else:
            query += " AND (nama_pelanggan LIKE ? OR kecamatan LIKE ? OR lokasi LIKE ?)"
            params += [f"%{q}%", f"%{q}%", f"%{q}%"]
        query += " ORDER BY tanggal_permohonan DESC"
        rows = db.execute(query, params).fetchall()

    return render_template("permohonan_cari.html", q=q, scope=scope, rows=rows)


def hitung_evaluasi_kinerja(db, tahun):
    """Rekap 12 bulan: berapa permohonan masuk, ditindaklanjuti, dan selisih,
    dipecah SIB/BK -- persis struktur "Evaluasi Kinerja Bagian Perencana &
    Supervisi" yang selama ini dibikin manual dari Excel. Dikelompokkan
    berdasarkan bulan tanggal_permohonan masuk (bukan bulan survey)."""
    baris = []
    total = {"masuk_sib": 0, "masuk_bk": 0, "lanjut_sib": 0, "lanjut_bk": 0, "selisih_sib": 0, "selisih_bk": 0}

    for i, nama_bulan in enumerate(NAMA_BULAN, start=1):
        bulan_str = f"{i:02d}"

        def hitung(jenis, kondisi_tambahan=""):
            q = f"""SELECT COUNT(*) c FROM permohonan
                    WHERE jenis=? AND strftime('%Y', tanggal_permohonan)=?
                          AND strftime('%m', tanggal_permohonan)=? {kondisi_tambahan}"""
            return db.execute(q, (jenis, tahun, bulan_str)).fetchone()["c"]

        masuk_sib = hitung("SIB")
        masuk_bk = hitung("BK")
        lanjut_sib = hitung("SIB", "AND ditindaklanjuti=1")
        lanjut_bk = hitung("BK", "AND ditindaklanjuti=1")
        selisih_sib = hitung("SIB", "AND ditindaklanjuti=0")
        selisih_bk = hitung("BK", "AND ditindaklanjuti=0")

        baris.append({
            "bulan": nama_bulan,
            "masuk_sib": masuk_sib, "masuk_bk": masuk_bk,
            "lanjut_sib": lanjut_sib, "lanjut_bk": lanjut_bk,
            "selisih_sib": selisih_sib, "selisih_bk": selisih_bk,
            "total_masuk": masuk_sib + masuk_bk,
            "total_lanjut": lanjut_sib + lanjut_bk,
            "total_selisih": selisih_sib + selisih_bk,
        })
        for k in total:
            total[k] += locals()[k]

    total["total_masuk"] = total["masuk_sib"] + total["masuk_bk"]
    total["total_lanjut"] = total["lanjut_sib"] + total["lanjut_bk"]
    total["total_selisih"] = total["selisih_sib"] + total["selisih_bk"]
    return baris, total


@app.route("/permohonan/laporan")
def permohonan_laporan():
    db = get_db()
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_permohonan) t FROM permohonan
               WHERE tanggal_permohonan IS NOT NULL AND strftime('%Y', tanggal_permohonan) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    tahun = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    baris, total = hitung_evaluasi_kinerja(db, tahun)
    return render_template("permohonan_laporan.html", tahun_list=tahun_list, tahun=tahun, baris=baris, total=total)


def ambil_rincian_bulanan(db, tahun, bulan, jenis):
    return db.execute(
        """SELECT nama_pelanggan, lokasi, kelurahan, kecamatan, tanggal_permohonan,
                  tanggal_survey, tanggal_kembali_hublang, petugas_survey, keterangan
           FROM permohonan
           WHERE jenis=? AND strftime('%Y', tanggal_permohonan)=? AND strftime('%m', tanggal_permohonan)=?
           ORDER BY tanggal_permohonan, id""",
        (jenis, tahun, bulan),
    ).fetchall()


def hitung_kinerja_teknik(db, tahun, bulan):
    """Rekap teknis per jenis (SIB/BK) untuk satu bulan: berapa masuk,
    disurvei, breakdown panjang pipa (diturunkan dari jenis_pipa), berapa
    dikirim ke Hublang, dan breakdown alasan khusus (butuh pipa distribusi,
    meter sudah terpasang, dst). JUMLAH di bagian keterangan itu sengaja
    HANYA jumlah dari 4 kategori khusus itu -- bukan sama dengan angka
    'dikirim ke Hublang' (yang jauh lebih besar, karena itu termasuk kasus
    normal yang gak butuh keterangan khusus)."""
    hasil = {}
    for jenis in ("SIB", "BK"):
        def hitung(kondisi_tambahan="", params_tambahan=()):
            q = f"""SELECT COUNT(*) c FROM permohonan
                    WHERE jenis=? AND strftime('%Y', tanggal_permohonan)=?
                          AND strftime('%m', tanggal_permohonan)=? {kondisi_tambahan}"""
            return db.execute(q, (jenis, tahun, bulan) + params_tambahan).fetchone()["c"]

        dari_hublang = hitung()
        disurvei = hitung("AND tanggal_survey IS NOT NULL")
        pipa_pendek = hitung("AND jenis_pipa='P.Dinas'")
        pipa_panjang = hitung("AND jenis_pipa='P.Distribusi'")
        dikirim = hitung("AND tanggal_dikirim_hublang IS NOT NULL")
        butuh_pipa = hitung("AND keterangan='Butuh Pipa Distribusi'")
        meter_terpasang = hitung("AND keterangan='Meter Air Sudah Terpasang'")
        bekas_pemutusan = hitung("AND keterangan='Bekas Pemutusan / Buka Kembali'")
        alamat_rumah_tt = hitung("AND keterangan IN ('Rumah Tidak Ditemukan','Alamat tidak dapat ditemukan')")

        hasil[jenis] = {
            "dari_hublang": dari_hublang, "ke_perencana": dari_hublang, "disurvei": disurvei,
            "pipa_pendek": pipa_pendek, "pipa_panjang": pipa_panjang, "dikirim": dikirim,
            "butuh_pipa": butuh_pipa, "meter_terpasang": meter_terpasang,
            "bekas_pemutusan": bekas_pemutusan, "alamat_rumah_tt": alamat_rumah_tt,
            "jumlah_ket": butuh_pipa + meter_terpasang + bekas_pemutusan + alamat_rumah_tt,
        }
    return hasil


@app.route("/permohonan/laporan/teknis")
def permohonan_laporan_teknis():
    db = get_db()
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_permohonan) t FROM permohonan
               WHERE tanggal_permohonan IS NOT NULL AND strftime('%Y', tanggal_permohonan) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    tahun = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    bulan = request.args.get("bulan", "") or datetime.now().strftime("%m")
    nama_bulan = NAMA_BULAN[int(bulan) - 1]

    hasil = hitung_kinerja_teknik(db, tahun, bulan)
    return render_template(
        "permohonan_laporan_teknis.html", tahun_list=tahun_list, tahun=tahun, bulan=bulan,
        nama_bulan=nama_bulan, nama_bulan_list=list(enumerate(NAMA_BULAN, start=1)), hasil=hasil,
    )


@app.route("/permohonan/laporan/teknis/unduh")
def permohonan_laporan_teknis_unduh():
    db = get_db()
    tahun = request.args.get("tahun", "")
    bulan = request.args.get("bulan", "")
    nama_bulan = NAMA_BULAN[int(bulan) - 1] if bulan else ""
    hasil = hitung_kinerja_teknik(db, tahun, bulan)

    wb = Workbook()
    ws = wb.active
    ws.title = "Kinerja Supervisi Teknik"
    bold = Font(bold=True)

    ws["A1"] = "LAPORAN KINERJA SUPERVISI TEKNIK/GAMBAR/PERHITUNGAN TEKNIS"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A2"] = f"BULAN: {nama_bulan.upper()} {tahun}"

    headers1 = ["Jenis Permohonan", "Dari Hublang", "Ke Perencana", "Disurvei",
                "<=10 Meter", ">10m (P.Distribusi)", "Dikirim ke Hublang",
                "Butuh Pipa Distribusi", "Meter Sudah Terpasang", "Bekas Pemutusan/Buka Kembali",
                "Alamat/Rumah Tidak Ditemukan", "Jumlah"]
    for col, h in enumerate(headers1, start=1):
        ws.cell(row=4, column=col, value=h).font = bold

    label_jenis = {"SIB": "Sambungan Instalasi Baru (SIB)", "BK": "Buka Kembali (BK)"}
    r = 5
    total = {k: 0 for k in ["dari_hublang", "ke_perencana", "disurvei", "pipa_pendek", "pipa_panjang",
                             "dikirim", "butuh_pipa", "meter_terpasang", "bekas_pemutusan",
                             "alamat_rumah_tt", "jumlah_ket"]}
    for jenis in ("SIB", "BK"):
        h = hasil[jenis]
        ws.cell(row=r, column=1, value=label_jenis[jenis])
        for col, key in enumerate(["dari_hublang", "ke_perencana", "disurvei", "pipa_pendek", "pipa_panjang",
                                    "dikirim", "butuh_pipa", "meter_terpasang", "bekas_pemutusan",
                                    "alamat_rumah_tt", "jumlah_ket"], start=2):
            ws.cell(row=r, column=col, value=h[key])
            total[key] += h[key]
        r += 1

    ws.cell(row=r, column=1, value="JUMLAH").font = bold
    for col, key in enumerate(["dari_hublang", "ke_perencana", "disurvei", "pipa_pendek", "pipa_panjang",
                                "dikirim", "butuh_pipa", "meter_terpasang", "bekas_pemutusan",
                                "alamat_rumah_tt", "jumlah_ket"], start=2):
        ws.cell(row=r, column=col, value=total[key]).font = bold

    for col, w in zip("ABCDEFGHIJKL", [26, 11, 11, 9, 9, 15, 13, 13, 14, 16, 17, 8]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name=f"laporan_kinerja_teknik_{tahun}-{bulan}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/permohonan/laporan/rincian")
def permohonan_laporan_rincian():
    db = get_db()
    tahun_list = [
        r["t"] for r in db.execute(
            """SELECT DISTINCT strftime('%Y', tanggal_permohonan) t FROM permohonan
               WHERE tanggal_permohonan IS NOT NULL AND strftime('%Y', tanggal_permohonan) IS NOT NULL
               ORDER BY t DESC"""
        ).fetchall()
    ]
    tahun = request.args.get("tahun", "") or (tahun_list[0] if tahun_list else datetime.now().strftime("%Y"))
    bulan = request.args.get("bulan", "") or datetime.now().strftime("%m")
    nama_bulan = NAMA_BULAN[int(bulan) - 1]

    sib_rows = ambil_rincian_bulanan(db, tahun, bulan, "SIB")
    bk_rows = ambil_rincian_bulanan(db, tahun, bulan, "BK")

    return render_template(
        "permohonan_laporan_rincian.html", tahun_list=tahun_list, tahun=tahun, bulan=bulan,
        nama_bulan=nama_bulan, nama_bulan_list=list(enumerate(NAMA_BULAN, start=1)),
        sib_rows=sib_rows, bk_rows=bk_rows,
    )


@app.route("/permohonan/laporan/rincian/unduh")
def permohonan_laporan_rincian_unduh():
    db = get_db()
    tahun = request.args.get("tahun", "")
    bulan = request.args.get("bulan", "")
    nama_bulan = NAMA_BULAN[int(bulan) - 1] if bulan else ""

    headers = ["No", "Nama Pelanggan", "Alamat", "Kelurahan", "Kecamatan",
               "Tgl Ke Perencana", "Tgl Survey", "Kembali Ke Hublang", "Petugas Survey", "Keterangan"]

    def tulis_sheet(ws, judul, rows):
        ws["A1"] = judul
        ws["A1"].font = Font(bold=True, size=12)
        for col, h in enumerate(headers, start=1):
            ws.cell(row=3, column=col, value=h).font = Font(bold=True)
        r = 4
        for i, row in enumerate(rows, start=1):
            ws.cell(row=r, column=1, value=i)
            ws.cell(row=r, column=2, value=row["nama_pelanggan"])
            ws.cell(row=r, column=3, value=row["lokasi"])
            ws.cell(row=r, column=4, value=row["kelurahan"])
            ws.cell(row=r, column=5, value=row["kecamatan"])
            ws.cell(row=r, column=6, value=row["tanggal_permohonan"])
            ws.cell(row=r, column=7, value=row["tanggal_survey"] or "")
            ws.cell(row=r, column=8, value=row["tanggal_kembali_hublang"] or "")
            ws.cell(row=r, column=9, value=row["petugas_survey"] or "")
            ws.cell(row=r, column=10, value=row["keterangan"] or "")
            r += 1
        widths = [5, 22, 26, 14, 14, 14, 12, 14, 14, 20]
        for col, w in zip("ABCDEFGHIJ", widths):
            ws.column_dimensions[col].width = w

    wb = Workbook()
    ws_sib = wb.active
    ws_sib.title = "Permohonan Masuk (SIB)"
    tulis_sheet(ws_sib, f"LAPORAN PERMOHONAN MASUK BULAN {nama_bulan.upper()} {tahun} DARI HUBUNGAN LANGGANAN",
                ambil_rincian_bulanan(db, tahun, bulan, "SIB"))

    ws_bk = wb.create_sheet("Buka Kembali (BK)")
    tulis_sheet(ws_bk, f"LAPORAN PERMOHONAN BUKA KEMBALI BULAN {nama_bulan.upper()} {tahun} DARI HUBUNGAN LANGGANAN",
                ambil_rincian_bulanan(db, tahun, bulan, "BK"))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name=f"laporan_permohonan_rincian_{tahun}-{bulan}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/permohonan/laporan/unduh")
def permohonan_laporan_unduh():
    db = get_db()
    tahun = request.args.get("tahun", "")
    baris, total = hitung_evaluasi_kinerja(db, tahun)

    wb = Workbook()
    ws = wb.active
    ws.title = "Evaluasi Kinerja"
    bold = Font(bold=True)

    ws["A1"] = "EVALUASI KINERJA BAGIAN PERENCANA & SUPERVISI"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"REKAP SURVEY PASANGAN INSTALASI BARU DAN BUKA KEMBALI - TAHUN {tahun}"

    headers1 = ["No", "Bulan", "Permohonan dari Hublang", "", "Ditindaklanjuti", "", "Selisih", "", "Total"]
    headers2 = ["", "", "SIB", "BK", "SIB", "BK", "SIB", "BK", ""]
    for col, h in enumerate(headers1, start=1):
        ws.cell(row=4, column=col, value=h).font = bold
    for col, h in enumerate(headers2, start=1):
        ws.cell(row=5, column=col, value=h).font = bold

    r = 6
    for i, b in enumerate(baris, start=1):
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=b["bulan"])
        ws.cell(row=r, column=3, value=b["masuk_sib"])
        ws.cell(row=r, column=4, value=b["masuk_bk"])
        ws.cell(row=r, column=5, value=b["lanjut_sib"])
        ws.cell(row=r, column=6, value=b["lanjut_bk"])
        ws.cell(row=r, column=7, value=b["selisih_sib"])
        ws.cell(row=r, column=8, value=b["selisih_bk"])
        ws.cell(row=r, column=9, value=b["total_masuk"])
        r += 1

    ws.cell(row=r, column=2, value="JUMLAH").font = bold
    for col, key in [(3, "masuk_sib"), (4, "masuk_bk"), (5, "lanjut_sib"), (6, "lanjut_bk"),
                     (7, "selisih_sib"), (8, "selisih_bk"), (9, "total_masuk")]:
        ws.cell(row=r, column=col, value=total[key]).font = bold

    for col, w in [("A", 5), ("B", 12), ("C", 8), ("D", 8), ("E", 8), ("F", 8), ("G", 8), ("H", 8), ("I", 10)]:
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True,
        download_name=f"evaluasi_kinerja_permohonan_{tahun}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.route("/permohonan/tambah", methods=["GET", "POST"])
def permohonan_tambah():
    db = get_db()
    if request.method == "POST":
        errors = validate_permohonan_data(request.form)
        if errors:
            flash("; ".join(errors), "error")
            return render_template("permohonan_form.html", permohonan=None, form_data=request.form,
                                    keterangan_pilihan=KETERANGAN_PERMOHONAN)

        db.execute(
            """INSERT INTO permohonan
               (nama_pelanggan, lokasi, kelurahan, kecamatan, jenis, tanggal_permohonan, keterangan)
               VALUES (?,?,?,?,?,?,?)""",
            (
                request.form["nama_pelanggan"],
                request.form.get("lokasi", ""),
                request.form.get("kelurahan", ""),
                request.form["kecamatan"],
                request.form["jenis"],
                request.form["tanggal_permohonan"],
                request.form.get("keterangan", ""),
            ),
        )
        db.commit()
        flash(f"Permohonan {request.form['nama_pelanggan']} tersimpan.")
        return redirect(url_for("permohonan_list"))

    return render_template("permohonan_form.html", permohonan=None, form_data=None,
                            keterangan_pilihan=KETERANGAN_PERMOHONAN)


@app.route("/permohonan/<int:pmid>/edit", methods=["GET", "POST"])
def permohonan_edit(pmid):
    db = get_db()
    permohonan = db.execute("SELECT * FROM permohonan WHERE id=?", (pmid,)).fetchone()
    if not permohonan:
        flash("Permohonan tidak ditemukan.", "error")
        return redirect(url_for("permohonan_list"))

    if request.method == "POST":
        errors = validate_permohonan_data(request.form)
        if errors:
            flash("; ".join(errors), "error")
            return render_template("permohonan_form.html", permohonan=permohonan, form_data=request.form,
                                    keterangan_pilihan=KETERANGAN_PERMOHONAN)

        ditindaklanjuti_raw = request.form.get("ditindaklanjuti", "")
        ditindaklanjuti = {"ya": 1, "tidak": 0}.get(ditindaklanjuti_raw)  # None kalau "belum"

        db.execute(
            """UPDATE permohonan SET
                 nama_pelanggan=?, lokasi=?, kelurahan=?, kecamatan=?, jenis=?, tanggal_permohonan=?,
                 tanggal_survey=?, petugas_survey=?, ditindaklanjuti=?, jenis_pipa=?,
                 tanggal_dikirim_hublang=?, tanggal_kembali_hublang=?, keterangan=?
               WHERE id=?""",
            (
                request.form["nama_pelanggan"], request.form.get("lokasi", ""),
                request.form.get("kelurahan", ""), request.form["kecamatan"], request.form["jenis"],
                request.form["tanggal_permohonan"], request.form.get("tanggal_survey", "") or None,
                request.form.get("petugas_survey") or None, ditindaklanjuti,
                request.form.get("jenis_pipa") or None, request.form.get("tanggal_dikirim_hublang", "") or None,
                request.form.get("tanggal_kembali_hublang", "") or None,
                request.form.get("keterangan", ""), pmid,
            ),
        )
        db.commit()
        flash("Permohonan diperbarui.")
        return redirect(url_for("permohonan_list"))

    return render_template("permohonan_form.html", permohonan=permohonan, form_data=None,
                            keterangan_pilihan=KETERANGAN_PERMOHONAN)


@app.route("/permohonan/<int:pmid>/jadikan-instalasi", methods=["GET", "POST"])
def permohonan_jadikan_instalasi(pmid):
    db = get_db()
    permohonan = db.execute("SELECT * FROM permohonan WHERE id=?", (pmid,)).fetchone()
    if not permohonan:
        flash("Permohonan tidak ditemukan.", "error")
        return redirect(url_for("permohonan_list"))

    if permohonan["instalasi_id"] is not None:
        flash("Permohonan ini sudah pernah dijadikan instalasi sebelumnya.", "error")
        return redirect(url_for("permohonan_list"))

    if request.method == "POST":
        nomor_instalasi = request.form["nomor_instalasi"].strip()
        dup = db.execute("SELECT 1 FROM instalasi WHERE nomor_instalasi=?", (nomor_instalasi,)).fetchone()
        if dup:
            flash(f"Nomor instalasi \"{nomor_instalasi}\" sudah dipakai instalasi lain.", "error")
            return render_template("permohonan_jadikan_instalasi.html", permohonan=permohonan, form_data=request.form)

        if not is_valid_iso_date(request.form.get("tanggal_pasang", "")):
            flash("Tanggal pasang tidak valid. Pakai date picker, jangan ketik manual.", "error")
            return render_template("permohonan_jadikan_instalasi.html", permohonan=permohonan, form_data=request.form)

        pelanggan_id = cari_atau_buat_pelanggan(
            db, permohonan["nama_pelanggan"], permohonan["lokasi"],
            permohonan["kelurahan"], permohonan["kecamatan"],
        )

        cur = db.execute(
            """INSERT INTO instalasi
               (pelanggan_id, nomor_instalasi, tanggal_pasang, diameter_pipa,
                tekanan_air, status, petugas, keterangan)
               VALUES (?,?,?,?,?,?,?,?)""",
            (
                pelanggan_id, nomor_instalasi, request.form["tanggal_pasang"],
                request.form.get("diameter_pipa", ""), request.form.get("tekanan_air") or None,
                permohonan["jenis"], request.form.get("petugas") or None,
                f"Dari permohonan #{pmid}",
            ),
        )
        instalasi_id = cur.lastrowid

        keterangan_baru = permohonan["keterangan"] or "Meter Air Sudah Terpasang"
        db.execute(
            "UPDATE permohonan SET instalasi_id=?, keterangan=? WHERE id=?",
            (instalasi_id, keterangan_baru, pmid),
        )
        db.commit()
        flash(f"Instalasi {nomor_instalasi} dibuat dari permohonan {permohonan['nama_pelanggan']}.")
        return redirect(url_for("permohonan_list"))

    return render_template("permohonan_jadikan_instalasi.html", permohonan=permohonan, form_data=None)


@app.route("/permohonan/<int:pmid>/hapus", methods=["POST"])
def permohonan_hapus(pmid):
    db = get_db()
    permohonan = db.execute("SELECT nama_pelanggan, instalasi_id FROM permohonan WHERE id=?", (pmid,)).fetchone()
    if not permohonan:
        flash("Permohonan tidak ditemukan.", "error")
        return redirect(url_for("permohonan_list"))

    db.execute("DELETE FROM permohonan WHERE id=?", (pmid,))
    db.commit()

    if permohonan["instalasi_id"]:
        flash(f"Permohonan {permohonan['nama_pelanggan']} dihapus. Instalasi yang sudah terlanjur dibuat dari permohonan ini TIDAK ikut terhapus.")
    else:
        flash(f"Permohonan {permohonan['nama_pelanggan']} dihapus.")
    return redirect(url_for("permohonan_list"))


if __name__ == "__main__":
    ensure_data_paths()

    # Otomatis buka browser saat file .exe dijalankan
    import webbrowser
    from threading import Timer

    def open_browser():
        webbrowser.open("http://127.0.0.1:5000")

    # Jalankan pembukaan browser dengan jeda 1.5 detik agar Flask siap mendengarkan port
    Timer(1.5, open_browser).start()

    # Matikan debug mode (debug=False) saat di-package ke exe agar tidak memicu reloading ganda
    # Mode ditentukan dari luar (run.bat), bukan hardcode di sini -- biar
    # app.py ini satu-satunya sumber kebenaran buat semua varian (local,
    # shared, Win7, Win10). Yang beda antar folder cukup run.bat-nya doang.
    sidap_mode = os.environ.get("SIDAP_MODE", "local")
    host = "0.0.0.0" if sidap_mode == "shared" else "127.0.0.1"
    print(f"SIDAP jalan dalam mode: {sidap_mode} (host={host})")

    app.run(debug=False, host=host, port=5000, threaded=True)