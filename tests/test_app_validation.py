import io
import os
import sqlite3
import tempfile
import unittest

from openpyxl import Workbook

import app


class AppValidationTests(unittest.TestCase):
    def test_parse_tanggal_accepts_multiple_formats(self):
        expected = "2026-02-15"
        for value in ["2026-02-15", "15-02-2026", "15/02/2026", "15-02-26"]:
            formatted, error = app.parse_tanggal(value)
            self.assertIsNone(error)
            self.assertEqual(formatted, expected)

    def test_parse_tanggal_rejects_invalid_date(self):
        formatted, error = app.parse_tanggal("31-02-2026")
        self.assertIsNone(formatted)
        self.assertIsNotNone(error)
        self.assertIn("Format tanggal", error)

    def test_parse_tanggal_rejects_unreasonable_year(self):
        formatted, error = app.parse_tanggal("15-02-1899")
        self.assertIsNone(formatted)
        self.assertIn("Tahun tanggal", error)

    def test_validate_permohonan_data_returns_no_errors_for_valid_input(self):
        data = {
            "nama_pelanggan": "Budi",
            "kecamatan": "Setiabudi",
            "jenis": "SIB",
            "tanggal_permohonan": "2026-03-01",
            "tanggal_survey": "2026-03-02",
            "tanggal_dikirim_hublang": "2026-03-05",
            "tanggal_kembali_hublang": "2026-03-10",
        }
        errors = app.validate_permohonan_data(data)
        self.assertEqual(errors, [])

    def test_validate_permohonan_data_returns_errors_for_invalid_input(self):
        data = {
            "nama_pelanggan": "",
            "kecamatan": "",
            "jenis": "BAD",
            "tanggal_permohonan": "2026-13-01",
            "tanggal_survey": "2026/03/02",
        }
        errors = app.validate_permohonan_data(data)
        self.assertGreaterEqual(len(errors), 3)
        self.assertIn("Nama pelanggan kosong", errors)
        self.assertIn("Kecamatan kosong", errors)
        self.assertIn("Jenis permohonan tidak valid", errors)

    def test_db_has_schema_detects_required_tables(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = os.path.join(tmpdir, "test_sidap.db")
            conn = sqlite3.connect(db_path)
            conn.execute("CREATE TABLE pelanggan (id INTEGER PRIMARY KEY)")
            conn.execute("CREATE TABLE instalasi (id INTEGER PRIMARY KEY)")
            conn.execute("CREATE TABLE permohonan (id INTEGER PRIMARY KEY)")
            conn.commit()
            conn.close()

            original_db_path = app.DB_PATH
            try:
                app.DB_PATH = db_path
                self.assertTrue(app.db_has_schema())
            finally:
                app.DB_PATH = original_db_path

    def test_db_has_schema_returns_false_for_missing_tables(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = os.path.join(tmpdir, "test_sidap.db")
            conn = sqlite3.connect(db_path)
            conn.execute("CREATE TABLE pelanggan (id INTEGER PRIMARY KEY)")
            conn.commit()
            conn.close()

            original_db_path = app.DB_PATH
            try:
                app.DB_PATH = db_path
                self.assertFalse(app.db_has_schema())
            finally:
                app.DB_PATH = original_db_path

    def test_cari_atau_buat_pelanggan_inserts_or_reuses(self):
        conn = sqlite3.connect(":memory:")
        conn.row_factory = sqlite3.Row
        conn.execute("CREATE TABLE pelanggan (id INTEGER PRIMARY KEY AUTOINCREMENT, nama TEXT, alamat TEXT, kelurahan TEXT, kecamatan TEXT)")
        conn.commit()

        first_id = app.cari_atau_buat_pelanggan(conn, "Adi", "Jl. Malioboro", "Kotabaru", "Yogyakarta")
        self.assertIsInstance(first_id, int)

        second_id = app.cari_atau_buat_pelanggan(conn, "Adi", "Jl. Malioboro", "Kotabaru", "Yogyakarta")
        self.assertEqual(first_id, second_id)

        third_id = app.cari_atau_buat_pelanggan(conn, "Adi", "Jl. Other", "Kotabaru", "Yogyakarta")
        self.assertNotEqual(first_id, third_id)

        conn.close()

    def test_parse_permohonan_upload_accepts_valid_rows(self):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "No", "Nama Pelanggan", "Lokasi", "Kelurahan", "Kecamatan",
            "Jenis", "Tanggal Permohonan", "Tanggal Survey", "Petugas Survey",
            "Ditindaklanjuti", "Jenis Pipa", "Tanggal Dikirim Hublang",
            "Tanggal Kembali Hublang", "Keterangan",
        ])
        ws.append([
            1, "Budi", "Jl. Contoh", "Kelurahan A", "Kecamatan B",
            "SIB", "2026-03-01", "2026-03-02", "Anto",
            "Belum", "P.Dinas", "2026-03-05", "2026-03-10", "Catatan",
        ])

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        class DummyDB:
            def execute(self, *args, **kwargs):
                return []

        rows = app.parse_permohonan_upload(buffer, DummyDB())
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["nama_pelanggan"], "Budi")
        self.assertEqual(rows[0]["jenis"], "SIB")
        self.assertEqual(rows[0]["tanggal_permohonan"], "2026-03-01")
        self.assertEqual(rows[0]["ditindaklanjuti"], None)
        self.assertEqual(rows[0]["errors"], [])

    def test_parse_permohonan_upload_detects_errors(self):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "No", "Nama Pelanggan", "Lokasi", "Kelurahan", "Kecamatan",
            "Jenis", "Tanggal Permohonan", "Tanggal Survey", "Petugas Survey",
            "Ditindaklanjuti", "Jenis Pipa", "Tanggal Dikirim Hublang",
            "Tanggal Kembali Hublang", "Keterangan",
        ])
        ws.append([
            1, "", "", "", "", "BAD", "31-02-2026", "2026-03-02", "Anto",
            "Maybe", "WrongType", "2026-03-05", "2026-03-10", "Catatan",
        ])

        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        class DummyDB:
            def execute(self, *args, **kwargs):
                return []

        rows = app.parse_permohonan_upload(buffer, DummyDB())
        self.assertEqual(len(rows), 1)
        self.assertNotEqual(rows[0]["errors"], [])
        self.assertIn("Nama pelanggan kosong", rows[0]["errors"])
        self.assertTrue(any("Jenis 'BAD' tidak valid" in e for e in rows[0]["errors"]))
        self.assertTrue(any("Format tanggal \"31-02-2026\" tidak dikenali" in e for e in rows[0]["errors"]))
        self.assertTrue(any("Ditindaklanjuti 'Maybe' tidak valid" in e for e in rows[0]["errors"]))


if __name__ == "__main__":
    unittest.main()
