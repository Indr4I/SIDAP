DROP TABLE IF EXISTS instalasi;
DROP TABLE IF EXISTS pelanggan;

CREATE TABLE pelanggan (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama TEXT NOT NULL,
    alamat TEXT,
    kelurahan TEXT,
    kecamatan TEXT NOT NULL
);

-- Satu pelanggan bisa punya lebih dari satu instalasi (contoh nyata dari
-- buku besar: "Sarpita Santi Hasibuan" muncul 3x dengan nomor instalasi
-- berbeda-beda). Makanya instalasi dipisah jadi tabel sendiri, bukan
-- digabung flat ke tabel pelanggan.
CREATE TABLE instalasi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    pelanggan_id INTEGER NOT NULL,
    nomor_instalasi TEXT UNIQUE NOT NULL,   -- simpan sebagai teks, "83.236" bukan angka
    tanggal_pasang TEXT NOT NULL,           -- format YYYY-MM-DD
    diameter_pipa TEXT,                     -- contoh: "\u00d82", "\u00d83"
    tekanan_air REAL,
    status TEXT NOT NULL CHECK (status IN ('SIB', 'BK')),  -- Sambungan Instalasi Baru / Buka Kembali
    petugas TEXT,                           -- nullable, sering kosong di lapangan
    keterangan TEXT,
    FOREIGN KEY (pelanggan_id) REFERENCES pelanggan(id)
);

CREATE INDEX idx_pelanggan_kecamatan ON pelanggan(kecamatan);
CREATE INDEX idx_instalasi_tanggal ON instalasi(tanggal_pasang);
CREATE INDEX idx_instalasi_pelanggan ON instalasi(pelanggan_id);

-- Permohonan: data yang dilempar dari Hublang ke unit perencanaan.
-- Tahap sebelum jadi instalasi resmi. Satu tabel ini jadi sumber buat
-- beberapa laporan (rekap masuk, rekap survey, evaluasi kinerja bulanan)
-- -- bedanya cuma query/filter, bukan tabel terpisah-pisah.
DROP TABLE IF EXISTS permohonan;
CREATE TABLE permohonan (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nama_pelanggan TEXT NOT NULL,
    lokasi TEXT,                     -- alamat/lokasi seperti tercatat di laporan Hublang
    kelurahan TEXT,
    kecamatan TEXT NOT NULL,
    jenis TEXT NOT NULL CHECK (jenis IN ('SIB', 'BK')),
    tanggal_permohonan TEXT NOT NULL,   -- kapan masuk dari Hublang ke perencana
    tanggal_survey TEXT,                -- nullable, kosong sampai disurvei
    petugas_survey TEXT,
    ditindaklanjuti INTEGER,            -- NULL = belum disurvei, 1 = ya, 0 = tidak (selisih)
    jenis_pipa TEXT,                    -- 'P.Dinas' / 'P.Distribusi', informational
    tanggal_dikirim_hublang TEXT,       -- gambar/perhitungan teknis selesai, dikirim ke Hublang (kasus SUKSES/mayoritas)
    tanggal_kembali_hublang TEXT,       -- diisi kalau dikembalikan (kasus SELISIH/gagal, minoritas)
    keterangan TEXT,
    instalasi_id INTEGER,               -- nullable, keisi begitu "Jadikan Instalasi" dieksekusi
    FOREIGN KEY (instalasi_id) REFERENCES instalasi(id)
);

CREATE INDEX idx_permohonan_tanggal ON permohonan(tanggal_permohonan);
CREATE INDEX idx_permohonan_kecamatan ON permohonan(kecamatan);
CREATE INDEX idx_permohonan_jenis ON permohonan(jenis);
