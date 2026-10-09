using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Akses database library\data\ckpn.db (SQLite).
    ///
    /// Pola pakai: satu PENGIRIM menulis, anggota tim lain hanya membaca.
    ///   - Penulis  : user yang terdaftar di config\pengirim.txt
    ///                (bila file kosong/tidak ada, semua user dianggap penulis).
    ///   - Pembaca  : koneksi dibuka "Read Only" → aman walau file di shared drive.
    ///
    /// Pengaturan untuk shared drive (SMB):
    ///   - journal_mode = DELETE (WAL tidak bekerja lewat jaringan).
    ///   - busy_timeout 10 detik: bila file sedang dipakai, tunggu sebentar
    ///     alih-alih langsung error "database is locked".
    ///
    /// Setiap penulisan besar (Simpan grup) didahului salinan cadangan
    /// ke library\backup\ (20 salinan terakhir disimpan).
    /// </summary>
    internal static class Database
    {
        public const int VersiSkema = 10;
        private static bool _skemaSiap;
        private static readonly object _kunci = new object();

        // ================================================================
        // Koneksi
        // ================================================================

        /// <summary>Buka koneksi. tulis=false → read-only.</summary>
        public static SQLiteConnection Buka(bool tulis)
        {
            NativeLoader.MuatSQLite();

            string path = AppPaths.FileDatabase;
            if (!tulis && !File.Exists(path))
                throw new FileNotFoundException("Database belum ada. Simpan grup pertama akan membuatnya.", path);

            // Skema lama (mis. dibuat Tahap 3a) diperbarui otomatis oleh PENGIRIM saat
            // pertama kali dibaca, supaya kolom/tabel baru tersedia untuk semua fitur baca.
            if (!tulis && !_skemaSiap)
            {
                string tidakPerlu;
                if (BolehMenulis(out tidakPerlu))
                    using (Buka(true)) { }
            }

            if (tulis) Directory.CreateDirectory(Path.GetDirectoryName(path));

            var sb = new SQLiteConnectionStringBuilder
            {
                // Path UNC (\\server\share\...) harus diawali 4 garis miring
                // agar tidak dipotong parser System.Data.SQLite.
                DataSource    = path.StartsWith(@"\\") ? @"\\" + path : path,
                Version       = 3,
                JournalMode   = SQLiteJournalModeEnum.Delete,
                ForeignKeys   = true,
                ReadOnly      = !tulis,
                FailIfMissing = !tulis
            };
            var con = new SQLiteConnection(sb.ConnectionString);
            con.Open();
            Exec(con, "PRAGMA busy_timeout = 10000");

            if (tulis) SiapkanSkema(con);
            return con;
        }

        /// <summary>True bila database sudah ada (untuk fitur baca).</summary>
        public static bool Ada { get { return File.Exists(AppPaths.FileDatabase); } }

        // ================================================================
        // Hak tulis
        // ================================================================

        /// <summary>Daftar pengirim dari config\pengirim.txt (kosong = semua boleh).</summary>
        public static List<string> DaftarPengirim()
        {
            var hasil = new List<string>();
            try
            {
                string f = AppPaths.FilePengirim;
                if (!File.Exists(f))
                {
                    // Buat template agar admin tahu formatnya.
                    Directory.CreateDirectory(Path.GetDirectoryName(f));
                    File.WriteAllText(f,
                        "# Daftar user Windows yang boleh menyimpan hasil ke database CKPN.\r\n" +
                        "# Satu nama per baris, contoh:\r\n" +
                        "# rukhan\r\n" +
                        "# Bila tidak ada nama sama sekali, semua user boleh menyimpan.\r\n");
                    return hasil;
                }
                foreach (var baris in File.ReadAllLines(f))
                {
                    string s = baris.Trim();
                    if (s.Length == 0 || s.StartsWith("#")) continue;
                    hasil.Add(s);
                }
            }
            catch { }
            return hasil;
        }

        public static bool BolehMenulis(out string alasan)
        {
            var daftar = DaftarPengirim();
            if (daftar.Count == 0)
            {
                alasan = "Daftar pengirim (library\\config\\pengirim.txt) masih kosong — semua user dapat menyimpan.";
                return true;
            }
            foreach (var u in daftar)
                if (string.Equals(u, Environment.UserName, StringComparison.OrdinalIgnoreCase))
                {
                    alasan = null;
                    return true;
                }
            alasan = "User '" + Environment.UserName + "' tidak terdaftar sebagai pengirim (library\\config\\pengirim.txt).";
            return false;
        }

        // ================================================================
        // Cadangan
        // ================================================================
        public static void Cadangkan()
        {
            try
            {
                string db = AppPaths.FileDatabase;
                if (!File.Exists(db)) return;
                Directory.CreateDirectory(AppPaths.FolderBackup);
                string tujuan = Path.Combine(AppPaths.FolderBackup,
                    "ckpn_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".db");
                File.Copy(db, tujuan, false);

                var lama = new List<string>(Directory.GetFiles(AppPaths.FolderBackup, "ckpn_*.db"));
                lama.Sort(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < lama.Count - 20; i++)
                    try { File.Delete(lama[i]); } catch { }
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Cadangan database gagal: " + ex.Message);
            }
        }

        /// <summary>
        /// Salinan arsip sebelum penghapusan data (Tahap 3d). Nama diawali "arsip_" sehingga
        /// tidak ikut dirotasi oleh Cadangkan(). Melempar error bila gagal — penghapusan
        /// tidak boleh berjalan tanpa arsip.
        /// </summary>
        public static string Arsipkan(string label)
        {
            string db = AppPaths.FileDatabase;
            if (!File.Exists(db)) throw new FileNotFoundException("Database belum ada.", db);
            Directory.CreateDirectory(AppPaths.FolderBackup);
            var aman = new System.Text.StringBuilder();
            foreach (char c in label ?? "")
                aman.Append(char.IsLetterOrDigit(c) || c == '-' ? c : '-');
            string tujuan = Path.Combine(AppPaths.FolderBackup,
                "arsip_" + aman + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".db");
            File.Copy(db, tujuan, false);
            return tujuan;
        }

        // ================================================================
        // Skema
        // ================================================================
        private static void SiapkanSkema(SQLiteConnection con)
        {
            lock (_kunci)
            {
                if (_skemaSiap) return;
                Exec(con, "CREATE TABLE IF NOT EXISTS meta (kunci TEXT PRIMARY KEY, nilai TEXT)");
                int versi = 0;
                using (var cmd = new SQLiteCommand("SELECT nilai FROM meta WHERE kunci='versi_skema'", con))
                {
                    object v = cmd.ExecuteScalar();
                    if (v != null) int.TryParse(Convert.ToString(v), out versi);
                }

                if (versi < 1)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV1) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','1')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: skema v1 dibuat di " + AppPaths.FileDatabase);
                }
                // ---- v2 (Tahap 3b): susunan grup, hapus versi, keputusan metode, kunci periode ----
                if (versi < 2)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV2) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','2')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v2 selesai");
                }
                // ---- v3 (Tahap 3c): metode konsolidasi per tahun + jurnal CKPN ----
                if (versi < 3)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV3) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','3')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v3 selesai");
                }
                // ---- v4 (Tahap 4): data analisis PD per kiriman (bahan sankey) ----
                if (versi < 4)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV4) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','4')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v4 selesai");
                }
                // ---- v5 (Tahap 4b): Top-N debitur CKPN Individu ditetapkan per tahun ----
                if (versi < 5)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV5) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','5')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v5 selesai");
                }
                // ---- v6 (Tahap 4c): Top-N per grup (menggantikan Top-N satu angka per tahun) ----
                if (versi < 6)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV6) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','6')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v6 selesai");
                }
                // ---- v7 (Tahap 4g): snapshot Overview Data per periode ----
                if (versi < 7)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV7) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','7')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v7 selesai");
                }
                // ---- v8 (Tahap 5h): parameter PD & LGD ABA per tahun (menimpa Summary C21:C23) ----
                if (versi < 8)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV8) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','8')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v8 selesai");
                }
                // ---- v9 (Tahap 6): dokumentasi — memo & bukti objektif penurunan nilai per kontrak ----
                if (versi < 9)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV9) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','9')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v9 selesai");
                }
                // ---- v10 (Tahap 6b): profil kontrak dari template APOLLO per periode ----
                if (versi < 10)
                {
                    using (var tx = con.BeginTransaction())
                    {
                        foreach (var sql in SkemaV10) Exec(con, sql);
                        Exec(con, "INSERT OR REPLACE INTO meta(kunci,nilai) VALUES('versi_skema','10')");
                        tx.Commit();
                    }
                    CatatanLog.Tulis("Database: migrasi ke skema v10 selesai");
                }
                // Migrasi versi berikutnya ditambahkan di sini: if (versi < 11) { ... }

                _skemaSiap = true;
            }
        }

        private static readonly string[] SkemaV1 =
        {
            // Periode laporan (tanggal posisi, mis. 2026-09-30)
            @"CREATE TABLE IF NOT EXISTS periode (
                id        INTEGER PRIMARY KEY,
                tanggal   TEXT NOT NULL UNIQUE,
                status    TEXT NOT NULL DEFAULT 'Terbuka',
                dibuat    TEXT NOT NULL)",

            // Satu baris per kiriman grup (kode KC yang dicentang saat dihitung).
            // Kirim ulang grup yang sama → versi naik, versi lama aktif=0.
            @"CREATE TABLE IF NOT EXISTS run_grup (
                id             INTEGER PRIMARY KEY,
                periode_id     INTEGER NOT NULL REFERENCES periode(id),
                kode_kc        TEXT NOT NULL,
                versi          INTEGER NOT NULL,
                aktif          INTEGER NOT NULL DEFAULT 1,
                status         TEXT NOT NULL,
                pengguna       TEXT,
                komputer       TEXT,
                waktu          TEXT NOT NULL,
                versi_addin    TEXT,
                parameter_json TEXT,
                catatan        TEXT,
                snapshot       TEXT,
                UNIQUE(periode_id, kode_kc, versi))",

            // Nilai ringkasan (Summary, LGD weighted, dll.) per kiriman
            @"CREATE TABLE IF NOT EXISTS ringkasan (
                run_id  INTEGER NOT NULL REFERENCES run_grup(id) ON DELETE CASCADE,
                kunci   TEXT NOT NULL,
                nilai   REAL,
                PRIMARY KEY (run_id, kunci))",

            // Isi akhir sheet A. CKPN - INDV per kiriman (setelah edit manual)
            @"CREATE TABLE IF NOT EXISTS hasil_individu (
                run_id          INTEGER NOT NULL REFERENCES run_grup(id) ON DELETE CASCADE,
                urut            INTEGER,
                kc              TEXT,
                cif             TEXT,
                nama            TEXT,
                no_kontrak      TEXT,
                os              REAL,
                ada_pn          TEXT,
                jaminan         REAL,
                biaya_jual      REAL,
                penurunan_nilai REAL,
                disesuaikan     INTEGER NOT NULL DEFAULT 0)",
            "CREATE INDEX IF NOT EXISTS ix_hasil_individu_kontrak ON hasil_individu(no_kontrak)",

            // Isi akhir sheet B4.LGD-CS MACET per kiriman
            @"CREATE TABLE IF NOT EXISTS hasil_lgdcs (
                run_id          INTEGER NOT NULL REFERENCES run_grup(id) ON DELETE CASCADE,
                no_rek          TEXT,
                nama            TEXT,
                pokok_awal      REAL,
                nilai_agunan    REAL,
                thn_diserahkan  TEXT,
                thn_eksekusi    TEXT,
                recovery        REAL,
                shortfall       REAL,
                sumber          TEXT)",
            "CREATE INDEX IF NOT EXISTS ix_hasil_lgdcs_rek ON hasil_lgdcs(no_rek)",

            // Penyesuaian manual CKPN Individu yang diterapkan ulang tiap perhitungan
            @"CREATE TABLE IF NOT EXISTS penyesuaian_individu (
                no_kontrak      TEXT PRIMARY KEY,
                jaminan         REAL,
                biaya_jual      REAL,
                jaminan_sistem  REAL,
                alasan          TEXT,
                pengguna        TEXT,
                waktu           TEXT,
                periode         TEXT)",

            // Penyesuaian manual LGD CS: ubah nilai / hapus (bukan eksekusi) / tambah baris
            @"CREATE TABLE IF NOT EXISTS penyesuaian_lgdcs (
                no_rek          TEXT PRIMARY KEY,
                jenis           TEXT NOT NULL CHECK (jenis IN ('ubah','hapus','tambah')),
                nilai_agunan    REAL,
                recovery        REAL,
                nama            TEXT,
                pokok_awal      REAL,
                thn_diserahkan  TEXT,
                thn_eksekusi    TEXT,
                alasan          TEXT,
                pengguna        TEXT,
                waktu           TEXT,
                periode         TEXT)",

            // Jejak audit setiap perubahan penyesuaian
            @"CREATE TABLE IF NOT EXISTS log_penyesuaian (
                id        INTEGER PRIMARY KEY,
                waktu     TEXT NOT NULL,
                pengguna  TEXT,
                periode   TEXT,
                modul     TEXT NOT NULL,
                kunci     TEXT NOT NULL,
                aksi      TEXT NOT NULL,
                sebelum   TEXT,
                sesudah   TEXT,
                alasan    TEXT)"
        };

        private static readonly string[] SkemaV2 =
        {
            // Susunan grup segmen per tahun (ditetapkan setahun sekali sesuai SOP)
            @"CREATE TABLE IF NOT EXISTS susunan_grup (
                id        INTEGER PRIMARY KEY,
                tahun     INTEGER NOT NULL,
                urut      INTEGER NOT NULL,
                nama      TEXT NOT NULL,
                kode_kc   TEXT NOT NULL,
                dasar     TEXT,
                pengguna  TEXT,
                waktu     TEXT,
                UNIQUE(tahun, kode_kc))",

            // Hapus versi kiriman = tandai (tidak dihapus fisik) supaya jejak audit tetap ada
            "ALTER TABLE run_grup ADD COLUMN dihapus INTEGER NOT NULL DEFAULT 0",
            "ALTER TABLE run_grup ADD COLUMN alasan_hapus TEXT",
            "ALTER TABLE run_grup ADD COLUMN pengguna_hapus TEXT",
            "ALTER TABLE run_grup ADD COLUMN waktu_hapus TEXT",

            // Kunci periode setelah konsolidasi
            "ALTER TABLE periode ADD COLUMN dikunci_oleh TEXT",
            "ALTER TABLE periode ADD COLUMN dikunci_waktu TEXT",
            "ALTER TABLE periode ADD COLUMN catatan_kunci TEXT",

            // Metode terpilih per grup saat konsolidasi vs PPKA
            @"CREATE TABLE IF NOT EXISTS keputusan_metode (
                periode_id  INTEGER NOT NULL REFERENCES periode(id),
                kode_kc     TEXT NOT NULL,
                nama_grup   TEXT,
                run_id      INTEGER REFERENCES run_grup(id),
                metode      TEXT NOT NULL CHECK (metode IN ('nf','mig')),
                ckpn_nf     REAL,
                ckpn_mig    REAL,
                ppka        REAL,
                pengguna    TEXT,
                waktu       TEXT,
                PRIMARY KEY (periode_id, kode_kc))",

            // Jejak aktivitas periode (hapus versi, kunci, buka kunci, susunan)
            @"CREATE TABLE IF NOT EXISTS log_aktivitas (
                id        INTEGER PRIMARY KEY,
                waktu     TEXT NOT NULL,
                pengguna  TEXT,
                periode   TEXT,
                aksi      TEXT NOT NULL,
                detail    TEXT)"
        };

        private static readonly string[] SkemaV3 =
        {
            // Pengaturan konsolidasi per tahun, ditetapkan bersama susunan grup:
            //   metode         : 'nf' | 'mig' — SATU metode untuk semua grup sepanjang tahun
            //   kebijakan_saldo: 'ckpn' (saldo CKPN = hasil metode) | 'maks' (saldo = maks(CKPN, PPKA))
            @"CREATE TABLE IF NOT EXISTS susunan_tahun (
                tahun            INTEGER PRIMARY KEY,
                metode           TEXT CHECK (metode IN ('nf','mig')),
                kebijakan_saldo  TEXT NOT NULL DEFAULT 'ckpn' CHECK (kebijakan_saldo IN ('ckpn','maks')),
                dasar            TEXT,
                pengguna         TEXT,
                waktu            TEXT)",

            // Ringkasan konsolidasi yang ditetapkan saat periode dikunci
            "ALTER TABLE periode ADD COLUMN metode TEXT",
            "ALTER TABLE periode ADD COLUMN kebijakan_saldo TEXT",
            "ALTER TABLE periode ADD COLUMN ckpn_total REAL",
            "ALTER TABLE periode ADD COLUMN ppka_total REAL",

            // Usulan jurnal CKPN per komponen (pembiayaan / ABA) saat periode dikunci
            @"CREATE TABLE IF NOT EXISTS jurnal_ckpn (
                periode_id    INTEGER NOT NULL REFERENCES periode(id),
                komponen      TEXT NOT NULL,
                jenis         TEXT NOT NULL,
                ckpn          REAL,
                ppka          REAL,
                saldo_target  REAL,
                saldo_awal    REAL,
                nominal       REAL,
                akun_debit    TEXT,
                akun_kredit   TEXT,
                keterangan    TEXT,
                pengguna      TEXT,
                waktu         TEXT,
                PRIMARY KEY (periode_id, komponen))"
        };

        private static readonly string[] SkemaV4 =
        {
            // Saldo bucket net flow, matriks migrasi per triwulan, dan PD resmi per kiriman (JSON)
            @"CREATE TABLE IF NOT EXISTS analisis_pd (
                run_id  INTEGER NOT NULL REFERENCES run_grup(id) ON DELETE CASCADE,
                jenis   TEXT NOT NULL,
                data    TEXT NOT NULL,
                PRIMARY KEY (run_id, jenis))"
        };

        private static readonly string[] SkemaV5 =
        {
            // Top-N debitur CKPN Individu sesuai SOP, ditetapkan bersama susunan grup & metode
            "ALTER TABLE susunan_tahun ADD COLUMN top_n INTEGER",
            // Top-N yang dipakai saat kiriman grup disimpan (jejak audit)
            "ALTER TABLE run_grup ADD COLUMN top_n INTEGER"
        };

        private static readonly string[] SkemaV6 =
        {
            // Top-N debitur CKPN Individu per grup (jumlah debitur tiap grup berbeda)
            "ALTER TABLE susunan_grup ADD COLUMN top_n INTEGER",
            // Nilai tahunan dari Tahap 4b (bila sudah diisi) menjadi nilai awal setiap grup tahun itu
            "UPDATE susunan_grup SET top_n=(SELECT t.top_n FROM susunan_tahun t WHERE t.tahun=susunan_grup.tahun) WHERE top_n IS NULL"
        };

        private static readonly string[] SkemaV7 =
        {
            // OS, EAD, PPKA per KC per kualitas + info keuangan (JSON dari DataOverviewBuilder)
            @"CREATE TABLE IF NOT EXISTS overview_periode (
                tanggal   TEXT PRIMARY KEY,
                data      TEXT NOT NULL,
                sumber    TEXT,
                pengguna  TEXT,
                waktu     TEXT)"
        };

        private static readonly string[] SkemaV8 =
        {
            // Rasio 0–1. pd = Summary!C21, lgd_dijamin = C22, lgd_atas = C23
            @"CREATE TABLE IF NOT EXISTS parameter_aba (
                tahun        INTEGER PRIMARY KEY,
                pd           REAL NOT NULL CHECK (pd BETWEEN 0 AND 1),
                lgd_dijamin  REAL NOT NULL CHECK (lgd_dijamin BETWEEN 0 AND 1),
                lgd_atas     REAL NOT NULL CHECK (lgd_atas BETWEEN 0 AND 1),
                dasar        TEXT,
                pengguna     TEXT,
                waktu        TEXT)"
        };

        private static readonly string[] SkemaV9 =
        {
            // Bukti objektif penurunan nilai per kontrak (dicatat modul CKPN Individu, disimpan saat Simpan grup)
            "ALTER TABLE hasil_individu ADD COLUMN kualitas INTEGER",
            "ALTER TABLE hasil_individu ADD COLUMN hari_tunggakan REAL",
            "ALTER TABLE hasil_individu ADD COLUMN restruktur INTEGER",
            "ALTER TABLE hasil_individu ADD COLUMN dasar_pn TEXT",

            // Dokumen / memo per periode:
            //   jenis 'ringkasan' (kunci '')      — catatan & penandatangan ringkasan bulanan
            //   jenis 'individu'  (kunci = CIF)   — memo penilaian penurunan nilai individual
            //   jenis 'lgdcs'     (kunci = no rek) — memo penetapan nilai realisasi agunan LGD CS
            //   data = JSON isian petugas; status 'Draf' | 'Final'
            @"CREATE TABLE IF NOT EXISTS memo (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                jenis         TEXT NOT NULL CHECK (jenis IN ('ringkasan','individu','lgdcs')),
                tanggal       TEXT NOT NULL,
                kunci         TEXT NOT NULL DEFAULT '',
                data          TEXT NOT NULL,
                status        TEXT NOT NULL DEFAULT 'Draf' CHECK (status IN ('Draf','Final')),
                dibuat_oleh   TEXT,
                dibuat_waktu  TEXT,
                diubah_oleh   TEXT,
                diubah_waktu  TEXT,
                final_oleh    TEXT,
                final_waktu   TEXT,
                UNIQUE (jenis, tanggal, kunci))",
            "CREATE INDEX IF NOT EXISTS ix_memo_kunci ON memo(jenis, kunci, tanggal)"
        };

        private static readonly string[] SkemaV10 =
        {
            // Profil per rekening dari template APOLLO (KC0600–KC1100, GB0500, KC2900) — JSON, lihat ProfilTemplate.cs
            @"CREATE TABLE IF NOT EXISTS profil_template (
                tanggal  TEXT NOT NULL,
                no_rek   TEXT NOT NULL,
                data     TEXT NOT NULL,
                waktu    TEXT,
                PRIMARY KEY (tanggal, no_rek))",
            @"CREATE TABLE IF NOT EXISTS profil_template_info (
                tanggal   TEXT PRIMARY KEY,
                file      TEXT,
                cap       TEXT,
                pengguna  TEXT,
                waktu     TEXT)"
        };

        public static void CatatAktivitas(SQLiteConnection con, string periode, string aksi, string detail)
        {
            Exec(con, "INSERT INTO log_aktivitas(waktu,pengguna,periode,aksi,detail) VALUES(@p0,@p1,@p2,@p3,@p4)",
                 Sekarang(), Environment.UserName, periode ?? "", aksi, detail ?? "");
        }

        // ================================================================
        // Util
        // ================================================================
        public static void Exec(SQLiteConnection con, string sql, params object[] p)
        {
            using (var cmd = Cmd(con, sql, p)) cmd.ExecuteNonQuery();
        }

        public static object Scalar(SQLiteConnection con, string sql, params object[] p)
        {
            using (var cmd = Cmd(con, sql, p)) return cmd.ExecuteScalar();
        }

        /// <summary>Parameter posisional: @p0, @p1, … sesuai urutan argumen.</summary>
        public static SQLiteCommand Cmd(SQLiteConnection con, string sql, params object[] p)
        {
            var cmd = new SQLiteCommand(sql, con);
            if (p != null)
                for (int i = 0; i < p.Length; i++)
                    cmd.Parameters.AddWithValue("@p" + i, p[i] ?? DBNull.Value);
            return cmd;
        }

        public static string Sekarang() { return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); }

        public static double? Dbl(object v)
        {
            if (v == null || v is DBNull) return null;
            return Convert.ToDouble(v);
        }
    }
}
