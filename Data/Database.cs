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
        public const int VersiSkema = 1;
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
                // Migrasi versi berikutnya (Tahap 3b dst.) ditambahkan di sini:
                // if (versi < 2) { ... }

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
