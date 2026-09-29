using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Query database untuk tampilan panel (tab Periode):
    /// daftar kiriman grup per periode, membuka snapshot, dan pengelolaan
    /// penyesuaian tersimpan (lihat, tambah, ubah, hapus, riwayat — Tahap 3d).
    /// </summary>
    internal static class DataPanel
    {
        // ================================================================
        // Kiriman grup per periode
        // ================================================================
        /// <param name="tanggal">"yyyy-MM-dd"; null/kosong = periode terbaru di database</param>
        public static Dictionary<string, object> DaftarStaging(string tanggal)
        {
            var hasil = new Dictionary<string, object> { { "adaDatabase", Database.Ada } };
            if (!Database.Ada) return hasil;

            using (var con = Database.Buka(false))
            {
                var periode = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT p.tanggal, p.status, COUNT(g.id) FROM periode p LEFT JOIN run_grup g ON g.periode_id=p.id AND g.aktif=1 " +
                    "GROUP BY p.id ORDER BY p.tanggal DESC"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        periode.Add(new Dictionary<string, object>
                        {
                            { "tanggal", rd.GetString(0) }, { "status", rd.GetString(1) }, { "jumlahGrup", Convert.ToInt32(rd[2]) }
                        });
                hasil["periode"] = periode;

                if (string.IsNullOrEmpty(tanggal) && periode.Count > 0)
                    tanggal = (string)((Dictionary<string, object>)periode[0])["tanggal"];
                hasil["tanggal"] = tanggal;
                if (string.IsNullOrEmpty(tanggal)) return hasil;

                var runs = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id, g.kode_kc, g.versi, g.aktif, g.status, g.pengguna, g.waktu, g.catatan, g.snapshot " +
                    "FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE p.tanggal=@p0 " +
                    "ORDER BY g.kode_kc, g.versi DESC", tanggal))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        runs.Add(new Dictionary<string, object>
                        {
                            { "id", rd.GetInt64(0) }, { "kodeKC", rd.GetString(1) }, { "versi", rd.GetInt32(2) },
                            { "aktif", rd.GetInt32(3) == 1 }, { "status", rd.GetString(4) },
                            { "pengguna", Str(rd[5]) }, { "waktu", Str(rd[6]) }, { "catatan", Str(rd[7]) },
                            { "adaSnapshot", !string.IsNullOrEmpty(Str(rd[8])) && File.Exists(Str(rd[8])) }
                        });

                // Nilai ringkasan utama per kiriman
                foreach (Dictionary<string, object> run in runs)
                {
                    var nilai = new Dictionary<string, object>();
                    using (var cmd = Database.Cmd(con, "SELECT kunci, nilai FROM ringkasan WHERE run_id=@p0", run["id"]))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) nilai[rd.GetString(0)] = Database.Dbl(rd[1]);
                    run["ringkasan"] = nilai;
                }
                hasil["runs"] = runs;
            }
            return hasil;
        }

        public static void BukaSnapshot(long runId)
        {
            using (var con = Database.Buka(false))
            {
                string path = Convert.ToString(Database.Scalar(con, "SELECT snapshot FROM run_grup WHERE id=@p0", runId) ?? "");
                if (path.Length == 0 || !File.Exists(path))
                    throw new FileNotFoundException("Snapshot tidak ditemukan.", path);
                // Hanya file di folder staging yang boleh dibuka dari panel
                if (!Path.GetFullPath(path).StartsWith(Path.GetFullPath(AppPaths.FolderStaging), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Lokasi snapshot di luar folder staging.");
                Process.Start(path);
            }
        }

        // ================================================================
        // Penyesuaian tersimpan — daftar untuk panel (Tahap 3d: + nama,
        // pemakaian terakhir, field lengkap untuk form edit)
        // ================================================================
        public static Dictionary<string, object> DaftarPenyesuaian()
        {
            var ind = new List<object>();
            var cs  = new List<object>();
            if (Database.Ada)
            {
                var pakaiInd = PemakaianTerakhir(
                    "SELECT h.no_kontrak, p.tanggal, g.kode_kc, h.nama, h.kc FROM hasil_individu h " +
                    "JOIN run_grup g ON g.id=h.run_id JOIN periode p ON p.id=g.periode_id " +
                    "WHERE g.aktif=1 AND g.dihapus=0 AND h.no_kontrak IN (SELECT no_kontrak FROM penyesuaian_individu) " +
                    "ORDER BY p.tanggal");
                var pakaiCs = PemakaianTerakhir(
                    "SELECT h.no_rek, p.tanggal, g.kode_kc, h.nama, '' FROM hasil_lgdcs h " +
                    "JOIN run_grup g ON g.id=h.run_id JOIN periode p ON p.id=g.periode_id " +
                    "WHERE g.aktif=1 AND g.dihapus=0 AND h.no_rek IN (SELECT no_rek FROM penyesuaian_lgdcs) " +
                    "ORDER BY p.tanggal");

                foreach (var p in Penyesuaian.MuatIndividu().Values)
                {
                    Dictionary<string, object> pk;
                    pakaiInd.TryGetValue(p.NoKontrak, out pk);
                    ind.Add(new Dictionary<string, object>
                    {
                        { "kunci", p.NoKontrak }, { "jaminan", p.Jaminan }, { "biaya", p.BiayaJual },
                        { "jaminanSistem", p.JaminanSistem }, { "alasan", p.Alasan },
                        { "pengguna", p.Pengguna }, { "waktu", p.Waktu }, { "periode", p.Periode },
                        { "nama", pk == null ? "" : pk["nama"] }, { "kc", pk == null ? "" : pk["kc"] },
                        { "dipakai", pk }
                    });
                }
                foreach (var p in Penyesuaian.MuatLgdCs().Values)
                {
                    Dictionary<string, object> pk;
                    pakaiCs.TryGetValue(p.NoRek, out pk);
                    cs.Add(new Dictionary<string, object>
                    {
                        { "kunci", p.NoRek }, { "jenis", p.Jenis },
                        { "nama", !string.IsNullOrEmpty(p.Nama) ? p.Nama : (pk == null ? "" : pk["nama"]) },
                        { "nilaiAgunan", p.NilaiAgunan }, { "recovery", p.Recovery }, { "pokok", p.PokokAwal },
                        { "thnSerah", p.ThnDiserahkan }, { "thnEks", p.ThnEksekusi },
                        { "alasan", p.Alasan }, { "pengguna", p.Pengguna }, { "waktu", p.Waktu }, { "periode", p.Periode },
                        { "dipakai", pk }
                    });
                }
            }
            string info;
            bool boleh = Database.BolehMenulis(out info);
            return new Dictionary<string, object>
            {
                { "individu", ind }, { "lgdcs", cs }, { "bolehMenulis", boleh }, { "infoPengirim", info }
            };
        }

        // kunci → { tanggal, kodeKC, nama, kc } dari kiriman aktif terakhir yang memuat kunci tsb
        private static Dictionary<string, Dictionary<string, object>> PemakaianTerakhir(string sql)
        {
            var hasil = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var con = Database.Buka(false))
                using (var cmd = Database.Cmd(con, sql))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                    {
                        if (rd[0] is DBNull) continue;
                        hasil[Convert.ToString(rd[0]).Trim()] = new Dictionary<string, object>
                        {
                            { "tanggal", Str(rd[1]) }, { "kodeKC", Str(rd[2]) }, { "nama", Str(rd[3]) }, { "kc", Str(rd[4]) }
                        };   // ORDER BY tanggal → baris terakhir = pemakaian terbaru
                    }
            }
            catch (Exception ex) { CatatanLog.Tulis("Pemakaian penyesuaian: " + ex.Message); }
            return hasil;
        }

        // ================================================================
        // Tambah / ubah penyesuaian dari panel (Tahap 3d)
        // Berlaku pada perhitungan BERIKUTNYA (tombol VBA maupun panel).
        // ================================================================
        public static Dictionary<string, object> SimpanPenyesuaian(string modul, Dictionary<string, object> d, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (d == null) throw new ArgumentException("Data penyesuaian kosong.");
            string kunci = Convert.ToString(Ambil(d, "kunci") ?? "").Trim();
            if (kunci.Length == 0)
                throw new InvalidOperationException(modul == "individu" ? "No. kontrak wajib diisi." : "No. rekening wajib diisi.");
            alasan = (alasan ?? "").Trim();

            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                string sebelum, sesudah, aksi;
                if (modul == "individu") SimpanIndividu(con, kunci, d, alasan, out sebelum, out sesudah, out aksi);
                else if (modul == "lgdcs") SimpanLgdCs(con, kunci, d, alasan, out sebelum, out sesudah, out aksi);
                else throw new ArgumentException("Modul tidak dikenal: " + modul);

                Database.Exec(con,
                    "INSERT INTO log_penyesuaian(waktu,pengguna,periode,modul,kunci,aksi,sebelum,sesudah,alasan) " +
                    "VALUES(@p0,@p1,'',@p2,@p3,@p4,@p5,@p6,@p7)",
                    Database.Sekarang(), Environment.UserName, modul, kunci, aksi, sebelum, sesudah, alasan);
                tx.Commit();
                CatatanLog.Tulis("Penyesuaian " + modul + " " + kunci + " " + aksi + ": " + sesudah);
                return new Dictionary<string, object> { { "aksi", aksi }, { "sesudah", sesudah } };
            }
        }

        private static void SimpanIndividu(System.Data.SQLite.SQLiteConnection con, string kunci, Dictionary<string, object> d,
                                           string alasan, out string sebelum, out string sesudah, out string aksi)
        {
            double? jaminan = Angka(Ambil(d, "jaminan"));
            double biaya = Angka(Ambil(d, "biaya")) ?? 0;
            if (jaminan == null || jaminan < 0) throw new InvalidOperationException("Nilai agunan wajib diisi (≥ 0).");
            if (biaya < 0) throw new InvalidOperationException("Biaya penjualan tidak boleh negatif.");

            double? jamSistem = null; string alasanLama = null; string periodeLama = "";
            sebelum = "";
            using (var cmd = Database.Cmd(con,
                "SELECT jaminan, biaya_jual, jaminan_sistem, alasan, periode FROM penyesuaian_individu WHERE no_kontrak=@p0", kunci))
            using (var rd = cmd.ExecuteReader())
                if (rd.Read())
                {
                    sebelum = "Agunan " + Rp(Database.Dbl(rd[0])) + " · Biaya " + Rp(Database.Dbl(rd[1]));
                    jamSistem = Database.Dbl(rd[2]); alasanLama = Str(rd[3]); periodeLama = Str(rd[4]);
                }
            aksi = sebelum.Length == 0 ? "tambah-dari-panel" : "ubah-dari-panel";

            // Kontrak baru: ambil nilai agunan sistem dari data dasar perhitungan terakhir di PC ini (bila ada)
            if (jamSistem == null) jamSistem = JaminanSistemDariDasar(kunci);

            sesudah = "Agunan " + Rp(jaminan) + " · Biaya " + Rp(biaya);
            Database.Exec(con,
                "INSERT OR REPLACE INTO penyesuaian_individu(no_kontrak,jaminan,biaya_jual,jaminan_sistem,alasan,pengguna,waktu,periode) " +
                "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7)",
                kunci, jaminan, biaya, jamSistem, alasan.Length > 0 ? alasan : (alasanLama ?? "Diubah dari panel"),
                Environment.UserName, Database.Sekarang(), periodeLama);
        }

        private static double? JaminanSistemDariDasar(string kunci)
        {
            try
            {
                var dasar = Penyesuaian.BacaDasar("individu.json");
                object baris;
                if (dasar == null || !dasar.TryGetValue("baris", out baris)) return null;
                foreach (var x in (System.Collections.IEnumerable)baris)
                {
                    var b = x as Dictionary<string, object>;
                    if (b != null && string.Equals(Convert.ToString(Ambil(b, "Kontrak")), kunci, StringComparison.OrdinalIgnoreCase))
                        return Angka(Ambil(b, "JaminanSistem"));
                }
            }
            catch { }
            return null;
        }

        private static void SimpanLgdCs(System.Data.SQLite.SQLiteConnection con, string kunci, Dictionary<string, object> d,
                                        string alasan, out string sebelum, out string sesudah, out string aksi)
        {
            string jenis = Convert.ToString(Ambil(d, "jenis") ?? "");
            if (jenis != "ubah" && jenis != "hapus" && jenis != "tambah")
                throw new InvalidOperationException("Jenis penyesuaian LGD CS harus ubah, hapus, atau tambah.");

            double? agunan = Angka(Ambil(d, "nilaiAgunan"));
            double? rec    = Angka(Ambil(d, "recovery"));
            double? pokok  = Angka(Ambil(d, "pokok"));
            string nama    = Convert.ToString(Ambil(d, "nama") ?? "").Trim();
            string thnS    = Convert.ToString(Ambil(d, "thnSerah") ?? "").Trim();
            string thnE    = Convert.ToString(Ambil(d, "thnEks") ?? "").Trim();

            if ((agunan ?? 0) < 0 || (rec ?? 0) < 0 || (pokok ?? 0) < 0)
                throw new InvalidOperationException("Nilai tidak boleh negatif.");
            foreach (var t in new[] { thnS, thnE })
            {
                int th;
                if (t.Length > 0 && (t.Length != 4 || !int.TryParse(t, out th)))
                    throw new InvalidOperationException("Tahun harus 4 digit (mis. 2024): " + t);
            }

            if (jenis == "ubah")
            {
                if (agunan == null && rec == null)
                    throw new InvalidOperationException("Isi nilai agunan dan/atau nilai realisasi untuk penyesuaian 'ubah'.");
                pokok = null; thnS = ""; thnE = "";
            }
            else if (jenis == "tambah")
            {
                if (pokok == null || pokok <= 0) throw new InvalidOperationException("Pokok awal wajib diisi untuk baris manual.");
            }
            else   // hapus = pengecualian
            {
                agunan = null; rec = null; thnS = ""; thnE = "";
            }

            string alasanLama = null, periodeLama = "";
            sebelum = "";
            using (var cmd = Database.Cmd(con,
                "SELECT jenis, nilai_agunan, recovery, pokok_awal, nama, alasan, periode FROM penyesuaian_lgdcs WHERE no_rek=@p0", kunci))
            using (var rd = cmd.ExecuteReader())
                if (rd.Read())
                {
                    sebelum = UraianCs(rd.GetString(0), Database.Dbl(rd[1]), Database.Dbl(rd[2]), Database.Dbl(rd[3]), Str(rd[4]));
                    if (nama.Length == 0) nama = Str(rd[4]);
                    alasanLama = Str(rd[5]); periodeLama = Str(rd[6]);
                }
            aksi = sebelum.Length == 0 ? "tambah-dari-panel" : "ubah-dari-panel";
            sesudah = UraianCs(jenis, agunan, rec, pokok, nama);

            Database.Exec(con,
                "INSERT OR REPLACE INTO penyesuaian_lgdcs(no_rek,jenis,nilai_agunan,recovery,nama,pokok_awal,thn_diserahkan,thn_eksekusi," +
                "alasan,pengguna,waktu,periode) VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11)",
                kunci, jenis, agunan, rec, nama, pokok, thnS, thnE,
                alasan.Length > 0 ? alasan : (alasanLama ?? "Diubah dari panel"),
                Environment.UserName, Database.Sekarang(), periodeLama);
        }

        private static string UraianCs(string jenis, double? agunan, double? rec, double? pokok, string nama)
        {
            string n = string.IsNullOrEmpty(nama) ? "" : nama + " · ";
            if (jenis == "hapus") return n + "Dikecualikan dari LGD CS";
            if (jenis == "tambah") return n + "Baris manual · pokok " + Rp(pokok) + " · agunan " + Rp(agunan) + " · realisasi " + Rp(rec);
            return n + "Diubah" + (agunan.HasValue ? " · agunan " + Rp(agunan) : "") + (rec.HasValue ? " · realisasi " + Rp(rec) : "");
        }

        /// <summary>Hapus satu penyesuaian. Alasan opsional sejak Tahap 3d.</summary>
        public static void HapusPenyesuaian(string modul, string kunci, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);

            string tabel = modul == "individu" ? "penyesuaian_individu" : modul == "lgdcs" ? "penyesuaian_lgdcs" : null;
            string kol   = modul == "individu" ? "no_kontrak" : "no_rek";
            if (tabel == null) throw new ArgumentException("Modul tidak dikenal: " + modul);
            alasan = string.IsNullOrWhiteSpace(alasan) ? "Dihapus dari panel" : alasan.Trim();

            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                int n = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM " + tabel + " WHERE " + kol + "=@p0", kunci));
                if (n == 0) throw new InvalidOperationException("Penyesuaian " + kunci + " tidak ditemukan.");
                Database.Exec(con, "DELETE FROM " + tabel + " WHERE " + kol + "=@p0", kunci);
                Database.Exec(con,
                    "INSERT INTO log_penyesuaian(waktu,pengguna,periode,modul,kunci,aksi,sebelum,sesudah,alasan) " +
                    "VALUES(@p0,@p1,'',@p2,@p3,'hapus-dari-panel','','Dihapus dari daftar penyesuaian',@p4)",
                    Database.Sekarang(), Environment.UserName, modul, kunci, alasan);
                tx.Commit();
            }
            CatatanLog.Tulis("Penyesuaian " + modul + " " + kunci + " dihapus: " + alasan);
        }

        /// <summary>Jejak perubahan satu kunci (maks. 50 terakhir).</summary>
        public static List<object> RiwayatPenyesuaian(string modul, string kunci)
        {
            var hasil = new List<object>();
            if (!Database.Ada) return hasil;
            using (var con = Database.Buka(false))
            using (var cmd = Database.Cmd(con,
                "SELECT waktu, pengguna, periode, aksi, sebelum, sesudah, alasan FROM log_penyesuaian " +
                "WHERE modul=@p0 AND kunci=@p1 ORDER BY id DESC LIMIT 50", modul, kunci))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    hasil.Add(new Dictionary<string, object>
                    {
                        { "waktu", Str(rd[0]) }, { "pengguna", Str(rd[1]) }, { "periode", Str(rd[2]) }, { "aksi", Str(rd[3]) },
                        { "sebelum", Str(rd[4]) }, { "sesudah", Str(rd[5]) }, { "alasan", Str(rd[6]) }
                    });
            return hasil;
        }

        // ---------------- util ----------------
        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        /// <summary>Angka dari JSON (number atau teks "1.234.567,5"); kosong → null.</summary>
        private static double? Angka(object v)
        {
            if (v == null) return null;
            if (v is string)
            {
                string t = ((string)v).Trim().Replace(" ", "");
                if (t.Length == 0) return null;
                t = t.Replace(".", "").Replace(",", ".");
                double x;
                if (double.TryParse(t, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out x)) return x;
                throw new InvalidOperationException("Angka tidak valid: " + v);
            }
            try { return Convert.ToDouble(v, System.Globalization.CultureInfo.InvariantCulture); }
            catch { throw new InvalidOperationException("Angka tidak valid: " + v); }
        }

        private static string Rp(double? v)
        {
            return v.HasValue ? v.Value.ToString("#,##0", System.Globalization.CultureInfo.GetCultureInfo("id-ID")) : "—";
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
