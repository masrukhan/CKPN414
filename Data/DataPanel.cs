using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Query database untuk tampilan panel (tab Periode):
    /// daftar kiriman grup per periode, daftar penyesuaian tersimpan,
    /// hapus penyesuaian, dan membuka snapshot.
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
        // Penyesuaian tersimpan
        // ================================================================
        public static Dictionary<string, object> DaftarPenyesuaian()
        {
            var ind = new List<object>();
            var cs  = new List<object>();
            if (Database.Ada)
            {
                foreach (var p in Penyesuaian.MuatIndividu().Values)
                    ind.Add(new Dictionary<string, object>
                    {
                        { "kunci", p.NoKontrak }, { "jaminan", p.Jaminan }, { "biaya", p.BiayaJual },
                        { "jaminanSistem", p.JaminanSistem }, { "alasan", p.Alasan },
                        { "pengguna", p.Pengguna }, { "waktu", p.Waktu }, { "periode", p.Periode }
                    });
                foreach (var p in Penyesuaian.MuatLgdCs().Values)
                    cs.Add(new Dictionary<string, object>
                    {
                        { "kunci", p.NoRek }, { "jenis", p.Jenis }, { "nama", p.Nama },
                        { "nilaiAgunan", p.NilaiAgunan }, { "recovery", p.Recovery }, { "pokok", p.PokokAwal },
                        { "alasan", p.Alasan }, { "pengguna", p.Pengguna }, { "waktu", p.Waktu }, { "periode", p.Periode }
                    });
            }
            string info;
            bool boleh = Database.BolehMenulis(out info);
            return new Dictionary<string, object>
            {
                { "individu", ind }, { "lgdcs", cs }, { "bolehMenulis", boleh }, { "infoPengirim", info }
            };
        }

        /// <summary>Hapus satu penyesuaian (mis. pengecualian LGD CS yang ternyata keliru).</summary>
        public static void HapusPenyesuaian(string modul, string kunci, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan penghapusan wajib diisi.");

            string tabel = modul == "individu" ? "penyesuaian_individu" : modul == "lgdcs" ? "penyesuaian_lgdcs" : null;
            string kol   = modul == "individu" ? "no_kontrak" : "no_rek";
            if (tabel == null) throw new ArgumentException("Modul tidak dikenal: " + modul);

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
                    Database.Sekarang(), Environment.UserName, modul, kunci, alasan.Trim());
                tx.Commit();
            }
            CatatanLog.Tulis("Penyesuaian " + modul + " " + kunci + " dihapus: " + alasan);
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
