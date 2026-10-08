using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Parameter PD &amp; LGD Aset Antar Bank (ABA - KC0500) per tahun — Tahap 5h.
    ///
    /// Ditetapkan di panel (⚙ Pengaturan) dan disimpan di database (tabel parameter_aba). Bila ada
    /// untuk tahun Master!C4, nilainya MENIMPA sel Summary:
    ///   C21 = PD industri (asumsi LPS)
    ///   C22 = LGD porsi dijamin LPS
    ///   C23 = LGD porsi di atas plafon LPS
    /// sehingga C24:C26 (CKPN ABA) mengikuti parameter tahunan. Ditulis pada:
    ///   - awal setiap perhitungan panel (setelah pemeriksaan proteksi) dan sesudah langkah Refresh Summary;
    ///   - Hitung ulang Summary dan Simpan grup (sebelum angka Summary dibaca).
    /// Bila belum ditetapkan untuk tahun itu, sel Summary tidak disentuh (perilaku lama).
    /// Seperti Top-N, parameter tidak dapat diubah/dihapus bila sudah ada periode Final di tahun itu.
    /// </summary>
    internal static class ParameterAba
    {
        public const string SheetSummary = "Summary";
        public static readonly string[] Sel = { "C21", "C22", "C23" };

        internal class Nilai
        {
            public int Tahun;
            public double Pd, LgdDijamin, LgdAtas;
            public string Dasar, Pengguna, Waktu;

            public double[] Larik { get { return new[] { Pd, LgdDijamin, LgdAtas }; } }

            public string Ringkas()
            {
                return "PD " + Persen(Pd) + " · LGD dijamin " + Persen(LgdDijamin) + " · LGD di atas plafon " + Persen(LgdAtas);
            }
        }

        public static string Persen(double v) { return (v * 100).ToString("0.####", LogProses.Id) + "%"; }

        // =============================================================
        // Database
        // =============================================================
        public static Nilai Baca(int tahun)
        {
            if (!Database.Ada) return null;
            try
            {
                using (var con = Database.Buka(false)) return Baca(con, tahun);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca parameter ABA " + tahun + " gagal: " + ex.Message);
                return null;
            }
        }

        private static Nilai Baca(SQLiteConnection con, int tahun)
        {
            using (var cmd = Database.Cmd(con,
                "SELECT pd, lgd_dijamin, lgd_atas, dasar, pengguna, waktu FROM parameter_aba WHERE tahun=@p0", tahun))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) return null;
                return new Nilai
                {
                    Tahun = tahun, Pd = rd.GetDouble(0), LgdDijamin = rd.GetDouble(1), LgdAtas = rd.GetDouble(2),
                    Dasar = rd[3] is DBNull ? "" : rd.GetString(3),
                    Pengguna = rd[4] is DBNull ? "" : rd.GetString(4),
                    Waktu = rd[5] is DBNull ? "" : rd.GetString(5)
                };
            }
        }

        private static bool AdaFinal(SQLiteConnection con, int tahun)
        {
            return Convert.ToInt32(Database.Scalar(con,
                "SELECT COUNT(*) FROM periode WHERE status='Final' AND substr(tanggal,1,4)=@p0", tahun.ToString())) > 0;
        }

        private static void CekRasio(string label, double v)
        {
            if (double.IsNaN(v) || v < 0 || v > 1)
                throw new InvalidOperationException(label + " harus antara 0% dan 100%.");
        }

        /// <summary>Simpan parameter ABA tahun ini (rasio 0–1).</summary>
        public static Dictionary<string, object> Simpan(int tahun, double pd, double lgdDijamin, double lgdAtas, string dasar)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (tahun < 2000 || tahun > 2100) throw new InvalidOperationException("Tahun tidak valid.");
            CekRasio("PD industri", pd);
            CekRasio("LGD porsi dijamin", lgdDijamin);
            CekRasio("LGD porsi di atas plafon", lgdAtas);
            if (string.IsNullOrWhiteSpace(dasar)) throw new InvalidOperationException("Dasar penetapan (kebijakan / addendum / memo) wajib diisi.");

            var baru = new Nilai { Tahun = tahun, Pd = pd, LgdDijamin = lgdDijamin, LgdAtas = lgdAtas, Dasar = dasar.Trim() };
            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                var lama = Baca(con, tahun);
                bool berubah = lama == null || !SamaNilai(lama.Larik, baru.Larik);
                if (lama != null && berubah && AdaFinal(con, tahun))
                    throw new InvalidOperationException("Parameter ABA tahun " + tahun + " tidak dapat diubah karena sudah ada periode Final di tahun itu (" +
                        lama.Ringkas() + "). Buka kunci periode tersebut terlebih dahulu bila perubahan memang disetujui.");

                Database.Exec(con,
                    "INSERT OR REPLACE INTO parameter_aba(tahun,pd,lgd_dijamin,lgd_atas,dasar,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6)",
                    tahun, pd, lgdDijamin, lgdAtas, baru.Dasar, Environment.UserName, Database.Sekarang());
                Database.CatatAktivitas(con, tahun.ToString(), "parameter-aba",
                    (lama == null ? "Ditetapkan: " : "Diubah: " + lama.Ringkas() + " → ") + baru.Ringkas() + " · dasar: " + baru.Dasar);
                tx.Commit();
            }
            LogProses.CatatPanel(tahun + "-12-31", "Parameter ABA", "Tahun " + tahun, LogProses.OK, LogProses.R()
                .Tambah("PD", Persen(pd)).Tambah("LGD dijamin", Persen(lgdDijamin)).Tambah("LGD di atas plafon", Persen(lgdAtas))
                .Tambah("Dasar", baru.Dasar));
            return Info(tahun, null);
        }

        /// <summary>Hapus parameter tahun ini → Summary C21:C23 kembali diisi manual / oleh modul.</summary>
        public static Dictionary<string, object> Hapus(int tahun, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan wajib diisi.");
            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                var lama = Baca(con, tahun);
                if (lama == null) return Info(tahun, null);
                if (AdaFinal(con, tahun))
                    throw new InvalidOperationException("Parameter ABA tahun " + tahun + " tidak dapat dihapus karena sudah ada periode Final di tahun itu.");
                Database.Exec(con, "DELETE FROM parameter_aba WHERE tahun=@p0", tahun);
                Database.CatatAktivitas(con, tahun.ToString(), "parameter-aba-hapus", lama.Ringkas() + " · alasan: " + alasan.Trim());
                tx.Commit();
            }
            return Info(tahun, null);
        }

        // =============================================================
        // Panel: info untuk ⚙ Pengaturan
        // =============================================================
        /// <param name="app">bila tidak null, nilai Summary!C21:C23 saat ini ikut dibaca (periode Master)</param>
        public static Dictionary<string, object> Info(int tahun, Excel.Application app)
        {
            var h = new Dictionary<string, object> { { "tahun", tahun } };
            Nilai n = null;
            bool adaFinal = false;
            if (Database.Ada)
            {
                try
                {
                    using (var con = Database.Buka(false)) { n = Baca(con, tahun); adaFinal = AdaFinal(con, tahun); }
                }
                catch (Exception ex) { h["error"] = ex.Message; }
            }
            h["tersimpan"] = n == null ? null : new Dictionary<string, object>
            {
                { "pd", n.Pd }, { "lgdDijamin", n.LgdDijamin }, { "lgdAtas", n.LgdAtas },
                { "dasar", n.Dasar }, { "pengguna", n.Pengguna }, { "waktu", n.Waktu }
            };
            h["terkunci"] = adaFinal;
            string info;
            h["bolehMenulis"] = Database.BolehMenulis(out info);

            if (app != null)
            {
                try
                {
                    Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
                    Excel.Worksheet ws = wb == null ? null : ParameterMaster.CariSheet(wb, SheetSummary);
                    if (ws != null)
                    {
                        var v = BacaSel(ws);
                        h["sheet"] = new Dictionary<string, object> { { "pd", v[0] }, { "lgdDijamin", v[1] }, { "lgdAtas", v[2] } };
                        DateTime tgl;
                        if (ParameterMaster.BacaTanggalLaporan(wb, out tgl)) h["tahunMaster"] = tgl.Year;
                    }
                }
                catch (Exception ex) { h["errorSheet"] = ex.Message; }
            }
            return h;
        }

        // =============================================================
        // Excel: tulis ke Summary!C21:C23
        // =============================================================
        private static double?[] BacaSel(Excel.Worksheet ws)
        {
            var v = new double?[3];
            for (int i = 0; i < 3; i++)
            {
                object o = ((Excel.Range)ws.Range[Sel[i]]).Value2;
                v[i] = o is double ? (double?)(double)o : null;
            }
            return v;
        }

        private static bool SamaNilai(double[] a, double[] b)
        {
            for (int i = 0; i < 3; i++) if (Math.Abs(a[i] - b[i]) > 1e-12) return false;
            return true;
        }

        /// <summary>
        /// Terapkan parameter tahun Master!C4 ke Summary!C21:C23. Tidak melakukan apa pun bila parameter
        /// belum ditetapkan atau isi sel sudah sama. Mengembalikan keterangan perubahan (null = tidak ada).
        /// </summary>
        public static string Terapkan(Excel.Workbook wb)
        {
            if (wb == null) return null;
            DateTime tgl;
            if (!ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return null;
            Nilai n = Baca(tgl.Year);
            if (n == null) return null;
            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetSummary);
            if (ws == null) throw new InvalidOperationException("Sheet Summary tidak ditemukan.");

            var lama = BacaSel(ws);
            bool sama = true;
            for (int i = 0; i < 3; i++) if (!lama[i].HasValue || Math.Abs(lama[i].Value - n.Larik[i]) > 1e-12) sama = false;
            if (sama) return null;

            Protection.UbahSheet(wb, SheetSummary, s =>
            {
                for (int i = 0; i < 3; i++) ((Excel.Range)s.Range[Sel[i]]).Value2 = n.Larik[i];
            });
            try { ((Excel.Range)ws.Range["C20:C26"]).Calculate(); } catch { }

            string ket = "Parameter ABA " + tgl.Year + " diterapkan ke Summary C21:C23: " + n.Ringkas() +
                         " (sebelumnya " + TeksLama(lama) + ")";
            CatatanLog.Tulis("  " + ket);
            LogProses.Catat("Parameter ABA", "Summary C21:C23", LogProses.Info, LogProses.R()
                .Tambah("Tahun", tgl.Year)
                .Tambah("PD", Persen(n.Pd)).Tambah("LGD dijamin", Persen(n.LgdDijamin)).Tambah("LGD di atas plafon", Persen(n.LgdAtas))
                .Tambah("Sebelumnya", TeksLama(lama)));
            return ket;
        }

        /// <summary>Peringatan bila Summary!C21:C23 berbeda dari parameter tahun Master (untuk Periksa simpan).</summary>
        public static string CekBeda(Excel.Workbook wb)
        {
            try
            {
                DateTime tgl;
                if (wb == null || !ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return null;
                Nilai n = Baca(tgl.Year);
                Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetSummary);
                if (n == null || ws == null) return null;
                var v = BacaSel(ws);
                for (int i = 0; i < 3; i++)
                    if (!v[i].HasValue || Math.Abs(v[i].Value - n.Larik[i]) > 1e-12)
                        return "PD/LGD ABA di Summary C21:C23 (" + TeksLama(v) + ") berbeda dari parameter ABA tahun " + tgl.Year +
                               " (" + n.Ringkas() + "). Nilai parameter diterapkan otomatis saat Simpan.";
            }
            catch { }
            return null;
        }

        private static string TeksLama(double?[] v)
        {
            var s = new string[3];
            for (int i = 0; i < 3; i++) s[i] = v[i].HasValue ? Persen(v[i].Value) : "kosong";
            return string.Join(" / ", s);
        }
    }
}
