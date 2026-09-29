using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Data analisis PD (Tahap 4) — bahan sankey PD Migration & PD Net Flow di tab Analisis.
    ///
    /// Sumber sel (dibaca apa adanya, tidak dihitung ulang):
    ///   B1.PD-Net Flow      D3:O3   label bulan (12 kolom)
    ///                       D4:O16  saldo bucket 0, 1-30, ..., 331-360 hari
    ///                       D18:O18 saldo > 360 hari, D19:O19 hapus buku (KC2900)
    ///   B2.PD-Migration     4 blok triwulan (baris awal 4 / 23 / 42 / 61), 5 baris kualitas:
    ///                       B = saldo awal, C..G = ke Kol 1..5, H = hapus buku, I = lainnya/keluar
    ///   B. CKPN - KOL INDV  E6:E19  PD Net Flow per bucket (14 bucket, sama dengan Dashboard lama)
    ///                       D38:D42 PD Migration per kualitas, E38 LGD weighted
    ///
    /// Data disimpan per kiriman grup (tabel analisis_pd, JSON) saat Simpan grup, sehingga
    /// grafik periode lama tetap bisa dibuka tanpa workbook. Kiriman sebelum Tahap 4 tidak
    /// punya data ini — panel menampilkannya dari workbook yang sedang terbuka saja.
    /// </summary>
    internal static class AnalisisPD
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static readonly string[] LabelBucket =
        {
            "0 hari", "1-30", "31-60", "61-90", "91-120", "121-150", "151-180",
            "181-210", "211-240", "241-270", "271-300", "301-330", "331-360", "> 360"
        };
        private static readonly int[] BarisTriwulan = { 4, 23, 42, 61 };
        private static readonly string[] NamaTriwulan = { "Triwulan I", "Triwulan II", "Triwulan III", "Triwulan IV" };

        // =================================================================
        // Baca dari workbook aplikasi yang sedang terbuka
        // =================================================================
        public static Dictionary<string, object> BacaWorkbook(Excel.Workbook wb)
        {
            var hasil = new Dictionary<string, object> { { "versiData", 1 } };

            // ---- PD Net Flow: saldo bucket per bulan ----
            Excel.Worksheet nf = ParameterMaster.CariSheet(wb, "B1.PD-Net Flow");
            if (nf != null)
            {
                object[,] v = (object[,])((Excel.Range)nf.Range["D3", "O19"]).Value2;   // 17 baris x 12 kolom
                var bulan = new List<string>();
                var saldo = new List<double[]>();
                for (int b = 0; b < 14; b++) saldo.Add(new double[12]);
                var wo = new double[12];
                for (int m = 0; m < 12; m++)
                {
                    bulan.Add(LabelBulan(v[1, m + 1], m));
                    for (int b = 0; b < 13; b++) saldo[b][m] = Angka(v[2 + b, m + 1]);   // baris 4..16
                    saldo[13][m] = Angka(v[16, m + 1]);                                  // baris 18
                    wo[m] = Angka(v[17, m + 1]);                                         // baris 19
                }
                hasil["netflow"] = new Dictionary<string, object>
                {
                    { "bucket", LabelBucket }, { "bulan", bulan }, { "saldo", saldo }, { "wo", wo }
                };
            }

            // ---- PD Migration: matriks per triwulan ----
            Excel.Worksheet mg = ParameterMaster.CariSheet(wb, "B2.PD-Migration");
            if (mg != null)
            {
                object[,] v = (object[,])((Excel.Range)mg.Range["B4", "I65"]).Value2;   // baris 4..65, kolom B..I
                var tw = new List<object>();
                for (int t = 0; t < 4; t++)
                {
                    var saldoAwal = new double[5];
                    var matriks = new List<double[]>();
                    double total = 0;
                    for (int k = 0; k < 5; k++)
                    {
                        int r = BarisTriwulan[t] - 4 + k + 1;   // indeks 1-based di array
                        saldoAwal[k] = Angka(v[r, 1]);
                        total += saldoAwal[k];
                        var baris = new double[7];               // Kol1..Kol5, hapus buku, lainnya
                        for (int j = 0; j < 7; j++) baris[j] = Angka(v[r, 2 + j]);
                        matriks.Add(baris);
                    }
                    tw.Add(new Dictionary<string, object>
                    {
                        { "nama", NamaTriwulan[t] }, { "ada", total > 0 }, { "saldoAwal", saldoAwal }, { "matriks", matriks }
                    });
                }
                hasil["migrasi"] = new Dictionary<string, object> { { "triwulan", tw } };
            }

            // ---- PD resmi dari sheet kolektif ----
            Excel.Worksheet kol = ParameterMaster.CariSheet(wb, "B. CKPN - KOL INDV");
            if (kol != null)
            {
                object[,] pdNf = (object[,])((Excel.Range)kol.Range["E6", "E19"]).Value2;
                object[,] pdMg = (object[,])((Excel.Range)kol.Range["D38", "D42"]).Value2;
                var a = new double?[14];
                for (int i = 0; i < 14; i++) a[i] = AngkaNull(pdNf[i + 1, 1]);
                var b = new double?[5];
                for (int i = 0; i < 5; i++) b[i] = AngkaNull(pdMg[i + 1, 1]);
                hasil["pdNetFlow"] = a;
                hasil["pdMigrasi"] = b;
                hasil["lgdWeighted"] = AngkaNull(((Excel.Range)kol.Range["E38"]).Value2);
            }
            return hasil;
        }

        /// <summary>Versi aman untuk Simpan grup: kegagalan baca tidak membatalkan penyimpanan.</summary>
        public static string BacaWorkbookJson(Excel.Workbook wb)
        {
            try { return _json.Serialize(BacaWorkbook(wb)); }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Analisis PD tidak tersimpan: " + ex.Message);
                return null;
            }
        }

        public static void SimpanRun(SQLiteConnection con, long runId, string json)
        {
            if (string.IsNullOrEmpty(json)) return;
            Database.Exec(con, "INSERT OR REPLACE INTO analisis_pd(run_id, jenis, data) VALUES(@p0,'pd',@p1)", runId, json);
        }

        // =================================================================
        // Untuk panel: daftar sumber & data per sumber
        // =================================================================

        /// <summary>
        /// Daftar pilihan sumber data analisis: kiriman aktif per periode (dengan nama grup
        /// menurut susunan tahunnya) dan opsi gabungan semua grup per periode.
        /// </summary>
        public static Dictionary<string, object> DaftarSumber()
        {
            var periode = new List<object>();
            if (Database.Ada)
                using (var con = Database.Buka(false))
                {
                    var daftar = new List<string>();
                    using (var cmd = Database.Cmd(con, "SELECT tanggal FROM periode ORDER BY tanggal DESC"))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) daftar.Add(rd.GetString(0));

                    foreach (string tgl in daftar)
                    {
                        int sumber;
                        var susunan = Periode.Susunan(con, int.Parse(tgl.Substring(0, 4)), out sumber);
                        var runs = new List<object>();
                        using (var cmd = Database.Cmd(con,
                            "SELECT g.id, g.kode_kc, g.versi, (SELECT COUNT(*) FROM analisis_pd a WHERE a.run_id=g.id) " +
                            "FROM run_grup g JOIN periode p ON p.id=g.periode_id " +
                            "WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0 ORDER BY g.kode_kc", tgl))
                        using (var rd = cmd.ExecuteReader())
                            while (rd.Read())
                            {
                                string kode = rd.GetString(1);
                                var g = susunan == null ? null : susunan.Find(x => x.KodeKC == kode);
                                runs.Add(new Dictionary<string, object>
                                {
                                    { "runId", rd.GetInt64(0) }, { "kodeKC", kode }, { "versi", rd.GetInt32(2) },
                                    { "nama", g == null ? kode : g.Nama }, { "adaData", Convert.ToInt32(rd[3]) > 0 }
                                });
                            }
                        if (runs.Count > 0) periode.Add(new Dictionary<string, object> { { "tanggal", tgl }, { "runs", runs } });
                    }
                }
            return new Dictionary<string, object> { { "periode", periode } };
        }

        public static Dictionary<string, object> BacaRun(long runId)
        {
            using (var con = Database.Buka(false))
            {
                object data = Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis='pd'", runId);
                if (data == null || data is DBNull)
                    throw new InvalidOperationException("Kiriman ini disimpan sebelum Tahap 4 sehingga belum punya data analisis PD. " +
                        "Buka grupnya (Buka untuk diedit) lalu simpan ulang, atau pilih sumber Workbook saat ini.");
                var hasil = _json.Deserialize<Dictionary<string, object>>(Convert.ToString(data));
                using (var cmd = Database.Cmd(con,
                    "SELECT p.tanggal, g.kode_kc, g.versi FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0", runId))
                using (var rd = cmd.ExecuteReader())
                    if (rd.Read())
                    {
                        hasil["tanggal"] = rd.GetString(0);
                        hasil["kodeKC"] = rd.GetString(1);
                        hasil["versi"] = rd.GetInt32(2);
                    }
                return hasil;
            }
        }

        /// <summary>
        /// Gabungan semua kiriman aktif satu periode: saldo net flow dan matriks migrasi
        /// dijumlahkan (nilai rupiah bersifat aditif). PD tidak dijumlahkan — PD berbeda per grup.
        /// Label bulan net flow diambil dari kiriman pertama.
        /// </summary>
        public static Dictionary<string, object> BacaGabungan(string tanggal)
        {
            var ids = new List<long>();
            var tanpaData = new List<string>();
            using (var con = Database.Buka(false))
            using (var cmd = Database.Cmd(con,
                "SELECT g.id, g.kode_kc, (SELECT COUNT(*) FROM analisis_pd a WHERE a.run_id=g.id) FROM run_grup g " +
                "JOIN periode p ON p.id=g.periode_id WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0", tanggal))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                {
                    if (Convert.ToInt32(rd[2]) > 0) ids.Add(rd.GetInt64(0));
                    else tanpaData.Add(rd.GetString(1));
                }
            if (ids.Count == 0) throw new InvalidOperationException("Belum ada kiriman periode " + tanggal + " yang memuat data analisis PD.");

            double[][] saldo = null; double[] wo = null; object bulan = null;
            double[][][] mat = new double[4][][]; double[][] awal = new double[4][];
            var kode = new List<string>();
            foreach (long id in ids)
            {
                var d = BacaRun(id);
                kode.Add(Convert.ToString(d["kodeKC"]));
                var nf = d.ContainsKey("netflow") ? d["netflow"] as Dictionary<string, object> : null;
                if (nf != null)
                {
                    if (bulan == null) bulan = nf["bulan"];
                    var s = KeMatriks(nf["saldo"]);
                    var w = KeVektor(nf["wo"]);
                    if (saldo == null) { saldo = s; wo = w; }
                    else { Tambah(saldo, s); Tambah(wo, w); }
                }
                var mg = d.ContainsKey("migrasi") ? d["migrasi"] as Dictionary<string, object> : null;
                if (mg != null)
                {
                    int t = 0;
                    foreach (var o in (IEnumerable)mg["triwulan"])
                    {
                        var tw = (Dictionary<string, object>)o;
                        var m = KeMatriks(tw["matriks"]); var a = KeVektor(tw["saldoAwal"]);
                        if (mat[t] == null) { mat[t] = m; awal[t] = a; }
                        else { Tambah(mat[t], m); Tambah(awal[t], a); }
                        if (++t == 4) break;
                    }
                }
            }

            var hasil = new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "kodeKC", string.Join(" + ", kode.ToArray()) }, { "gabungan", true },
                { "tanpaData", tanpaData }
            };
            if (saldo != null)
                hasil["netflow"] = new Dictionary<string, object> { { "bucket", LabelBucket }, { "bulan", bulan }, { "saldo", saldo }, { "wo", wo } };
            var twList = new List<object>();
            for (int t = 0; t < 4; t++)
                if (mat[t] != null)
                {
                    double total = 0; foreach (var x in awal[t]) total += x;
                    twList.Add(new Dictionary<string, object>
                    {
                        { "nama", NamaTriwulan[t] }, { "ada", total > 0 }, { "saldoAwal", awal[t] }, { "matriks", mat[t] }
                    });
                }
            if (twList.Count > 0) hasil["migrasi"] = new Dictionary<string, object> { { "triwulan", twList } };
            return hasil;
        }

        // =================================================================
        // Util
        // =================================================================
        private static string LabelBulan(object v, int m)
        {
            if (v is double)
            {
                double d = (double)v;
                if (d > 20000 && d < 80000)   // tanggal Excel (OADate) → "Sep 2026"
                    return DateTime.FromOADate(d).ToString("MMM yyyy", new CultureInfo("id-ID"));
                return d.ToString(CultureInfo.InvariantCulture);
            }
            string s = v == null ? "" : Convert.ToString(v).Trim();
            return s.Length > 0 ? s : "Bulan " + (m + 1);
        }

        private static double Angka(object v) { return v is double ? (double)v : 0; }
        private static double? AngkaNull(object v) { return v is double ? (double?)v : null; }

        private static double[] KeVektor(object o)
        {
            var hasil = new List<double>();
            foreach (var x in (IEnumerable)o) hasil.Add(x == null ? 0 : Convert.ToDouble(x, CultureInfo.InvariantCulture));
            return hasil.ToArray();
        }

        private static double[][] KeMatriks(object o)
        {
            var hasil = new List<double[]>();
            foreach (var baris in (IEnumerable)o) hasil.Add(KeVektor(baris));
            return hasil.ToArray();
        }

        private static void Tambah(double[] a, double[] b)
        {
            for (int i = 0; i < Math.Min(a.Length, b.Length); i++) a[i] += b[i];
        }

        private static void Tambah(double[][] a, double[][] b)
        {
            for (int i = 0; i < Math.Min(a.Length, b.Length); i++) Tambah(a[i], b[i]);
        }
    }
}
