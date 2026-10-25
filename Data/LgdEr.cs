using System;
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
    /// Tahap 9d — rincian LGD Expected Recoveries per kiriman grup.
    ///
    /// Sheet "B3.LGD-ER" dibangun ulang modul LGDExpectedRecoveries setiap hitung, dengan pola blok per
    /// kohort tahun hapus buku (lihat BangunCohortBlock):
    ///   C = "Tahun hapus buku YYYY"  (E = tahun kohort)          ← baris judul blok
    ///   F..K = tahun pengamatan                                   ← baris tahun
    ///   C = "Nama Nasabah"                                        ← baris label
    ///   pasangan baris per rekening:
    ///     saldo    : C = no. rekening, D = tahun HB, E = balance hapus buku, F..K = saldo per tahun
    ///     recovery : C = "-", E = "recovery", F..K = recovery per tahun, N = jumlah, O = recovery rate
    ///   C = "combine"  : E, F..K, N, O total kohort
    ///   C = "TOTAL COMBINE" : E = total WO, N = total recovery, O = rate
    ///   "Total WO" / "Total Recovery" / "LGD Weighted" di kolom C, nilainya di D
    ///
    /// Sheet dibaca apa adanya (nilai, bukan rumus) dengan memindai kolom C, sehingga tidak bergantung
    /// pada jumlah baris. Disimpan sebagai JSON di analisis_pd (jenis 'lgder', tanpa perubahan skema)
    /// dan ikut tersalin saat koreksi panel (KoreksiRun menyalin semua jenis analisis_pd).
    /// </summary>
    internal static class LgdEr
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        public const string Jenis = "lgder";
        private const string SheetEr = "B3.LGD-ER";

        public static string BacaWorkbookJson(Excel.Workbook wb)
        {
            try
            {
                var d = BacaWorkbook(wb);
                return d == null ? null : _json.Serialize(d);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Rincian LGD ER tidak tersimpan: " + ex.Message);
                return null;
            }
        }

        public static Dictionary<string, object> BacaWorkbook(Excel.Workbook wb)
        {
            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetEr);
            if (ws == null) return null;
            Excel.Range akhirC = (Excel.Range)((Excel.Range)ws.Cells[ws.Rows.Count, "C"]).End[Excel.XlDirection.xlUp];
            int last = Math.Max(akhirC.Row, 10);
            object[,] v = (object[,])((Excel.Range)ws.Range["C1", "O" + last]).Value2;   // kolom C..O → indeks 1..13
            const int C = 1, D = 2, E = 3, F = 4, N = 12, O = 13;

            var kohort = new List<object>();
            Dictionary<string, object> blok = null, rek = null;
            List<object> daftarRek = null;
            var hasil = new Dictionary<string, object> { { "versiData", 1 }, { "kohort", kohort } };

            for (int r = 1; r <= last; r++)
            {
                string c = Teks(v[r, C]);
                if (c.StartsWith("Tahun hapus buku", StringComparison.OrdinalIgnoreCase) && !c.Equals("Tahun hapus buku berjalan", StringComparison.OrdinalIgnoreCase))
                {
                    daftarRek = new List<object>();
                    blok = new Dictionary<string, object> { { "tahun", Bulat(v[r, E]) ?? TahunDariTeks(c) }, { "rekening", daftarRek } };
                    // baris berikutnya: tahun pengamatan F..K
                    if (r + 1 <= last) blok["tahunKolom"] = Deret(v, r + 1, F, 6, true);
                    kohort.Add(blok);
                    rek = null;
                    continue;
                }
                if (blok == null)
                {
                    continue;
                }
                if (c.Equals("Nama Nasabah", StringComparison.OrdinalIgnoreCase)) continue;
                if (c.Equals("combine", StringComparison.OrdinalIgnoreCase))
                {
                    blok["total"] = new Dictionary<string, object>
                    {
                        { "balance", Angka(v[r, E]) }, { "saldo", Deret(v, r, F, 6, false) }, { "recovery", Angka(v[r, N]) }, { "rate", Angka(v[r, O]) }
                    };
                    rek = null;
                    continue;
                }
                if (c.Equals("TOTAL COMBINE", StringComparison.OrdinalIgnoreCase))
                {
                    hasil["total"] = new Dictionary<string, object> { { "wo", Angka(v[r, E]) }, { "recovery", Angka(v[r, N]) }, { "rate", Angka(v[r, O]) } };
                    blok = null;
                    continue;
                }
                if (c == "-")
                {
                    // baris recovery milik rekening di atasnya
                    if (rek != null)
                    {
                        rek["recovery"] = Deret(v, r, F, 6, false);
                        rek["jumlahRecovery"] = Angka(v[r, N]);
                        rek["rate"] = Angka(v[r, O]);
                    }
                    continue;
                }
                if (c.Length == 0 || c.StartsWith("(tidak ada data", StringComparison.OrdinalIgnoreCase)) { rek = null; continue; }
                // baris saldo rekening
                rek = new Dictionary<string, object>
                {
                    { "rek", c }, { "tahunHB", Bulat(v[r, D]) }, { "balance", Angka(v[r, E]) }, { "saldo", Deret(v, r, F, 6, false) }
                };
                daftarRek.Add(rek);
            }

            // Ringkasan LGD Weighted (label di C, nilai di D)
            for (int r = 1; r <= last; r++)
            {
                string c = Teks(v[r, C]);
                if (c == "Total WO") hasil["totalWo"] = Angka(v[r, D]);
                else if (c == "Total Recovery") hasil["totalRecovery"] = Angka(v[r, D]);
                else if (c == "LGD Weighted") hasil["lgd"] = Angka(v[r, D]);
            }
            return kohort.Count == 0 ? null : hasil;
        }

        public static void SimpanRun(SQLiteConnection con, long runId, string json)
        {
            if (string.IsNullOrEmpty(json)) return;
            Database.Exec(con, "INSERT OR REPLACE INTO analisis_pd(run_id, jenis, data) VALUES(@p0,@p1,@p2)", runId, Jenis, json);
        }

        /// <summary>Mode acuan (PD & LGD Desember): LGD ER milik versi acuan disalin, bukan isi sheet yang tidak dihitung ulang.</summary>
        public static void SalinDariRun(SQLiteConnection con, long runIdBaru, long runIdAcuan)
        {
            Database.Exec(con, "INSERT OR REPLACE INTO analisis_pd(run_id, jenis, data) " +
                "SELECT @p0, jenis, data FROM analisis_pd WHERE run_id=@p1 AND jenis=@p2", runIdBaru, runIdAcuan, Jenis);
        }

        // ------------------------------------------------------------
        private static object[] Deret(object[,] v, int r, int kolAwal, int n, bool bulat)
        {
            var a = new object[n];
            for (int i = 0; i < n; i++)
            {
                object x = v[r, kolAwal + i];
                a[i] = x is double ? (bulat ? (object)Convert.ToInt32((double)x) : (double)x) : null;
            }
            return a;
        }

        private static double? Angka(object x) { return x is double ? (double?)(double)x : null; }
        private static int? Bulat(object x) { return x is double ? (int?)Convert.ToInt32((double)x) : null; }

        private static int? TahunDariTeks(string s)
        {
            for (int i = 0; i + 4 <= s.Length; i++)
            {
                int t;
                if (int.TryParse(s.Substring(i, 4), NumberStyles.None, CultureInfo.InvariantCulture, out t) && t > 1990 && t < 2200) return t;
            }
            return null;
        }

        private static string Teks(object x)
        {
            if (x == null || x is int) return "";   // int = kode error sel
            if (x is double && Math.Abs((double)x % 1) < 1e-9 && Math.Abs((double)x) >= 1e5)
                return ((double)x).ToString("0", CultureInfo.InvariantCulture);   // no. rekening tersimpan sebagai angka
            return Convert.ToString(x, CultureInfo.InvariantCulture).Trim();
        }
    }
}
