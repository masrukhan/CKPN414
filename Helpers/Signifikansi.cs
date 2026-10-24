using System;
using System.Globalization;
using CKPNLibrary.Panel;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Tahap 9b — penentuan debitur signifikan CKPN Individu.
    ///
    /// Dua metode, ditetapkan per grup per tahun di susunan grup:
    ///   • ranking : Top-N CIF dengan baki debet terbesar (perilaku lama, Master!C10).
    ///   • nominal : semua CIF dengan total baki debet ≥ ambang (mis. Rp1 M), sesuai opsi PAPSI BPRS 2026.
    ///
    /// Dasar nilai SAMA untuk kedua metode: baki debet per CIF (agregat semua kontrak di semua sheet KC).
    /// Plafon awal tidak tersedia di template APOLLO (kolom nilai kontrak = pokok + margin awal),
    /// dan kontrak hasil restrukturisasi tetap memakai baki debetnya, seperti ranking.
    ///
    /// Panel menulis aturan grup ke Master!H10 (kolom tersembunyi), mis. "nominal:1000000000".
    /// Kosong / "ranking" = Top-N Master!C10 (perilaku lama, juga untuk tombol Master tanpa panel).
    ///
    /// Modul memanggil <see cref="Batas"/> pada daftar CIF yang SUDAH diurutkan menurun berdasarkan
    /// baki debet, sehingga CKPN Individu, pengecualian Net Flow, Migration, dan bucket kolektif
    /// tahunan selalu memilih CIF yang sama:
    ///
    ///   int batas = Signifikansi.Batas(topN, pairs.Count, j => pairs[j].OS);
    /// </summary>
    internal static class Signifikansi
    {
        public const string Ranking = "ranking", Nominal = "nominal";
        public const string Sel = "H10", SelLabel = "G10", SelCatatan = "E10";
        public const double AmbangBawaan = 1000000000d;       // Rp1 M
        private const string AwalanCatatan = "Signifikansi:";

        /// <summary>Aturan yang berlaku saat ini (dibaca dari Master!H10 workbook aplikasi).</summary>
        public sealed class Aturan
        {
            public string Metode = Ranking;
            public double Ambang;
            public bool IsNominal { get { return Metode == Nominal; } }

            /// <summary>
            /// Jumlah CIF signifikan dari daftar terurut menurun (baki debet). Ranking → min(topN, jumlah);
            /// nominal → banyaknya CIF dengan baki debet ≥ ambang (topN diabaikan).
            /// </summary>
            public int Batas(int topN, int jumlah, Func<int, double> bakiKe)
            {
                if (jumlah <= 0) return 0;
                if (!IsNominal) return Math.Max(0, Math.Min(topN, jumlah));
                int n = 0;
                while (n < jumlah && bakiKe(n) >= Ambang) n++;
                return n;
            }
        }

        // ------------------------------------------------------------
        // Dipanggil modul perhitungan
        // ------------------------------------------------------------

        /// <summary>Jumlah CIF signifikan (lihat <see cref="Aturan.Batas"/>). bakiKe(i) = baki debet CIF ke-i.</summary>
        public static int Batas(int topN, int jumlah, Func<int, double> bakiKe)
        {
            return Baca().Batas(topN, jumlah, bakiKe);
        }

        /// <summary>
        /// Pengganti "if (topN &lt;= 0) return ..." di modul: pada metode nominal pengecualian tetap
        /// dijalankan walaupun Master!C10 kosong/0.
        /// </summary>
        public static bool Aktif(int topN)
        {
            return topN > 0 || Baca().IsNominal;
        }

        /// <summary>Aturan dari Master!H10 workbook aplikasi. Gagal membaca → ranking (perilaku lama).</summary>
        public static Aturan Baca()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app) ?? app.ActiveWorkbook;
                return Baca(wb);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca aturan signifikansi Master!" + Sel + " gagal: " + ex.Message);
                return new Aturan();
            }
        }

        public static Aturan Baca(Excel.Workbook wb)
        {
            Excel.Worksheet m = wb == null ? null : ParameterMaster.CariSheet(wb, "Master");
            return Urai(BacaMentah(m));
        }

        public static string BacaMentah(Excel.Worksheet master)
        {
            if (master == null) return "";
            try
            {
                object v = ((Excel.Range)master.Range[Sel]).Value2;
                return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            }
            catch { return ""; }
        }

        /// <summary>"nominal:1000000000" → nominal Rp1 M; selain itu ranking.</summary>
        public static Aturan Urai(string isi)
        {
            var a = new Aturan();
            if (string.IsNullOrEmpty(isi)) return a;
            int i = isi.IndexOf(':');
            string metode = (i < 0 ? isi : isi.Substring(0, i)).Trim().ToLowerInvariant();
            if (metode != Nominal) return a;
            double ambang;
            if (i < 0 || !double.TryParse(isi.Substring(i + 1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out ambang) || ambang <= 0)
                ambang = AmbangBawaan;
            a.Metode = Nominal;
            a.Ambang = ambang;
            return a;
        }

        public static string Isi(string metode, double? ambang)
        {
            if (metode != Nominal) return "";
            return Nominal + ":" + (ambang.HasValue && ambang.Value > 0 ? ambang.Value : AmbangBawaan).ToString("0", CultureInfo.InvariantCulture);
        }

        // ------------------------------------------------------------
        // Ditulis panel sebelum hitung grup
        // ------------------------------------------------------------

        /// <summary>
        /// Tulis aturan grup ke Master!H10, plus catatan terlihat di Master!E10 (di samping "Jumlah Top Debitur")
        /// supaya user tombol Master tahu Top-N tidak dipakai. Proteksi dibuka hanya bila isi berubah.
        /// </summary>
        public static void Tulis(Excel.Workbook wb, string metode, double? ambang)
        {
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            if (master == null) throw new InvalidOperationException("Sheet 'Master' tidak ditemukan.");
            string isi = Isi(metode, ambang);
            string catatan = isi.Length == 0 ? "" : AwalanCatatan + " baki debet per CIF ≥ " + Rupiah(Urai(isi).Ambang) +
                                                    " (Top-N tidak dipakai) — diisi panel";
            string catatanLama = "";
            try { object v = ((Excel.Range)master.Range[SelCatatan]).Value2; catatanLama = v == null ? "" : Convert.ToString(v); } catch { }
            // E10 hanya diubah bila kosong atau memang berisi catatan panel (jangan menimpa isian user)
            bool bolehCatatan = catatanLama.Length == 0 || catatanLama.StartsWith(AwalanCatatan, StringComparison.Ordinal);
            if (BacaMentah(master) == isi && (!bolehCatatan || catatanLama == catatan)) return;

            Protection.UbahMaster(wb, ms =>
            {
                Excel.Range r = (Excel.Range)ms.Range[Sel];
                r.NumberFormat = "@";
                r.Value2 = isi;
                try
                {
                    Excel.Range lbl = (Excel.Range)ms.Range[SelLabel];
                    if (lbl.Value2 == null) lbl.Value2 = "Signifikansi individu (diisi panel)";
                }
                catch { }
                if (bolehCatatan)
                    try
                    {
                        Excel.Range c = (Excel.Range)ms.Range[SelCatatan];
                        c.Value2 = catatan.Length == 0 ? null : catatan;
                        if (catatan.Length > 0) { c.Font.Italic = true; c.Font.Color = 0x2F5A8A; }   // BGR: coklat tua
                    }
                    catch { }
            });
        }

        // ------------------------------------------------------------
        // Label
        // ------------------------------------------------------------

        /// <summary>Label singkat, mis. "Top-N 25" atau "baki debet ≥ Rp1 M".</summary>
        public static string Label(string metode, int? topN, double? ambang)
        {
            if (metode == Nominal) return "baki debet ≥ " + Rupiah(ambang ?? AmbangBawaan);
            return topN.HasValue ? "Top-N " + topN.Value : "Top-N belum ditetapkan";
        }

        public static string Label(Aturan a, int topN)
        {
            return a.IsNominal ? Label(Nominal, null, a.Ambang) : Label(Ranking, topN, null);
        }

        /// <summary>Rp1 M / Rp500 jt / Rp1,5 M (format Indonesia ringkas).</summary>
        public static string Rupiah(double v)
        {
            var id = new CultureInfo("id-ID");
            if (v >= 1e12) return "Rp" + (v / 1e12).ToString("0.##", id) + " T";
            if (v >= 1e9) return "Rp" + (v / 1e9).ToString("0.##", id) + " M";
            if (v >= 1e6) return "Rp" + (v / 1e6).ToString("0.##", id) + " jt";
            return "Rp" + v.ToString("#,0", id);
        }
    }
}
