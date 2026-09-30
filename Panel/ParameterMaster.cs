using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    // =====================================================================
    // Parameter per langkah — hasil pembacaan sheet Master oleh C#.
    //
    // PRINSIP: pembacaan di sini MENIRU PERSIS makro VBA yang selama ini
    // dipakai tombol Master (sel, default, dan format string yang dikirim ke
    // modul). Tujuannya: angka hasil hitungan dari panel identik dengan hasil
    // dari tombol. Kalau makro VBA diubah, bagian terkait di sini ikut diubah.
    //
    // Sumber tiap bagian:
    //   Individu   ← modCKPN_Individu.HitungCKPNIndividu
    //   Net Flow   ← IsiBucketPDNetflow_AllMonths
    //   Migration  ← HitungPDMigration / JalankanSemuaTriwulan
    //   LGD ER     ← HitungLGDExpectedRecoveries
    //   LGD CS     ← HitungLGDCollateralShortfallMacet
    //   Summary    ← RefreshLinkSummary
    // =====================================================================

    internal class ParamIndividu
    {
        public string FilePath;
        public int    TopN;
        public bool   NPF, Kol2, Restru;
    }

    internal class ParamNetFlow
    {
        // Disimpan dalam bentuk string "a|b|c|" persis seperti yang dikirim VBA,
        // lalu dipecah dengan aturan yang sama seperti CKPNFunctions.
        public string FilePaths, BulanLabels, RefDates;
        public int    TopN;
        public int    JumlahBulan;
    }

    internal class ParamMigration
    {
        public bool   RunAll;
        public string Triwulan;                  // mode satu triwulan
        public string PathAwal, PathAkhir;       // mode satu triwulan
        public string TglAwal, TglAkhir;         // "yyyyMMdd"
        public string AwalGab, AkhirGab, TglAwalGab, TglAkhirGab;   // mode Run All
        public int    TopN;
    }

    internal class ParamLGDER
    {
        public int    CurrYear;
        public string FilePathsStr;              // "2021=path|2022=path|..."
    }

    internal class ParamLGDCS
    {
        public string FileConfigStr;             // "label=tahun=bulan=path|..."
        public double Haircut;
        public int    JumlahPeriode;
    }

    internal class ParamSummary
    {
        public string FilePath;
    }

    /// <summary>Hasil baca Master: parameter + daftar error per langkah.</summary>
    internal class ParameterCKPN
    {
        public string   BulanLaporan = "";
        public string   KCList = "";               // "KC0600,KC0700"
        public string[] KC     = new string[0];

        public ParamIndividu  Individu;
        public ParamNetFlow   NetFlow;
        public ParamMigration Migration;
        public ParamLGDER     LgdEr;
        public ParamLGDCS     LgdCs;
        public ParamSummary   Summary;

        /// <summary>langkahId → daftar masalah. Langkah tanpa entri = siap.</summary>
        public readonly Dictionary<string, List<string>> Error =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>langkahId → ringkasan parameter untuk ditampilkan di panel.</summary>
        public readonly Dictionary<string, string> Ringkas =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public void Tambah(string langkah, string pesan)
        {
            List<string> l;
            if (!Error.TryGetValue(langkah, out l)) { l = new List<string>(); Error[langkah] = l; }
            l.Add(pesan);
        }

        public bool Siap(string langkah) { return !Error.ContainsKey(langkah); }
    }

    internal static class ParameterMaster
    {
        public static readonly string[] DaftarKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        // =============================================================
        // Entry point: baca seluruh parameter dari sheet Master.
        // Harus dipanggil di konteks makro (thread utama Excel).
        // =============================================================
        public static ParameterCKPN Baca(Excel.Workbook wb)
        {
            var p = new ParameterCKPN();
            Excel.Worksheet m = CariSheet(wb, "Master");
            if (m == null)
            {
                foreach (var id in CKPNPipeline.IdLangkahHitung) p.Tambah(id, "Sheet Master tidak ditemukan.");
                return p;
            }

            p.BulanLaporan = TeksTanggal(Nilai(m, "C4"));

            // ---- KC dicentang (semua langkah memakai daftar yang sama) ----
            var kc = BacaKCDicentang(m);
            p.KC     = kc.ToArray();
            p.KCList = string.Join(",", p.KC);
            if (kc.Count == 0)
                foreach (var id in CKPNPipeline.IdLangkahHitung) p.Tambah(id, "Tidak ada sheet KC yang dicentang.");

            BacaIndividu(m, p);
            BacaNetFlow(m, p);
            BacaMigration(m, p);
            BacaLgdEr(m, p);
            BacaLgdCs(m, p);
            BacaSummary(m, p);
            return p;
        }

        // ---------------- CKPN Individu ----------------
        private static void BacaIndividu(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "individu";
            var x = new ParamIndividu();
            x.FilePath = Teks(m, "D14");
            if (x.FilePath.Length == 0)       p.Tambah(id, "Path file (Master!D14) belum diisi.");
            else if (!File.Exists(x.FilePath)) p.Tambah(id, "File tidak ditemukan: " + x.FilePath);

            // VBA: topN = 10; If IsNumeric(C10) Then topN = CLng(C10); If topN < 1 Then topN = 10
            x.TopN = 10;
            object c10 = Nilai(m, "C10");
            if (IsNumericVBA(c10)) x.TopN = CLngVBA(c10);
            if (x.TopN < 1) x.TopN = 10;

            // VBA: IIf(UCase(Trim(CStr(D12))) = "YA", "Ya", "Tidak")
            x.NPF    = Teks(m, "D12").ToUpperInvariant() == "YA";
            x.Kol2   = Teks(m, "D13").ToUpperInvariant() == "YA";
            x.Restru = Teks(m, "F12").ToUpperInvariant() == "YA";

            p.Individu = x;
            p.Ringkas[id] = "Top-N " + x.TopN + " · NPF " + YaTidak(x.NPF) + " · Kol-2 " + YaTidak(x.Kol2) +
                            " · Restru " + YaTidak(x.Restru) + " · " + Path.GetFileName(x.FilePath);
        }

        // ---------------- PD Net Flow ----------------
        private static void BacaNetFlow(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "netflow";
            var x = new ParamNetFlow { FilePaths = "", BulanLabels = "", RefDates = "" };

            // VBA: For r = 20 To 32 — path di D, label di B, tanggal ref di F.
            // Baris dengan path kosong / "pilih folder" dilewati.
            for (int r = 20; r <= 32; r++)
            {
                string path = Teks(m, "D" + r);
                if (path.Length == 0 || path.ToLowerInvariant() == "pilih folder") continue;
                x.FilePaths   += path + "|";
                x.BulanLabels += Teks(m, "B" + r) + "|";
                x.RefDates    += TeksCStr(Nilai(m, "F" + r)) + "|";
                x.JumlahBulan++;
            }

            // VBA: If IsNumeric(C10) Then topN = CLng(C10)  (tanpa default → 0)
            object c10 = Nilai(m, "C10");
            x.TopN = IsNumericVBA(c10) ? CLngVBA(c10) : 0;

            if (x.JumlahBulan == 0) p.Tambah(id, "Belum ada path bulan yang diisi (Master!D20:D32).");

            p.NetFlow = x;
            p.Ringkas[id] = x.JumlahBulan + " dari 13 bulan terisi · Top-N " + x.TopN;
        }

        // ---------------- PD Migration ----------------
        private static void BacaMigration(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "migration";
            var x = new ParamMigration();
            object c10 = Nilai(m, "C10");
            x.TopN = IsNumericVBA(c10) ? CLngVBA(c10) : 0;
            p.Migration = x;

            string tw = Teks(m, "C52");
            if (tw.Length == 0) { p.Tambah(id, "Triwulan Dianalisis (Master!C52) belum diisi."); return; }

            if (string.Equals(tw, "Run All Triwulan", StringComparison.OrdinalIgnoreCase))
            {
                // ---- Mode Run All: 4 pasang file wajib ada ----
                x.RunAll = true;
                string[] urut = { "Triwulan I", "Triwulan II", "Triwulan III", "Triwulan IV" };
                var aw = new List<string>(); var ak = new List<string>();
                var ta = new List<string>(); var tk = new List<string>();
                foreach (var t in urut)
                {
                    string pa, pk, da, dk;
                    if (!BacaPathTriwulan(m, t, out pa, out pk)) { p.Tambah(id, "Path Periode Awal/Akhir untuk " + t + " belum diisi."); continue; }
                    if (!File.Exists(pa)) p.Tambah(id, "File Awal " + t + " tidak ditemukan: " + pa);
                    if (!File.Exists(pk)) p.Tambah(id, "File Akhir " + t + " tidak ditemukan: " + pk);
                    BacaTglTriwulan(m, t, out da, out dk);   // opsional di mode Run All
                    aw.Add(pa); ak.Add(pk); ta.Add(da); tk.Add(dk);
                }
                x.AwalGab     = string.Join("|", aw.ToArray());
                x.AkhirGab    = string.Join("|", ak.ToArray());
                x.TglAwalGab  = string.Join("|", ta.ToArray());
                x.TglAkhirGab = string.Join("|", tk.ToArray());
                p.Ringkas[id] = "Run All Triwulan (I–IV) · Top-N " + x.TopN;
                return;
            }

            // ---- Mode satu triwulan ----
            if (tw != "Triwulan I" && tw != "Triwulan II" && tw != "Triwulan III" && tw != "Triwulan IV")
            {
                p.Tambah(id, "Triwulan tidak dikenali: " + tw);
                return;
            }
            x.Triwulan = tw;
            string pAwal, pAkhir;
            if (!BacaPathTriwulan(m, tw, out pAwal, out pAkhir))
                p.Tambah(id, "Path Periode Awal/Akhir untuk " + tw + " belum diisi.");
            else
            {
                if (!File.Exists(pAwal))  p.Tambah(id, "File Awal tidak ditemukan: " + pAwal);
                if (!File.Exists(pAkhir)) p.Tambah(id, "File Akhir tidak ditemukan: " + pAkhir);
            }
            x.PathAwal = pAwal; x.PathAkhir = pAkhir;

            string tA, tK;
            if (!BacaTglTriwulan(m, tw, out tA, out tK))
                p.Tambah(id, "Periode WO untuk " + tw + " tidak terbaca di Master!B40:D43.");
            x.TglAwal = tA; x.TglAkhir = tK;

            p.Ringkas[id] = tw + " · Top-N " + x.TopN;
        }

        // VBA BacaPathTriwulan: baris 47..50, B = nama triwulan, D = awal, F = akhir
        private static bool BacaPathTriwulan(Excel.Worksheet m, string tw, out string awal, out string akhir)
        {
            awal = ""; akhir = "";
            for (int r = 47; r <= 50; r++)
            {
                if (Teks(m, "B" + r) == tw)   // VBA "=" pada string: peka huruf besar/kecil
                {
                    awal  = Teks(m, "D" + r);
                    akhir = Teks(m, "F" + r);
                    break;
                }
            }
            return awal.Length > 0 && akhir.Length > 0;
        }

        // VBA BacaTglTriwulan: baris 40..43, C = tgl awal (+1 hari), D = tgl akhir → "yyyymmdd"
        private static bool BacaTglTriwulan(Excel.Worksheet m, string tw, out string awal, out string akhir)
        {
            awal = ""; akhir = "";
            for (int r = 40; r <= 43; r++)
            {
                if (Teks(m, "B" + r) == tw)
                {
                    DateTime d;
                    // VBA IsDate(): sel angka biasa (bukan berformat tanggal) TIDAK dianggap tanggal.
                    if (KeTanggal(Nilai(m, "C" + r), out d, false)) awal  = d.AddDays(1).ToString("yyyyMMdd");
                    if (KeTanggal(Nilai(m, "D" + r), out d, false)) akhir = d.ToString("yyyyMMdd");
                    break;
                }
            }
            return awal.Length > 0 && akhir.Length > 0;
        }

        // ---------------- LGD Expected Recoveries ----------------
        private static void BacaLgdEr(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "lgder";
            var x = new ParamLGDER { FilePathsStr = "" };
            p.LgdEr = x;

            // VBA: currYear = YEAR(EDATE(C4, -11))
            DateTime c4;
            if (!KeTanggal(Nilai(m, "C4"), out c4))
            {
                p.Tambah(id, "Tanggal pelaporan (Master!C4) belum diisi / tidak dapat dibaca sebagai tanggal.");
                return;
            }
            x.CurrYear = c4.AddMonths(-11).Year;
            if (x.CurrYear < 2000 || x.CurrYear > 2100)
            {
                p.Tambah(id, "currYear hasil kalkulasi tidak valid: " + x.CurrYear + ". Periksa Master!C4.");
                return;
            }

            // VBA: For r = 60 To 65 — tahun di F, path di D. Semua wajib ada.
            for (int r = 60; r <= 65; r++)
            {
                object f = Nilai(m, "F" + r);
                if (!IsNumericVBA(f)) { p.Tambah(id, "Tahun di Master!F" + r + " tidak valid."); continue; }
                int y = CLngVBA(f);
                string path = Teks(m, "D" + r);
                if (path.Length == 0 || path.ToLowerInvariant().Contains("pilih"))
                {
                    p.Tambah(id, "Path file tahun " + y + " (Master!D" + r + ") belum diisi.");
                    continue;
                }
                if (!File.Exists(path)) p.Tambah(id, "File tahun " + y + " tidak ditemukan: " + path);
                if (x.FilePathsStr.Length > 0) x.FilePathsStr += "|";
                x.FilePathsStr += y + "=" + path;
            }
            p.Ringkas[id] = "Tahun berjalan " + x.CurrYear + " · cohort " + (x.CurrYear - 5) + "–" + x.CurrYear;
        }

        // ---------------- LGD Collateral Shortfall ----------------
        private static void BacaLgdCs(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "lgdcs";
            var x = new ParamLGDCS { FileConfigStr = "" };
            p.LgdCs = x;

            // VBA: Do While B(cfgRow) tidak kosong, mulai baris 72, maksimal ~200
            const int start = 72;
            int row = start;
            while (Teks(m, "B" + row).Length > 0)
            {
                row++;
                if (row > 200) break;
            }
            int n = row - start;
            if (n < 1)
            {
                p.Tambah(id, "Konfigurasi file referensi (Master B72:F...) kosong.");
                return;
            }

            // VBA: haircut = 0.2; If IsNumeric(C70) Then haircut = CDbl(C70)
            x.Haircut = 0.2;
            object c70 = Nilai(m, "C70");
            if (IsNumericVBA(c70)) x.Haircut = CDblVBA(c70);

            int valid = 0;
            for (int i = 0; i < n; i++)
            {
                int r = start + i;
                string lbl = Teks(m, "B" + r);
                object f = Nilai(m, "F" + r);
                if (!IsNumericVBA(f)) { p.Tambah(id, "Tahun di Master!F" + r + " tidak valid."); return; }
                int yr = CLngVBA(f);
                string path = Teks(m, "D" + r);
                int bln = lbl.IndexOf("juni", StringComparison.OrdinalIgnoreCase) >= 0 ? 6 : 12;

                if (x.FileConfigStr.Length > 0) x.FileConfigStr += "|";
                x.FileConfigStr += lbl + "=" + yr + "=" + bln + "=" + path;

                if (path.Length > 0 && path.ToLowerInvariant() != "pilih folder" && File.Exists(path)) valid++;
            }
            x.JumlahPeriode = n;
            p.Ringkas[id] = valid + " dari " + n + " file terbaca · haircut " + (x.Haircut * 100).ToString("0.0") + "%";
        }

        // ---------------- Refresh Summary ----------------
        private static void BacaSummary(Excel.Worksheet m, ParameterCKPN p)
        {
            const string id = "summary";
            var x = new ParamSummary { FilePath = Teks(m, "D14") };
            if (x.FilePath.Length == 0)        p.Tambah(id, "Path file sumber (Master!D14) belum diisi.");
            else if (!File.Exists(x.FilePath)) p.Tambah(id, "File sumber tidak ditemukan: " + x.FilePath);
            p.Summary = x;
            p.Ringkas[id] = "CKPN, PPKA & ABA dari " + Path.GetFileName(x.FilePath);
        }

        // =============================================================
        // Tanggal laporan (Master!C4) — kunci periode di database
        // =============================================================
        public static bool BacaTanggalLaporan(Excel.Workbook wb, out DateTime tanggal)
        {
            tanggal = DateTime.MinValue;
            Excel.Worksheet m = CariSheet(wb, "Master");
            return m != null && KeTanggal(Nilai(m, "C4"), out tanggal);
        }

        /// <summary>Path file template CKPN/SLIK (Master!D14) — dipakai juga oleh Overview Data.</summary>
        public static string BacaPathTemplate(Excel.Workbook wb)
        {
            Excel.Worksheet m = CariSheet(wb, "Master");
            return m == null ? "" : Teks(m, "D14");
        }

        // =============================================================
        // Checkbox KC di Master (Form Control lama, dibaca via late binding)
        // =============================================================
        /// <summary>
        /// Ubah centang KC di Master sesuai grup yang akan dihitung dari panel.
        /// Sheet Master terproteksi, jadi dibuka sementara lewat Protection.UbahMaster.
        /// xlOn = 1, xlOff = -4146.
        /// </summary>
        public static void SetKCDicentang(Excel.Workbook wb, ICollection<string> kcDipilih)
        {
            var pilih = new HashSet<string>(kcDipilih, StringComparer.OrdinalIgnoreCase);
            Protection.UbahMaster(wb, master =>
            {
                foreach (string kc in DaftarKC)
                {
                    object chk;
                    try { chk = master.CheckBoxes(kc); }
                    catch { throw new InvalidOperationException("Checkbox '" + kc + "' tidak ditemukan di sheet Master."); }
                    chk.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, chk,
                                               new object[] { pilih.Contains(kc) ? 1 : -4146 });
                }
            });

            // Verifikasi: centang yang terbaca harus sama persis dengan yang diminta
            var terbaca = BacaKCDicentang(CariSheet(wb, "Master"));
            if (terbaca.Count != pilih.Count || !terbaca.TrueForAll(pilih.Contains))
                throw new InvalidOperationException("Centang KC di Master gagal diubah (terbaca: " + string.Join(",", terbaca.ToArray()) + ").");
        }

        /// <summary>
        /// Tulis Top-N tahunan ke Master!C10 (dibaca Individu, Net Flow, dan Migration)
        /// sebelum hitung grup dari panel. Master terproteksi → lewat Protection.UbahMaster.
        /// </summary>
        public static void SetTopN(Excel.Workbook wb, int topN)
        {
            Excel.Worksheet m = CariSheet(wb, "Master");
            if (m == null) throw new InvalidOperationException("Sheet 'Master' tidak ditemukan.");
            object lama = Nilai(m, "C10");
            if (IsNumericVBA(lama) && CLngVBA(lama) == topN) return;   // sudah sesuai, tidak perlu buka proteksi
            Protection.UbahMaster(wb, master => { ((Excel.Range)master.Range["C10"]).Value2 = topN; });
        }

        /// <summary>Top-N yang sedang tertulis di Master!C10 (aturan sama dengan makro Individu; default 10).</summary>
        public static int BacaTopN(Excel.Workbook wb)
        {
            Excel.Worksheet m = CariSheet(wb, "Master");
            if (m == null) return 10;
            object c10 = Nilai(m, "C10");
            int n = IsNumericVBA(c10) ? CLngVBA(c10) : 10;
            return n < 1 ? 10 : n;
        }

        public static List<string> BacaKCDicentang(Excel.Worksheet master)
        {
            var hasil = new List<string>();
            foreach (string kc in DaftarKC)
            {
                try
                {
                    object chk = master.CheckBoxes(kc);
                    object val = chk.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, chk, null);
                    if (Convert.ToInt32(val) == 1) hasil.Add(kc);   // xlOn = 1
                }
                catch { /* checkbox tidak ada → dianggap tidak dicentang */ }
            }
            return hasil;
        }

        // =============================================================
        // Emulasi fungsi VBA yang dipakai makro
        // =============================================================

        /// <summary>Range.Value (bukan Value2): tanggal kembali sebagai DateTime, seperti di VBA.</summary>
        private static object Nilai(Excel.Worksheet ws, string alamat)
        {
            return ((Excel.Range)ws.Range[alamat]).get_Value(Type.Missing);
        }

        /// <summary>Trim(CStr(cell.Value)) — kosong bila null/error.</summary>
        private static string Teks(Excel.Worksheet ws, string alamat)
        {
            return TeksCStr(Nilai(ws, alamat)).Trim();
        }

        /// <summary>
        /// Meniru CStr(Value): tanggal → format tanggal pendek Windows,
        /// angka → representasi kultur aktif. Dipakai untuk label/tanggal ref
        /// Net Flow yang hanya dicatat di Audit Log.
        /// </summary>
        private static string TeksCStr(object v)
        {
            if (v == null) return "";
            if (v is int) return "";                 // nilai error sel (#N/A dll.) di interop
            if (v is DateTime) return ((DateTime)v).ToString("d", CultureInfo.CurrentCulture);
            if (v is double) return ((double)v).ToString(CultureInfo.CurrentCulture);
            return Convert.ToString(v, CultureInfo.CurrentCulture) ?? "";
        }

        /// <summary>IsNumeric VBA: Empty dianggap numerik (bernilai 0), teks angka juga.</summary>
        private static bool IsNumericVBA(object v)
        {
            if (v == null) return true;
            if (v is double || v is decimal || v is bool) return true;
            if (v is int) return false;              // sel error
            if (v is DateTime) return false;
            string s = Convert.ToString(v, CultureInfo.CurrentCulture).Trim();
            if (s.Length == 0) return false;
            double d;
            return double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d);
        }

        private static double CDblVBA(object v)
        {
            if (v == null) return 0;
            if (v is bool) return (bool)v ? -1 : 0;   // CDbl(True) = -1 di VBA
            if (v is double) return (double)v;
            if (v is decimal) return (double)(decimal)v;
            double d;
            double.TryParse(Convert.ToString(v, CultureInfo.CurrentCulture).Trim(),
                            NumberStyles.Any, CultureInfo.CurrentCulture, out d);
            return d;
        }

        /// <summary>CLng VBA: pembulatan banker's rounding (ke genap terdekat).</summary>
        private static int CLngVBA(object v)
        {
            return (int)Math.Round(CDblVBA(v), MidpointRounding.ToEven);
        }

        /// <summary>
        /// IsDate + CDate VBA untuk sel: DateTime atau teks tanggal; angka serial
        /// hanya diterima bila terimaAngka = true (makro LGD ER menerima
        /// IsNumeric(C4), sedangkan BacaTglTriwulan hanya IsDate).
        /// </summary>
        private static bool KeTanggal(object v, out DateTime d, bool terimaAngka = true)
        {
            d = DateTime.MinValue;
            if (v == null) return false;
            if (v is DateTime) { d = (DateTime)v; return true; }
            if (v is double)
            {
                if (!terimaAngka) return false;
                double x = (double)v;
                if (x <= 0 || x > 2958465) return false;
                d = DateTime.FromOADate(x);
                return true;
            }
            string s = Convert.ToString(v, CultureInfo.CurrentCulture).Trim();
            return s.Length > 0 && DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out d);
        }

        private static string TeksTanggal(object v)
        {
            DateTime d;
            if (KeTanggal(v, out d)) return d.ToString("dd MMMM yyyy", new CultureInfo("id-ID"));
            return TeksCStr(v);
        }

        private static string YaTidak(bool b) { return b ? "Ya" : "Tidak"; }

        internal static Excel.Worksheet CariSheet(Excel.Workbook wb, string nama)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (sh.Name.Equals(nama, StringComparison.OrdinalIgnoreCase)) return sh;
            return null;
        }
    }
}
