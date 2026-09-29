using System;
using System.Collections.Generic;
using System.IO;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;                 // SheetSpec, SheetMode
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Data Overview — identitas & informasi keuangan, outstanding (OS) dan EAD per segmen
    /// per kualitas, PPKA per kualitas, serta NPF.
    ///
    /// Dua cara pakai:
    ///
    ///  1. Panel CKPN (Tahap 3c) — UntukPanel(wbAplikasi)
    ///       Path langsung dari sheet Master workbook aplikasi yang sedang terbuka:
    ///         Master!C4  -> Bulan Laporan
    ///         Master!D14 -> file template CKPN (GB0101, GB0200, KC0600..KC1100)
    ///       Hasil dikembalikan sebagai data untuk tab "Overview Data"; tidak menulis sheet apa pun.
    ///
    ///  2. Workbook Dashboard lama — Hitung() (CKPN_RefreshDataOverview)
    ///       Mengisi sheet "Data Overview". Path diambil dari Master bila workbook aktif
    ///       adalah aplikasi CKPN; bila tidak, dari rantai lama 'Sumber Data'!E7 -> aplikasi -> Master.
    ///
    /// Definisi kolom (tidak berubah dari versi sebelumnya):
    ///   OS  KC0600-0800 : AC            KC0900 : AA
    ///       KC1000      : AG (saldo modal)
    ///       KC1100      : AB - AI (nilai kontrak - akumulasi penyusutan)
    ///   EAD KC0600-0900 : sama dengan OS (kolom SheetSpec)
    ///       KC1000      : AL + AN (tunggakan pokok + basil)
    ///       KC1100      : AL + AM (tunggakan pokok + ujroh)
    ///   Kualitas        : kolom SheetSpec.ColKualitas (1..5); baris tanpa nama (agunan) dilewati
    ///   PPKA            : MapCKPN per sheet
    ///   NPF             : (Kurang Lancar + Diragukan + Macet) / total OS
    /// </summary>
    internal class DataOverviewBuilder
    {
        private readonly Excel.Application _app;

        private const string SheetDataSource   = "Sumber Data";
        private const string SheetDataOverview = "Data Overview";
        private const string SheetLog          = "Audit Log";

        private static readonly string[] SegmenOverview =
            { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };
        private const int OvKualitasRowFirst = 12;   // C12..H17
        private const int PPKARowOut         = 19;   // C19..H19
        private const int OvEADRowFirst      = 25;   // C25..H30

        private static readonly Dictionary<string, string> KolomCKPN = new Dictionary<string, string>
        {
            { "KC0600", "AV" }, { "KC0700", "AV" }, { "KC0800", "AV" },
            { "KC0900", "AT" }, { "KC1000", "BB" }, { "KC1100", "BA" }
        };

        // --- Kolom khusus KC1000 & KC1100 ---
        private const string KC1000_OS        = "AG";
        private const string KC1100_Nilai     = "AB";
        private const string KC1100_Susut     = "AI";
        private const string KC1000_TungPokok = "AL";
        private const string KC1000_TungBasil = "AN";
        private const string KC1100_TungPokok = "AL";
        private const string KC1100_Ujroh     = "AM";

        public DataOverviewBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ================================================================
        // Model hasil baca template
        // ================================================================
        private class Segmen
        {
            public string Kode;
            public bool   Ada;
            public string Pesan = "";
            public int    Debitur;
            public double[] OS   = new double[6];   // index 1..5
            public double[] EAD  = new double[6];
            public double[] PPKA = new double[6];
        }

        private class HasilOverview
        {
            public string NamaBPR = "";
            public bool   AdaKeuangan;
            public double Asset, OSNeraca, CKPNNeraca, LabaLalu, LabaBerjalan;
            public List<Segmen> Segmen = new List<Segmen>();
            public List<string> Log = new List<string>();
        }

        // ================================================================
        // 1. Panel: path dari Master workbook aplikasi
        // ================================================================
        public Dictionary<string, object> UntukPanel(Excel.Workbook wbAplikasi)
        {
            string pathTemplate = CKPNLibrary.Panel.ParameterMaster.BacaPathTemplate(wbAplikasi);
            DateTime tgl;
            string periode = CKPNLibrary.Panel.ParameterMaster.BacaTanggalLaporan(wbAplikasi, out tgl)
                ? tgl.ToString("yyyy-MM-dd") : "";

            if (string.IsNullOrEmpty(pathTemplate))
                throw new InvalidOperationException("Path file template (Master!D14) belum diisi.");
            if (!File.Exists(pathTemplate))
                throw new InvalidOperationException("File template tidak ditemukan: " + pathTemplate);

            HasilOverview h = BacaDariFile(pathTemplate);

            // ---- susun untuk panel ----
            var segmen = new List<object>();
            double[] totOS = new double[6], totEAD = new double[6], totPPKA = new double[6];
            int totDebitur = 0;
            foreach (var s in h.Segmen)
            {
                for (int k = 1; k <= 5; k++) { totOS[k] += s.OS[k]; totEAD[k] += s.EAD[k]; totPPKA[k] += s.PPKA[k]; }
                totDebitur += s.Debitur;
                segmen.Add(KeData(s.Kode, s.Ada, s.Pesan, s.Debitur, s.OS, s.EAD, s.PPKA));
            }

            return new Dictionary<string, object>
            {
                { "periode", periode },
                { "fileTemplate", Path.GetFileName(pathTemplate) },
                { "pathTemplate", pathTemplate },
                { "waktu", DateTime.Now.ToString("yyyy-MM-dd HH:mm") },
                { "namaBPR", h.NamaBPR },
                { "keuangan", h.AdaKeuangan ? new Dictionary<string, object>
                    {
                        { "asset", h.Asset }, { "osNeraca", h.OSNeraca }, { "ckpnNeraca", h.CKPNNeraca },
                        { "labaLalu", h.LabaLalu }, { "labaBerjalan", h.LabaBerjalan }
                    } : null },
                { "segmen", segmen },
                { "total", KeData("Total", true, "", totDebitur, totOS, totEAD, totPPKA) },
                { "log", h.Log }
            };
        }

        private static Dictionary<string, object> KeData(string kode, bool ada, string pesan, int debitur,
                                                         double[] os, double[] ead, double[] ppka)
        {
            double tOS = Jumlah(os), tEAD = Jumlah(ead), tPPKA = Jumlah(ppka);
            double npfOS = os[3] + os[4] + os[5];
            double npfEAD = ead[3] + ead[4] + ead[5];
            return new Dictionary<string, object>
            {
                { "kode", kode }, { "ada", ada }, { "pesan", pesan }, { "debitur", debitur },
                { "os", Potong(os) }, { "ead", Potong(ead) }, { "ppka", Potong(ppka) },
                { "totalOS", tOS }, { "totalEAD", tEAD }, { "totalPPKA", tPPKA },
                { "selisih", tEAD - tOS },
                { "npfNominal", npfOS },
                { "npf", tOS > 0 ? npfOS / tOS : (double?)null },
                { "npfEAD", tEAD > 0 ? npfEAD / tEAD : (double?)null }
            };
        }

        private static double Jumlah(double[] b) { double t = 0; for (int k = 1; k <= 5; k++) t += b[k]; return t; }
        private static double[] Potong(double[] b) { return new[] { b[1], b[2], b[3], b[4], b[5] }; }

        /// <summary>
        /// Buka template read-only (atau pakai yang sudah terbuka), baca, lalu tutup
        /// hanya bila dibuka oleh fungsi ini.
        /// </summary>
        private HasilOverview BacaDariFile(string pathTemplate)
        {
            Excel.Workbook wbTpl = null;
            bool dibukaDiSini = false;
            bool layarLama = _app.ScreenUpdating, alertLama = _app.DisplayAlerts;
            try
            {
                _app.ScreenUpdating = false;
                _app.DisplayAlerts = false;

                string full = Path.GetFullPath(pathTemplate);
                foreach (Excel.Workbook wb in _app.Workbooks)
                    if (string.Equals(wb.FullName, full, StringComparison.OrdinalIgnoreCase)) { wbTpl = wb; break; }
                if (wbTpl == null)
                {
                    wbTpl = _app.Workbooks.Open(pathTemplate, UpdateLinks: 0, ReadOnly: true);
                    dibukaDiSini = true;
                }
                return BacaTemplate(wbTpl);
            }
            finally
            {
                if (dibukaDiSini && wbTpl != null) try { wbTpl.Close(false); } catch { }
                try { _app.DisplayAlerts = alertLama; _app.ScreenUpdating = layarLama; } catch { }
            }
        }

        // ================================================================
        // Inti: baca seluruh isi dari workbook template
        // ================================================================
        private HasilOverview BacaTemplate(Excel.Workbook wbTpl)
        {
            var h = new HasilOverview();

            if (ExcelHelper.SheetAda(wbTpl, "GB0101"))
            {
                var ws = (Excel.Worksheet)wbTpl.Worksheets["GB0101"];
                h.NamaBPR = ToStr(((Excel.Range)ws.Range["E3"]).Value2);
            }
            else h.Log.Add("Sheet 'GB0101' tidak ada di template.");

            if (ExcelHelper.SheetAda(wbTpl, "GB0200"))
            {
                var ws = (Excel.Worksheet)wbTpl.Worksheets["GB0200"];
                Func<string, double> v = a => ToDouble(((Excel.Range)ws.Range[a]).Value2);
                h.AdaKeuangan  = true;
                h.Asset        = v("E38");
                h.OSNeraca     = v("E7") + v("E16") + v("E21") + v("E24") - v("E15");
                h.CKPNNeraca   = v("E36");
                h.LabaLalu     = v("E71");
                h.LabaBerjalan = v("E73");
            }
            else h.Log.Add("Sheet 'GB0200' tidak ada di template.");

            foreach (string kode in SegmenOverview)
                h.Segmen.Add(BacaSegmen(wbTpl, kode, h.Log));
            return h;
        }

        private Segmen BacaSegmen(Excel.Workbook wbTpl, string kode, List<string> log)
        {
            var s = new Segmen { Kode = kode };
            SheetSpec spec;
            try { spec = SheetSpec.Buat(kode); }
            catch { s.Pesan = "spesifikasi sheet tidak dikenal"; log.Add(kode + ": " + s.Pesan); return s; }

            if (!ExcelHelper.SheetAda(wbTpl, kode))
            {
                s.Pesan = "sheet tidak ada di template";
                log.Add(kode + ": " + s.Pesan);
                return s;
            }
            s.Ada = true;

            var ws = (Excel.Worksheet)wbTpl.Worksheets[kode];
            int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
            if (lastRow <= spec.HeaderRow) { s.Pesan = "tidak ada data"; return s; }

            int dataStart = spec.HeaderRow + 1;
            string[] namaArr = ExcelHelper.BacaKolomString(ws, spec.ColNama, dataStart, lastRow);
            string[] kualArr = ExcelHelper.BacaKolomString(ws, spec.ColKualitas, dataStart, lastRow);
            double[] osArr   = BacaNilaiSegmen(ws, spec, kode, false, dataStart, lastRow);
            double[] eadArr  = BacaNilaiSegmen(ws, spec, kode, true,  dataStart, lastRow);
            string colCkpn;
            double[] ppkaArr = KolomCKPN.TryGetValue(kode, out colCkpn)
                ? ExcelHelper.BacaKolomDouble(ws, colCkpn, dataStart, lastRow)
                : new double[namaArr.Length];

            for (int i = 0; i < namaArr.Length; i++)
            {
                int kual = NormKualitas(kualArr[i]);
                if (kual < 1 || kual > 5) continue;
                // PPKA dijumlah untuk semua baris berkualitas (sama seperti CONTROL PPKA di Dashboard)
                s.PPKA[kual] += ppkaArr[i];
                if (string.IsNullOrEmpty(namaArr[i])) continue;   // baris agunan
                s.OS[kual]  += osArr[i];
                s.EAD[kual] += eadArr[i];
                s.Debitur++;
            }
            return s;
        }

        private double[] BacaNilaiSegmen(Excel.Worksheet ws, SheetSpec spec, string shName,
                                         bool ead, int dataStart, int lastRow)
        {
            if (shName == "KC1000")
                return ead
                    ? Jumlah2(ws, KC1000_TungPokok, KC1000_TungBasil, dataStart, lastRow)
                    : ExcelHelper.BacaKolomDouble(ws, KC1000_OS, dataStart, lastRow);
            if (shName == "KC1100")
                return ead
                    ? Jumlah2(ws, KC1100_TungPokok, KC1100_Ujroh, dataStart, lastRow)
                    : Selisih2(ws, KC1100_Nilai, KC1100_Susut, dataStart, lastRow);

            switch (spec.Mode)
            {
                case SheetMode.AF:
                    return ExcelHelper.BacaKolomDouble(ws, spec.ColOS, dataStart, lastRow);
                case SheetMode.Tunggakan1:
                    return ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok, dataStart, lastRow);
                default:
                    return Jumlah2(ws, spec.ColNomPokok, spec.ColNomBasil, dataStart, lastRow);
            }
        }

        private static double[] Jumlah2(Excel.Worksheet ws, string colA, string colB, int start, int end)
        {
            double[] a = ExcelHelper.BacaKolomDouble(ws, colA, start, end);
            double[] b = ExcelHelper.BacaKolomDouble(ws, colB, start, end);
            double[] r = new double[a.Length];
            for (int i = 0; i < a.Length; i++) r[i] = a[i] + b[i];
            return r;
        }

        private static double[] Selisih2(Excel.Worksheet ws, string colA, string colB, int start, int end)
        {
            double[] a = ExcelHelper.BacaKolomDouble(ws, colA, start, end);
            double[] b = ExcelHelper.BacaKolomDouble(ws, colB, start, end);
            double[] r = new double[a.Length];
            for (int i = 0; i < a.Length; i++) r[i] = a[i] - b[i];
            return r;
        }

        // ================================================================
        // 2. Sheet "Data Overview" (workbook Dashboard lama)
        // ================================================================
        public void Hitung()
        {
            var wb = _app.ActiveWorkbook;
            var wsOv  = CariSheet(wb, SheetDataOverview);
            var wsLog = CariSheet(wb, SheetLog);
            if (wsOv == null) throw new InvalidOperationException("Sheet '" + SheetDataOverview + "' tidak ditemukan.");

            var hasilLog = new List<string>();
            object bulanLaporan = null;
            string pathTemplate = "";

            // Path langsung dari Master bila workbook aktif punya sheet Master (aplikasi CKPN)
            var wsMaster = CariSheet(wb, "Master");
            if (wsMaster != null)
            {
                bulanLaporan = ((Excel.Range)wsMaster.Range["C4"]).Value2;
                pathTemplate = ToStr(((Excel.Range)wsMaster.Range["D14"]).Value2);
                hasilLog.Add("Path dari Master!D14 workbook aktif.");
            }
            else
            {
                // Rantai lama: 'Sumber Data'!E7 -> aplikasi CKPN -> Master!C4/D14
                var wsData = CariSheet(wb, SheetDataSource);
                if (wsData == null)
                    throw new InvalidOperationException("Sheet 'Master' atau '" + SheetDataSource + "' tidak ditemukan.");
                string pathApp = ToStr(((Excel.Range)wsData.Range["E7"]).Value2);
                if (string.IsNullOrEmpty(pathApp) || !File.Exists(pathApp))
                {
                    KosongkanSemua(wsOv);
                    hasilLog.Add("Path 'Sumber Data'!E7 kosong / tidak ditemukan -> semua dikosongkan.");
                    Selesai(wsLog, hasilLog);
                    return;
                }
                Excel.Workbook wbApp = null;
                try
                {
                    wbApp = _app.Workbooks.Open(pathApp, UpdateLinks: 0, ReadOnly: true);
                    var m = CariSheet(wbApp, "Master");
                    if (m != null)
                    {
                        bulanLaporan = ((Excel.Range)m.Range["C4"]).Value2;
                        pathTemplate = ToStr(((Excel.Range)m.Range["D14"]).Value2);
                    }
                    else hasilLog.Add("Sheet 'Master' tidak ada di file aplikasi (E7).");
                }
                finally { if (wbApp != null) try { wbApp.Close(false); } catch { } }
            }

            ((Excel.Range)wsOv.Range["B3"]).Value2 = bulanLaporan ?? "";

            if (string.IsNullOrEmpty(pathTemplate) || !File.Exists(pathTemplate))
            {
                KosongkanTemplate(wsOv);
                hasilLog.Add("Path Master!D14 kosong / tidak ditemukan (" + pathTemplate + ").");
                Selesai(wsLog, hasilLog);
                return;
            }

            try
            {
                HasilOverview h = BacaDariFile(pathTemplate);
                hasilLog.AddRange(h.Log);

                ((Excel.Range)wsOv.Range["B2"]).Value2 = h.NamaBPR;
                if (h.AdaKeuangan)
                {
                    ((Excel.Range)wsOv.Range["B5"]).Value2 = h.Asset;
                    ((Excel.Range)wsOv.Range["B6"]).Value2 = h.OSNeraca;
                    ((Excel.Range)wsOv.Range["B7"]).Value2 = h.CKPNNeraca;
                    ((Excel.Range)wsOv.Range["B8"]).Value2 = h.LabaLalu;
                    ((Excel.Range)wsOv.Range["B9"]).Value2 = h.LabaBerjalan;
                }
                else KosongkanInfoKeuangan(wsOv);

                double[] ppka = new double[6];
                for (int i = 0; i < h.Segmen.Count; i++)
                {
                    var s = h.Segmen[i];
                    TulisBaris(wsOv, OvKualitasRowFirst + i, s.OS);
                    TulisBaris(wsOv, OvEADRowFirst + i, s.EAD);
                    for (int k = 1; k <= 5; k++) ppka[k] += s.PPKA[k];
                    hasilLog.Add(s.Kode + (s.Ada ? " OK (OS=" + Jumlah(s.OS).ToString("#,##0") + ")" : ": " + s.Pesan));
                }
                TulisBaris(wsOv, PPKARowOut, ppka);
            }
            catch (Exception ex)
            {
                hasilLog.Add("Gagal: " + ex.Message);
            }

            _app.Calculate();
            Selesai(wsLog, hasilLog);
        }

        private static void TulisBaris(Excel.Worksheet ws, int row, double[] b)
        {
            // C..G = kualitas 1..5. Kolom H tetap formula =SUM(C:G) di sheet.
            ((Excel.Range)ws.Cells[row, "C"]).Value2 = b[1];
            ((Excel.Range)ws.Cells[row, "D"]).Value2 = b[2];
            ((Excel.Range)ws.Cells[row, "E"]).Value2 = b[3];
            ((Excel.Range)ws.Cells[row, "F"]).Value2 = b[4];
            ((Excel.Range)ws.Cells[row, "G"]).Value2 = b[5];
        }

        private void KosongkanSemua(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B3"]).Value2 = "";
            KosongkanTemplate(ws);
        }

        private void KosongkanTemplate(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B2"]).Value2 = "";
            KosongkanInfoKeuangan(ws);
            var nol = new double[6];
            for (int r = 0; r < SegmenOverview.Length; r++)
            {
                TulisBaris(ws, OvKualitasRowFirst + r, nol);
                TulisBaris(ws, OvEADRowFirst + r, nol);
            }
            TulisBaris(ws, PPKARowOut, nol);
        }

        private void KosongkanInfoKeuangan(Excel.Worksheet ws)
        {
            foreach (var a in new[] { "B5", "B6", "B7", "B8", "B9" })
                ((Excel.Range)ws.Range[a]).Value2 = 0;
        }

        private void Selesai(Excel.Worksheet wsLog, List<string> hasilLog)
        {
            if (wsLog != null)
            {
                try
                {
                    int r = NextAuditLogRow(wsLog);
                    ((Excel.Range)wsLog.Cells[r, 1]).Value2 = DateTime.Now.ToOADate();
                    ((Excel.Range)wsLog.Cells[r, 1]).NumberFormat = "m/d/yyyy h:mm";
                    ((Excel.Range)wsLog.Cells[r, 2]).Value2 = "Data Overview Refresh";
                    ((Excel.Range)wsLog.Cells[r, 3]).Value2 = string.Join(" | ", hasilLog);
                }
                catch { }
            }
            Pemberitahu.Info("Data Overview CKPN", "Data Overview selesai di-refresh.\n\n" + string.Join("\n", hasilLog));
        }

        private static int NextAuditLogRow(Excel.Worksheet wsLog)
        {
            for (int r = 5; r <= 1000000; r++)
            {
                if (string.IsNullOrEmpty(ToStr(((Excel.Range)wsLog.Cells[r, 1]).Value2)) &&
                    string.IsNullOrEmpty(ToStr(((Excel.Range)wsLog.Cells[r, 2]).Value2)))
                    return r;
            }
            return 1000001;
        }

        // ================================================================
        // Utility
        // ================================================================
        private static int NormKualitas(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return 0;
            string s = val.Trim();
            int i = 0;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            if (i == 0) return 0;
            int n;
            return (int.TryParse(s.Substring(0, i), out n) && n >= 1 && n <= 5) ? n : 0;
        }

        private static string ToStr(object val) { return val == null ? "" : val.ToString().Trim(); }

        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int i) return i;
            if (val is long l) return l;
            double r;
            return double.TryParse(val.ToString(), out r) ? r : 0;
        }

        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }
    }
}
