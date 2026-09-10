using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;                 // SheetSpec, SheetMode
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Mengisi sheet "Data Overview" — ringkasan identitas, informasi keuangan,
    /// tabel kualitas pembiayaan per segmen, dan PPKA per kualitas untuk BPR Syariah.
    ///
    /// Rantai path (sama seperti DashboardBuilder / PDMigration):
    ///   'Sumber Data'!E7   -> file "aplikasi CKPN" (punya sheet Master)
    ///   aplikasi Master!C4 -> Bulan Laporan            (-> B3)
    ///   aplikasi Master!D14-> file "template CKPN" (GB0101, GB0200, KC0600..KC1100)
    ///
    /// Isi Data Overview:
    ///   B2 Nama BPR (GB0101!E38)
    ///   B3 Bulan Laporan (aplikasi Master!C4)
    ///   B5 Asset (GB0200!E38)
    ///   B6 Outstanding/PYD (GB0200: E7+E16+E21+E24-E15)
    ///   B7 CKPN Neraca (GB0200!E36)
    ///   B8 Laba Rugi Tahun Lalu (GB0200!E71)
    ///   B9 Laba Rugi Berjalan (GB0200!E73)
    ///   C12:G17 Kualitas pembiayaan per segmen (OS/saldo modal per kualitas)
    ///   C19:G19 PPKA per kualitas (kolom CKPN tiap sheet KC, dijumlah lintas segmen)
    ///   Kolom H (JUMLAH) dibiarkan sebagai formula =SUM(...) di sheet.
    /// </summary>
    internal class DataOverviewBuilder
    {
        private readonly Excel.Application _app;

        private const string SheetDataSource   = "Sumber Data";
        private const string SheetDataOverview = "Data Overview";
        private const string SheetLog          = "Audit Log";

        // Segmen tabel kualitas -> baris C12..H17
        private static readonly string[] SegmenOverview =
            { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };
        private const int OvKualitasRowFirst = 12;   // C12..H17
        private const int PPKARowOut         = 19;   // C19..H19

        // Kolom CKPN per sheet KC (identik dgn CONTROL PPKA per KC di DashboardBuilder),
        // mulai baris data = spec.HeaderRow + 1 (baris 5).
        private static readonly object[][] MapCKPN = new object[][]
        {
            new object[] { "KC0600", "AV" },
            new object[] { "KC0700", "AV" },
            new object[] { "KC0800", "AV" },
            new object[] { "KC0900", "AT" },
            new object[] { "KC1000", "BB" },
            new object[] { "KC1100", "BA" },
        };

        public DataOverviewBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ----------------------------------------------------------------
        // Entry point — dipanggil dari CKPNFunctions (CKPN_RefreshDataOverview)
        // ----------------------------------------------------------------
        public void Hitung()
        {
            var wb = _app.ActiveWorkbook;

            var wsData = CariSheet(wb, SheetDataSource);
            var wsOv   = CariSheet(wb, SheetDataOverview);
            var wsLog  = CariSheet(wb, SheetLog);

            if (wsData == null) throw new InvalidOperationException("Sheet '" + SheetDataSource   + "' tidak ditemukan.");
            if (wsOv   == null) throw new InvalidOperationException("Sheet '" + SheetDataOverview + "' tidak ditemukan.");

            var hasilLog = new List<string>();

            // ---- Tahap 1: path aplikasi CKPN dari 'Sumber Data'!E7 ----
            string pathIndeks = ToStr(((Excel.Range)wsData.Range["E7"]).Value2);
            if (string.IsNullOrEmpty(pathIndeks) || !System.IO.File.Exists(pathIndeks))
            {
                KosongkanSemua(wsOv);
                hasilLog.Add("Path 'Sumber Data'!E7 kosong / tidak ditemukan -> semua dikosongkan.");
                Selesai(wsLog, hasilLog);
                return;
            }

            Excel.Workbook wbApp = null;   // aplikasi CKPN
            Excel.Workbook wbTpl = null;   // template CKPN
            try
            {
                // ---- Tahap 2: buka aplikasi, ambil Bulan Laporan (C4) & path template (D14) ----
                wbApp = _app.Workbooks.Open(pathIndeks, UpdateLinks: 0, ReadOnly: true);

                object bulanLaporan = null;
                string pathTemplate = "";
                if (ExcelHelper.SheetAda(wbApp, "Master"))
                {
                    var wsMasterApp = (Excel.Worksheet)wbApp.Worksheets["Master"];
                    bulanLaporan = ((Excel.Range)wsMasterApp.Range["C4"]).Value2;
                    pathTemplate = ToStr(((Excel.Range)wsMasterApp.Range["D14"]).Value2);
                }
                else
                {
                    hasilLog.Add("Sheet 'Master' tidak ada di file aplikasi (E7).");
                }

                ((Excel.Range)wsOv.Range["B3"]).Value2 = bulanLaporan ?? "";

                wbApp.Close(false); wbApp = null;

                if (string.IsNullOrEmpty(pathTemplate) || !System.IO.File.Exists(pathTemplate))
                {
                    KosongkanTemplate(wsOv);
                    hasilLog.Add("Path Master!D14 kosong / tidak ditemukan (" + pathTemplate + ").");
                    Selesai(wsLog, hasilLog);
                    return;
                }

                // ---- Tahap 3: buka template, isi seluruh Data Overview ----
                wbTpl = _app.Workbooks.Open(pathTemplate, UpdateLinks: 0, ReadOnly: true);

                IsiInformasiKeuangan(wbTpl, wsOv, hasilLog);                        // B2, B5:B9
                IsiTabelSegmen(wbTpl, wsOv, OvKualitasRowFirst, false, hasilLog);   // C12:G17  (outstanding)
                IsiTabelSegmen(wbTpl, wsOv, OvEADRowFirst,      true,  hasilLog);   // C25:G30  (EAD/tunggakan)
                IsiPPKAperKualitas(wbTpl, wsOv, hasilLog);                          // C19:G19

                wbTpl.Close(false); wbTpl = null;
            }
            catch (Exception ex)
            {
                hasilLog.Add("Gagal: " + ex.Message);
            }
            finally
            {
                if (wbApp != null) try { wbApp.Close(false); } catch { }
                if (wbTpl != null) try { wbTpl.Close(false); } catch { }
            }

            _app.Calculate();
            Selesai(wsLog, hasilLog);
        }

        // ================================================================
        // B2 + B5:B9 — identitas & informasi keuangan dari template
        // ================================================================
        private void IsiInformasiKeuangan(Excel.Workbook wbTpl, Excel.Worksheet wsOv, List<string> hasilLog)
        {
            // Nama BPR Syariah -> B2 (GB0101!E3)
            if (ExcelHelper.SheetAda(wbTpl, "GB0101"))
            {
                var wsGb01 = (Excel.Worksheet)wbTpl.Worksheets["GB0101"];
                ((Excel.Range)wsOv.Range["B2"]).Value2 = ToStr(((Excel.Range)wsGb01.Range["E3"]).Value2);
            }
            else
            {
                ((Excel.Range)wsOv.Range["B2"]).Value2 = "";
                hasilLog.Add("Sheet 'GB0101' tidak ada di template.");
            }

            // Informasi keuangan -> B5:B9 (semua dari GB0200)
            if (ExcelHelper.SheetAda(wbTpl, "GB0200"))
            {
                var wsGb = (Excel.Worksheet)wbTpl.Worksheets["GB0200"];

                double e7  = ToDouble(((Excel.Range)wsGb.Range["E7"]).Value2);
                double e15 = ToDouble(((Excel.Range)wsGb.Range["E15"]).Value2);
                double e16 = ToDouble(((Excel.Range)wsGb.Range["E16"]).Value2);
                double e21 = ToDouble(((Excel.Range)wsGb.Range["E21"]).Value2);
                double e24 = ToDouble(((Excel.Range)wsGb.Range["E24"]).Value2);

                double asset        = ToDouble(((Excel.Range)wsGb.Range["E38"]).Value2);
                double outstanding  = e7 + e16 + e21  + e24 - e15;
                double ckpnNeraca   = ToDouble(((Excel.Range)wsGb.Range["E36"]).Value2);
                double labaLalu     = ToDouble(((Excel.Range)wsGb.Range["E71"]).Value2);
                double labaBerjalan = ToDouble(((Excel.Range)wsGb.Range["E73"]).Value2);

                ((Excel.Range)wsOv.Range["B5"]).Value2 = asset;
                ((Excel.Range)wsOv.Range["B6"]).Value2 = outstanding;
                ((Excel.Range)wsOv.Range["B7"]).Value2 = ckpnNeraca;
                ((Excel.Range)wsOv.Range["B8"]).Value2 = labaLalu;
                ((Excel.Range)wsOv.Range["B9"]).Value2 = labaBerjalan;

                hasilLog.Add("Info keuangan OK (Asset=" + asset.ToString("#,##0") +
                             ", OS=" + outstanding.ToString("#,##0") + ").");
            }
            else
            {
                KosongkanInfoKeuangan(wsOv);
                hasilLog.Add("Sheet 'GB0200' tidak ada di template.");
            }
        }

        private const int OvEADRowFirst      = 25;   // C25..H30 (EAD / tunggakan)

        // --- Kolom khusus KC1000 & KC1100 ---
        // Kualitas aktiva = outstanding
        private const string KC1000_OS   = "AG";     // saldo modal
        private const string KC1100_Nilai = "AB";    // total nilai kontrak
        private const string KC1100_Susut = "AI";    // akumulasi penyusutan
        // EAD = tunggakan
        private const string KC1000_TungPokok = "AL";
        private const string KC1000_TungBasil = "AN";
        private const string KC1100_TungPokok = "AL";
        private const string KC1100_Ujroh     = "AM";

        // ead=false -> tabel Rincian Kualitas Aktiva (outstanding)
        // ead=true  -> tabel Rincian EAD CKPN (tunggakan untuk KC1000/KC1100)
        private void IsiTabelSegmen(Excel.Workbook wbTpl, Excel.Worksheet wsOv,
                                    int rowFirst, bool ead, List<string> hasilLog)
        {
            string tag = ead ? "EAD" : "Kualitas";

            for (int s = 0; s < SegmenOverview.Length; s++)
            {
                string   shName = SegmenOverview[s];
                int      rowOut = rowFirst + s;
                double[] bucket = new double[6];            // index 1..5 = Lancar..Macet

                SheetSpec spec;
                try { spec = SheetSpec.Buat(shName); }
                catch { NolkanKualitas(wsOv, rowOut); hasilLog.Add(tag + " " + shName + ": spec gagal."); continue; }

                if (!ExcelHelper.SheetAda(wbTpl, shName))
                {
                    NolkanKualitas(wsOv, rowOut);
                    hasilLog.Add(tag + " " + shName + ": sheet tidak ada di template.");
                    continue;
                }

                var ws      = (Excel.Worksheet)wbTpl.Worksheets[shName];
                int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
                if (lastRow <= spec.HeaderRow) { NolkanKualitas(wsOv, rowOut); continue; }

                int      dataStart = spec.HeaderRow + 1;
                string[] namaArr   = ExcelHelper.BacaKolomString(ws, spec.ColNama,     dataStart, lastRow);
                string[] kualArr   = ExcelHelper.BacaKolomString(ws, spec.ColKualitas, dataStart, lastRow);

                // Nilai per baris tergantung tabel (outstanding vs EAD) & segmen
                double[] nilaiArr = BacaNilaiSegmen(ws, spec, shName, ead, dataStart, lastRow);

                for (int i = 0; i < namaArr.Length; i++)
                {
                    if (string.IsNullOrEmpty(namaArr[i])) continue;   // lewati baris agunan
                    int kual = NormKualitas(kualArr[i]);              // 1..5
                    if (kual >= 1 && kual <= 5) bucket[kual] += nilaiArr[i];
                }

                // C..G = kualitas 1..5. Kolom H tetap formula =SUM(C:G) di sheet.
                ((Excel.Range)wsOv.Cells[rowOut, "C"]).Value2 = bucket[1];  // Lancar
                ((Excel.Range)wsOv.Cells[rowOut, "D"]).Value2 = bucket[2];  // DPK
                ((Excel.Range)wsOv.Cells[rowOut, "E"]).Value2 = bucket[3];  // Kurang Lancar
                ((Excel.Range)wsOv.Cells[rowOut, "F"]).Value2 = bucket[4];  // Diragukan
                ((Excel.Range)wsOv.Cells[rowOut, "G"]).Value2 = bucket[5];  // Macet

                hasilLog.Add(tag + " " + shName + " OK (L=" + bucket[1].ToString("#,##0") +
                            " Macet=" + bucket[5].ToString("#,##0") + ").");
            }
        }

        // Mengembalikan array nilai per baris (dataStart..lastRow)
        private double[] BacaNilaiSegmen(Excel.Worksheet ws, SheetSpec spec, string shName,
                                        bool ead, int dataStart, int lastRow)
        {
            // ---- KC1000 & KC1100: kolom khusus ----
            if (shName == "KC1000")
            {
                return ead
                    ? Jumlah2 (ws, KC1000_TungPokok, KC1000_TungBasil, dataStart, lastRow) // EAD: AL + AN
                    : ExcelHelper.BacaKolomDouble(ws, KC1000_OS, dataStart, lastRow);       // Kualitas: AG
            }
            if (shName == "KC1100")
            {
                return ead
                    ? Jumlah2 (ws, KC1100_TungPokok, KC1100_Ujroh, dataStart, lastRow)      // EAD: AL + AM
                    : Selisih2(ws, KC1100_Nilai,     KC1100_Susut, dataStart, lastRow);      // Kualitas: AB - AI
            }

            // ---- KC0600..KC0900: pakai SheetSpec (sama utk Kualitas & EAD) ----
            switch (spec.Mode)
            {
                case SheetMode.AF:
                    return ExcelHelper.BacaKolomDouble(ws, spec.ColOS,       dataStart, lastRow);
                case SheetMode.Tunggakan1:
                    return ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok, dataStart, lastRow);
                default: // SheetMode.Tunggakan2
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
        // C19:G19 — PPKA per kualitas, dijumlah lintas semua sheet KC
        // ================================================================
        private void IsiPPKAperKualitas(Excel.Workbook wbTpl, Excel.Worksheet wsOv, List<string> hasilLog)
        {
            double[] bucket = new double[6];   // index 1..5, kumulatif lintas KC

            foreach (var m in MapCKPN)
            {
                string shName  = (string)m[0];
                string colCkpn = (string)m[1];

                SheetSpec spec;
                try { spec = SheetSpec.Buat(shName); }
                catch { hasilLog.Add("PPKA " + shName + ": spec gagal."); continue; }

                if (!ExcelHelper.SheetAda(wbTpl, shName))
                {
                    hasilLog.Add("PPKA " + shName + ": sheet tidak ada di template.");
                    continue;
                }

                var ws      = (Excel.Worksheet)wbTpl.Worksheets[shName];
                int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
                if (lastRow <= spec.HeaderRow) continue;

                int      dataStart = spec.HeaderRow + 1;   // = baris 5
                string[] kualArr   = ExcelHelper.BacaKolomString(ws, spec.ColKualitas, dataStart, lastRow);
                double[] ckpnArr   = ExcelHelper.BacaKolomDouble(ws, colCkpn,          dataStart, lastRow);

                for (int i = 0; i < kualArr.Length; i++)
                {
                    int kual = NormKualitas(kualArr[i]);   // 1..5
                    if (kual >= 1 && kual <= 5) bucket[kual] += ckpnArr[i];
                }

                hasilLog.Add("PPKA " + shName + " (" + colCkpn + dataStart + ":bawah) OK.");
            }

            // C..G = kualitas 1..5. Kolom H19 tetap formula =SUM(C19:G19).
            ((Excel.Range)wsOv.Cells[PPKARowOut, "C"]).Value2 = bucket[1];  // Lancar
            ((Excel.Range)wsOv.Cells[PPKARowOut, "D"]).Value2 = bucket[2];  // DPK
            ((Excel.Range)wsOv.Cells[PPKARowOut, "E"]).Value2 = bucket[3];  // Kurang Lancar
            ((Excel.Range)wsOv.Cells[PPKARowOut, "F"]).Value2 = bucket[4];  // Diragukan
            ((Excel.Range)wsOv.Cells[PPKARowOut, "G"]).Value2 = bucket[5];  // Macet
        }

        // ================================================================
        // Helper pengosongan
        // ================================================================
        private void KosongkanSemua(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B2"]).Value2 = "";
            ((Excel.Range)ws.Range["B3"]).Value2 = "";
            KosongkanTemplate(ws);
        }

        private void KosongkanTemplate(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B2"]).Value2 = "";
            KosongkanInfoKeuangan(ws);
            for (int r = OvKualitasRowFirst; r < OvKualitasRowFirst + SegmenOverview.Length; r++)
                NolkanKualitas(ws, r);
            NolkanKualitas(ws, PPKARowOut);
        }

        private void KosongkanInfoKeuangan(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B5"]).Value2 = 0;
            ((Excel.Range)ws.Range["B6"]).Value2 = 0;
            ((Excel.Range)ws.Range["B7"]).Value2 = 0;
            ((Excel.Range)ws.Range["B8"]).Value2 = 0;
            ((Excel.Range)ws.Range["B9"]).Value2 = 0;
        }

        private void NolkanKualitas(Excel.Worksheet ws, int row)
        {
            ((Excel.Range)ws.Cells[row, "C"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "D"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "E"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "F"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "G"]).Value2 = 0;
        }

        // ================================================================
        // Selesai: tulis Audit Log (bila ada) + MessageBox
        // ================================================================
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

            System.Windows.Forms.MessageBox.Show(
                "Data Overview selesai di-refresh.\n\n" + string.Join("\n", hasilLog),
                "Data Overview CKPN",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
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

        private static string ToStr(object val)
        {
            return val == null ? "" : val.ToString().Trim();
        }

        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int    i) return i;
            if (val is long   l) return l;
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