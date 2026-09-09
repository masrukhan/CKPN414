using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Refresh sheet "Data Overview" — identitas & informasi keuangan BPR Syariah.
    /// Rantai path (sama seperti TarikPPKAperKC):
    ///   'Sumber Data'!E7   -> file APLIKASI CKPN (punya sheet Master)
    ///   aplikasi Master!D14 -> file TEMPLATE CKPN (punya sheet GB0101, GB0200)
    ///   aplikasi Master!C4  -> Bulan Laporan (diambil dari APLIKASI, bukan template)
    /// </summary>
    internal class DataOverviewBuilder
    {
        private readonly Excel.Application _app;

        private const string SheetDataSource   = "Sumber Data";
        private const string SheetDataOverview = "Data Overview";
        private const string SheetLog          = "Audit Log";

        public DataOverviewBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        public void Hitung()
        {
            var wb     = _app.ActiveWorkbook;
            var wsData = CariSheet(wb, SheetDataSource);
            var wsOv   = CariSheet(wb, SheetDataOverview);
            var wsLog  = CariSheet(wb, SheetLog);

            if (wsData == null) throw new InvalidOperationException("Sheet '" + SheetDataSource + "' tidak ditemukan.");
            if (wsOv   == null) throw new InvalidOperationException("Sheet '" + SheetDataOverview + "' tidak ditemukan.");

            var hasilLog = new List<string>();

            // ---- Tahap 1: path APLIKASI CKPN dari 'Sumber Data'!E7 ----
            string pathAplikasi = ToStr(((Excel.Range)wsData.Range["E7"]).Value2);
            if (string.IsNullOrEmpty(pathAplikasi) || !System.IO.File.Exists(pathAplikasi))
            {
                KosongkanSemua(wsOv);
                hasilLog.Add("Data Overview: path E7 kosong/tidak ditemukan -> semua kosong");
                Selesai(wsLog, hasilLog);
                return;
            }

            Excel.Workbook wbApp = null;   // aplikasi CKPN
            Excel.Workbook wbTpl = null;   // template CKPN
            try
            {
                // ---- Tahap 2: buka APLIKASI, ambil Bulan Laporan (C4) & path TEMPLATE (D14) ----
                wbApp = _app.Workbooks.Open(pathAplikasi, UpdateLinks: 0, ReadOnly: true);

                object bulanLaporan = null;
                string pathTemplate = "";
                if (ExcelHelper.SheetAda(wbApp, "Master"))
                {
                    var wsMasterApp = (Excel.Worksheet)wbApp.Worksheets["Master"];
                    bulanLaporan = ((Excel.Range)wsMasterApp.Range["C4"]).Value2;   // B3
                    pathTemplate = ToStr(((Excel.Range)wsMasterApp.Range["D14"]).Value2);
                }
                else
                {
                    hasilLog.Add("Data Overview: sheet 'Master' tidak ada di file E7");
                }

                // Bulan Laporan -> B3 (dari aplikasi)
                ((Excel.Range)wsOv.Range["B3"]).Value2 = bulanLaporan ?? "";

                wbApp.Close(false); wbApp = null;

                if (string.IsNullOrEmpty(pathTemplate) || !System.IO.File.Exists(pathTemplate))
                {
                    KosongkanTemplate(wsOv);
                    hasilLog.Add("Data Overview: path Master!D14 kosong/tidak ditemukan (" + pathTemplate + ")");
                    Selesai(wsLog, hasilLog);
                    return;
                }

                // ---- Tahap 3: buka TEMPLATE, ambil identitas & informasi keuangan ----
                wbTpl = _app.Workbooks.Open(pathTemplate, UpdateLinks: 0, ReadOnly: true);

                // Nama BPR Syariah -> B2 (GB0101!E3)
                if (ExcelHelper.SheetAda(wbTpl, "GB0101"))
                {
                    var wsGb0101 = (Excel.Worksheet)wbTpl.Worksheets["GB0101"];
                    ((Excel.Range)wsOv.Range["B2"]).Value2 = ToStr(((Excel.Range)wsGb0101.Range["E3"]).Value2);
                }
                else
                {
                    ((Excel.Range)wsOv.Range["B2"]).Value2 = "";
                    hasilLog.Add("Data Overview: sheet 'GB0101' tidak ada di template");
                }

                // Informasi keuangan -> B5:B9 (semua dari GB0200)
                if (ExcelHelper.SheetAda(wbTpl, "GB0200"))
                {
                    var wsGb = (Excel.Worksheet)wbTpl.Worksheets["GB0200"];

                    double e7  = ToDouble(((Excel.Range)wsGb.Range["E7"]).Value2);
                    double e15 = ToDouble(((Excel.Range)wsGb.Range["E15"]).Value2);
                    double e16 = ToDouble(((Excel.Range)wsGb.Range["E16"]).Value2);
                    double e21 = ToDouble(((Excel.Range)wsGb.Range["E21"]).Value2);
                    double e22 = ToDouble(((Excel.Range)wsGb.Range["E22"]).Value2);
                    double e23 = ToDouble(((Excel.Range)wsGb.Range["E23"]).Value2);
                    double e24 = ToDouble(((Excel.Range)wsGb.Range["E24"]).Value2);

                    double asset        = ToDouble(((Excel.Range)wsGb.Range["E38"]).Value2);           // B5
                    double outstanding  = e7 + e16 + e21 + e22 - e23 + e24 - e15;                       // B6
                    double ckpnNeraca   = ToDouble(((Excel.Range)wsGb.Range["E36"]).Value2);           // B7
                    double labaLalu     = ToDouble(((Excel.Range)wsGb.Range["E71"]).Value2);           // B8
                    double labaBerjalan = ToDouble(((Excel.Range)wsGb.Range["E73"]).Value2);           // B9

                    ((Excel.Range)wsOv.Range["B5"]).Value2 = asset;
                    ((Excel.Range)wsOv.Range["B6"]).Value2 = outstanding;
                    ((Excel.Range)wsOv.Range["B7"]).Value2 = ckpnNeraca;
                    ((Excel.Range)wsOv.Range["B8"]).Value2 = labaLalu;
                    ((Excel.Range)wsOv.Range["B9"]).Value2 = labaBerjalan;

                    hasilLog.Add("Data Overview: OK (Asset=" + asset.ToString("#,##0") +
                                 ", Outstanding=" + outstanding.ToString("#,##0") + ")");
                }
                else
                {
                    KosongkanInfoKeuangan(wsOv);
                    hasilLog.Add("Data Overview: sheet 'GB0200' tidak ada di template");
                }

                wbTpl.Close(false); wbTpl = null;
            }
            catch (Exception ex)
            {
                hasilLog.Add("Data Overview: gagal (" + ex.Message + ")");
            }
            finally
            {
                if (wbApp != null) try { wbApp.Close(false); } catch { }
                if (wbTpl != null) try { wbTpl.Close(false); } catch { }
            }

            _app.Calculate();
            Selesai(wsLog, hasilLog);
        }

        // ---------- helper kosong ----------
        private void KosongkanSemua(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B2"]).Value2 = "";
            ((Excel.Range)ws.Range["B3"]).Value2 = "";
            KosongkanInfoKeuangan(ws);
        }
        private void KosongkanTemplate(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B2"]).Value2 = "";
            KosongkanInfoKeuangan(ws);
        }
        private void KosongkanInfoKeuangan(Excel.Worksheet ws)
        {
            ((Excel.Range)ws.Range["B5"]).Value2 = 0;
            ((Excel.Range)ws.Range["B6"]).Value2 = 0;
            ((Excel.Range)ws.Range["B7"]).Value2 = 0;
            ((Excel.Range)ws.Range["B8"]).Value2 = 0;
            ((Excel.Range)ws.Range["B9"]).Value2 = 0;
        }

        private void Selesai(Excel.Worksheet wsLog, List<string> hasilLog)
        {
            if (wsLog != null) { try { TulisAuditLog(wsLog, hasilLog); } catch { } }
            System.Windows.Forms.MessageBox.Show(
                "Data Overview selesai di-refresh.\n\n" + string.Join("\n", hasilLog),
                "Data Overview CKPN", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void TulisAuditLog(Excel.Worksheet wsLog, List<string> hasil)
        {
            int nextRow = NextAuditLogRow(wsLog);
            ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = DateTime.Now.ToOADate();
            ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "Data Overview Refresh";
            ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = string.Join(" | ", hasil);
        }
        private static int NextAuditLogRow(Excel.Worksheet wsLog)
        {
            for (int r = 5; r <= 1000000; r++)
                if (string.IsNullOrEmpty(ToStr(((Excel.Range)wsLog.Cells[r, 1]).Value2)) &&
                    string.IsNullOrEmpty(ToStr(((Excel.Range)wsLog.Cells[r, 2]).Value2)))
                    return r;
            return 1000001;
        }

        private static string ToStr(object val) => val == null ? "" : val.ToString().Trim();
        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int i) return i;
            double r; return double.TryParse(val.ToString(), out r) ? r : 0;
        }
        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) return sh;
            return null;
        }
    }
}