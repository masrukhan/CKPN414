using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Membangun/refresh sheet "Dashboard" — ringkasan CKPN vs PPKA
    /// (Individual & Kolektif) untuk seluruh segmen, snapshot dari sheet
    /// "Summary" pada masing-masing file hasil export CKPN per segmen.
    ///
    /// Sumber path: sheet "Sumber Data" kolom C (Segmen KC) & E (Path),
    /// baris 7..12 (6 segmen).
    /// </summary>
    internal class DashboardBuilder
    {
        private readonly Excel.Application _app;

        private const string SheetDataSource = "Sumber Data";
        private const string SheetDashboard  = "Dashboard";
        private const string SheetLog        = "Audit Log";
        private const string SheetSummary    = "Summary";
        private const string SheetRiwayat = "Riwayat CKPN";

        private const int DataSourceFirstRow = 7;
        private const int DataSourceLastRow  = 12;   // 6 segmen

        private const int DashARowFirst = 10;   // Section A (PD Net Flow) data start
        private const int DashCRowFirst = 21;  // Section C (PD Migration) data start

        public DashboardBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        public void Hitung()
        {
            var wb = _app.ActiveWorkbook;

            var wsData = CariSheet(wb, SheetDataSource);
            var wsDash = CariSheet(wb, SheetDashboard);
            var wsLog  = CariSheet(wb, SheetLog);

            if (wsData == null) throw new InvalidOperationException("Sheet '" + SheetDataSource + "' tidak ditemukan.");
            if (wsDash == null) throw new InvalidOperationException("Sheet '" + SheetDashboard  + "' tidak ditemukan.");

            int nSegmen = DataSourceLastRow - DataSourceFirstRow + 1;
            var hasilLog = new List<string>();

            const int DashDRowFirst = 61;   // Section D (PD Migration & LGD) data start

            double abaCkpn  = 0;   // Summary!C26 -> F9 & F20
            double abaPpka  = 0;   // Summary!H2  -> G9 & G20
            bool   abaFound = false;

            for (int i = 0; i < nSegmen; i++)
            {
                int rowSrc   = DataSourceFirstRow + i;
                int rowDashA = DashARowFirst + i;
                int rowDashC = DashCRowFirst + i;
                int rowDashD = DashDRowFirst + i;

                string segmen = ToStr(((Excel.Range)wsData.Cells[rowSrc, "C"]).Value2);
                string path   = ToStr(((Excel.Range)wsData.Cells[rowSrc, "E"]).Value2);

                if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
                {
                    hasilLog.Add(segmen + ": path kosong/tidak ditemukan");
                    TulisKosong(wsDash, rowDashA);
                    TulisKosong(wsDash, rowDashC);
                    TulisKosongPDLGD(wsDash, rowDashD);
                    continue;
                }

                Excel.Workbook wbSrc = null;
                try
                {
                    wbSrc = _app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true);

                    // ---- Top N (Individual) dari sheet Master!C10 ----
                    double topN = 0;
                    if (ExcelHelper.SheetAda(wbSrc, "Master"))
                    {
                        var wsMaster = (Excel.Worksheet)wbSrc.Worksheets["Master"];
                        topN = ToDouble(((Excel.Range)wsMaster.Range["C10"]).Value2);
                    }

                    if (!ExcelHelper.SheetAda(wbSrc, SheetSummary))
                    {
                        hasilLog.Add(segmen + ": sheet 'Summary' tidak ditemukan");
                        TulisKosong(wsDash, rowDashA);
                        TulisKosong(wsDash, rowDashC);
                    }
                    else
                    {
                        var wsSum = (Excel.Worksheet)wbSrc.Worksheets[SheetSummary];

                        if (!abaFound)
                        {
                            abaCkpn  = ToDouble(((Excel.Range)wsSum.Range["C26"]).Value2);
                            abaPpka  = ToDouble(((Excel.Range)wsSum.Range["H2"]).Value2);
                            abaFound = true;
                        }

                        ((Excel.Range)wsDash.Cells[rowDashA, "C"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["B6"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashA, "D"]).Value2 = topN;
                        ((Excel.Range)wsDash.Cells[rowDashA, "E"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["C6"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashA, "F"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["B7"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashA, "G"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["C7"]).Value2);

                        ((Excel.Range)wsDash.Cells[rowDashC, "C"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["B13"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashC, "D"]).Value2 = topN;
                        ((Excel.Range)wsDash.Cells[rowDashC, "E"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["C13"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashC, "F"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["B14"]).Value2);
                        ((Excel.Range)wsDash.Cells[rowDashC, "G"]).Value2 = ToDouble(((Excel.Range)wsSum.Range["C14"]).Value2);

                        hasilLog.Add(segmen + ": OK (Top N=" + topN + ")");
                    }

                    // ---- Section D & E: dari sheet 'B. CKPN - KOL INDV' ----
                    if (!ExcelHelper.SheetAda(wbSrc, "B. CKPN - KOL INDV"))
                    {
                        hasilLog.Add(segmen + ": sheet 'B. CKPN - KOL INDV' tidak ditemukan");
                        TulisKosongPDLGD(wsDash, rowDashD);
                        TulisKosongPDNetFlow(wsDash, i);
                    }
                    else
                    {
                        var wsKol = (Excel.Worksheet)wbSrc.Worksheets["B. CKPN - KOL INDV"];

                        // Section D: PD Migration per Kualitas (D38:D42) + LGD Weighted (E38)
                        for (int k = 0; k < 5; k++)
                        {
                            int rowSumber = 38 + k;
                            string colDash = ((char)('C' + k)).ToString();
                            ((Excel.Range)wsDash.Cells[rowDashD, colDash]).Value2 =
                                ToDouble(((Excel.Range)wsKol.Cells[rowSumber, "D"]).Value2);
                        }
                        ((Excel.Range)wsDash.Cells[rowDashD, "H"]).Value2 =
                            ToDouble(((Excel.Range)wsKol.Range["E38"]).Value2);

                        // Section E: PD Net Flow per Bucket Hari Tunggakan (E6:E19 — 14 bucket)
                        string colDashE = ((char)('D' + i)).ToString();   // D,E,F,G,H,I untuk segmen ke-0..5
                        for (int b = 0; b < 14; b++)
                        {
                            int rowSumber = 6 + b;    // E6..E19
                            int rowDashE  = 73 + b;   // baris 61..74
                            ((Excel.Range)wsDash.Cells[rowDashE, colDashE]).Value2 =
                                ToDouble(((Excel.Range)wsKol.Cells[rowSumber, "E"]).Value2);
                        }
                    }

                    wbSrc.Close(false); wbSrc = null;
                }
                catch (Exception ex)
                {
                    hasilLog.Add(segmen + ": gagal dibaca (" + ex.Message + ")");
                    TulisKosong(wsDash, rowDashA);
                    TulisKosong(wsDash, rowDashC);
                    TulisKosongPDLGD(wsDash, rowDashD);
                }
                finally
                {
                    if (wbSrc != null) try { wbSrc.Close(false); } catch { }
                }
            }

            ((Excel.Range)wsDash.Range["F9"]).Value2  = abaCkpn;   // CKPN Kolektif
            ((Excel.Range)wsDash.Range["G9"]).Value2  = abaPpka;   // PPKA Kolektif
            ((Excel.Range)wsDash.Range["F20"]).Value2 = abaCkpn;
            ((Excel.Range)wsDash.Range["G20"]).Value2 = abaPpka;

            ((Excel.Range)wsDash.Range["C5"]).Value2 = DateTime.Now.ToOADate();
            _app.Calculate();

            if (wsLog != null)
            {
                try { TulisAuditLog(wsLog, hasilLog); } catch { }
            }

            var wsRiwayat = CariSheet(wb, SheetRiwayat);
            if (wsRiwayat != null)
            {
                try { TulisRiwayat(wsRiwayat, wsDash); } catch { }
            }

            System.Windows.Forms.MessageBox.Show(
                "Dashboard selesai di-refresh.\n\n" + string.Join("\n", hasilLog),
                "Dashboard CKPN", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void TulisKosongPDLGD(Excel.Worksheet ws, int row)
        {
            ((Excel.Range)ws.Cells[row, "C"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "D"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "E"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "F"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "G"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "H"]).Value2 = 0;
        }

        // ----------------------------------------------------------------
        // TulisRiwayat: tambahkan 1 baris snapshot ke sheet "Riwayat CKPN"
        // setiap kali Dashboard di-refresh — bukan menimpa, tapi APPEND
        // di baris kosong berikutnya (mulai baris 6).
        // ----------------------------------------------------------------
        private void TulisRiwayat(Excel.Worksheet wsRiwayat, Excel.Worksheet wsDash)
        {
            int nextRow = 6;
            while (!string.IsNullOrEmpty(ToStr(((Excel.Range)wsRiwayat.Cells[nextRow, "B"]).Value2)))
                nextRow++;

            ((Excel.Range)wsRiwayat.Cells[nextRow, "B"]).Value2 = DateTime.Now.ToOADate();
            ((Excel.Range)wsRiwayat.Cells[nextRow, "B"]).NumberFormat = "m/d/yyyy h:mm";

            // Grand Total Section A (PD Net Flow) — baris 15 di Dashboard
            ((Excel.Range)wsRiwayat.Cells[nextRow, "C"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["H16"]).Value2);
            ((Excel.Range)wsRiwayat.Cells[nextRow, "D"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["I16"]).Value2);
            ((Excel.Range)wsRiwayat.Cells[nextRow, "E"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["J16"]).Value2);

            // Grand Total Section C (PD Migration) — baris 25 di Dashboard
            ((Excel.Range)wsRiwayat.Cells[nextRow, "F"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["H27"]).Value2);
            ((Excel.Range)wsRiwayat.Cells[nextRow, "G"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["I27"]).Value2);
            ((Excel.Range)wsRiwayat.Cells[nextRow, "H"]).Value2 = ToDouble(((Excel.Range)wsDash.Range["J27"]).Value2);
        }

        private void TulisKosong(Excel.Worksheet ws, int row)
        {
            ((Excel.Range)ws.Cells[row, "C"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "D"]).Value2 = "";   // Top N tidak diketahui — bukan 0
            ((Excel.Range)ws.Cells[row, "E"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "F"]).Value2 = 0;
            ((Excel.Range)ws.Cells[row, "G"]).Value2 = 0;
        }

        private void TulisAuditLog(Excel.Worksheet wsLog, List<string> hasil)
        {
            int nextRow = NextAuditLogRow(wsLog);
            ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = DateTime.Now.ToOADate();
            ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "Dashboard Refresh";
            ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = string.Join(" | ", hasil);
        }

        private void TulisKosongPDNetFlow(Excel.Worksheet ws, int segmenIndex)
        {
            string col = ((char)('D' + segmenIndex)).ToString();
            for (int b = 0; b < 14; b++)
                ((Excel.Range)ws.Cells[73 + b, col]).Value2 = 0;
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

        private static string ToStr(object val) => val == null ? "" : val.ToString().Trim();

        private static double ToDouble(object val)
        {
            if (val == null) return 0;
            if (val is double d) return d;
            if (val is int i) return i;
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