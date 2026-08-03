using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    internal class DokumentasiCKPNBuilder
    {
        private readonly Excel.Application _app;
        private const string SheetDataSource = "Sumber Data";
        private const string SheetDok        = "Dokumentasi CKPN";
        private const int DataSourceFirstRow = 7;
        private const int DataSourceLastRow  = 12;
        private const int DebiturFirstRow    = 25;

        public DokumentasiCKPNBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        public void Hitung()
        {
            var wb = _app.ActiveWorkbook;
            var wsData = CariSheet(wb, SheetDataSource);
            var wsDok  = CariSheet(wb, SheetDok);
            if (wsData == null) throw new InvalidOperationException("Sheet '" + SheetDataSource + "' tidak ditemukan.");
            if (wsDok  == null) throw new InvalidOperationException("Sheet '" + SheetDok + "' tidak ditemukan.");

            int nSegmen = DataSourceLastRow - DataSourceFirstRow + 1;
            var hasilLog = new List<string>();

            string pathKC0600 = ToStr(((Excel.Range)wsData.Cells[DataSourceFirstRow, "E"]).Value2);
            if (string.IsNullOrEmpty(pathKC0600) || !System.IO.File.Exists(pathKC0600))
                throw new InvalidOperationException("Path file KC0600 (Sumber Data baris 7) kosong/tidak ditemukan.");

            Excel.Workbook wbKC0600 = null;
            try
            {
                wbKC0600 = _app.Workbooks.Open(pathKC0600, UpdateLinks: 0, ReadOnly: true);

                if (ExcelHelper.SheetAda(wbKC0600, "Master"))
                {
                    var wsMaster = (Excel.Worksheet)wbKC0600.Worksheets["Master"];
                    ((Excel.Range)wsDok.Range["D6"]).Value2 = ((Excel.Range)wsMaster.Range["C4"]).Value2;
                    ((Excel.Range)wsDok.Range["D19"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["D12"]).Value2);
                    ((Excel.Range)wsDok.Range["D20"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["D13"]).Value2);
                    ((Excel.Range)wsDok.Range["D21"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["F12"]).Value2);
                }

                if (ExcelHelper.SheetAda(wbKC0600, "Master"))
                {
                    var wsMaster = (Excel.Worksheet)wbKC0600.Worksheets["Master"];
                    ((Excel.Range)wsDok.Range["D6"]).Value2  = ((Excel.Range)wsMaster.Range["C4"]).Value2;
                    ((Excel.Range)wsDok.Range["D19"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["D12"]).Value2);
                    ((Excel.Range)wsDok.Range["D20"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["D13"]).Value2);
                    ((Excel.Range)wsDok.Range["D21"]).Value2 = ToStr(((Excel.Range)wsMaster.Range["F12"]).Value2);

                    // ---- Section E: Periodesasi Perhitungan ----
                    ((Excel.Range)wsDok.Range["D34"]).Value2 = BacaRangeSebagaiTeks(wsMaster, "B20:B32");  // PD Net Flow
                    ((Excel.Range)wsDok.Range["D35"]).Value2 = BacaRangeSebagaiTeks(wsMaster, "B39:D43");  // PD Migration
                    ((Excel.Range)wsDok.Range["D36"]).Value2 = BacaRangeSebagaiTeks(wsMaster, "B60:B65");  // LGD Expected Recoveries
                    ((Excel.Range)wsDok.Range["D37"]).Value2 = BacaRangeSebagaiTeks(wsMaster, "B72:B83");  // LGD Collateral Shortfall
                }

                hasilLog.Add("KC0600 (Master/Audit Log): OK");
                wbKC0600.Close(false); wbKC0600 = null;
            }
            catch (Exception ex)
            {
                hasilLog.Add("KC0600: gagal dibaca (" + ex.Message + ")");
            }
            finally
            {
                if (wbKC0600 != null) try { wbKC0600.Close(false); } catch { }
            }

            for (int i = 0; i < nSegmen; i++)
            {
                int rowSrc = DataSourceFirstRow + i;
                int rowDok = DebiturFirstRow + i;
                string segmen = ToStr(((Excel.Range)wsData.Cells[rowSrc, "C"]).Value2);
                string path   = ToStr(((Excel.Range)wsData.Cells[rowSrc, "E"]).Value2);

                if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
                {
                    hasilLog.Add(segmen + ": path kosong/tidak ditemukan");
                    ((Excel.Range)wsDok.Cells[rowDok, "C"]).Value2 = 0;
                    continue;
                }

                Excel.Workbook wbSrc = null;
                try
                {
                    wbSrc = _app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true);
                    if (!ExcelHelper.SheetAda(wbSrc, "A. CKPN - INDV"))
                    {
                        hasilLog.Add(segmen + ": sheet 'A. CKPN - INDV' tidak ditemukan");
                        ((Excel.Range)wsDok.Cells[rowDok, "C"]).Value2 = 0;
                    }
                    else
                    {
                        var wsIndv = (Excel.Worksheet)wbSrc.Worksheets["A. CKPN - INDV"];
                        int jumlah = HitungDebiturCKPN(wsIndv);
                        ((Excel.Range)wsDok.Cells[rowDok, "C"]).Value2 = jumlah;
                        hasilLog.Add(segmen + ": " + jumlah + " debitur");
                    }
                    wbSrc.Close(false); wbSrc = null;
                }
                catch (Exception ex)
                {
                    hasilLog.Add(segmen + ": gagal dibaca (" + ex.Message + ")");
                    ((Excel.Range)wsDok.Cells[rowDok, "C"]).Value2 = 0;
                }
                finally
                {
                    if (wbSrc != null) try { wbSrc.Close(false); } catch { }
                }
            }

            _app.Calculate();
            System.Windows.Forms.MessageBox.Show(
                "Dokumentasi CKPN selesai diperbarui.\n\n" + string.Join("\n", hasilLog),
                "Dokumentasi CKPN", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private int HitungDebiturCKPN(Excel.Worksheet ws)
        {
            int lastRow = ExcelHelper.CariLastRow(ws, "B", 5);
            int count = 0;
            for (int r = 6; r <= lastRow; r++)
            {
                string idRek = ToStr(((Excel.Range)ws.Cells[r, "B"]).Value2);
                if (idRek.Equals("TOTAL", StringComparison.OrdinalIgnoreCase)) break;
                double penurunanNilai = ToDouble(((Excel.Range)ws.Cells[r, "K"]).Value2);
                if (penurunanNilai > 0) count++;
            }
            return count;
        }

        private static string ToStr(object val) { return val == null ? "" : val.ToString().Trim(); }
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
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase)) return sh;
            return null;
        }

        // ----------------------------------------------------------------
        // BacaRangeSebagaiTeks: baca range manapun (1 atau banyak kolom),
        // gabungkan tiap baris (kolom dipisah " - "), baris dipisah "; ".
        // Sel kosong diabaikan.
        // ----------------------------------------------------------------
        private string BacaRangeSebagaiTeks(Excel.Worksheet ws, string rangeAddr)
        {
            var rng = (Excel.Range)ws.Range[rangeAddr];
            object[,] arr = (object[,])rng.Value2;
            var barisList = new List<string>();
            int nRows = arr.GetLength(0);
            int nCols = arr.GetLength(1);

            for (int r = 1; r <= nRows; r++)
            {
                var kolom = new List<string>();
                for (int c = 1; c <= nCols; c++)
                {
                    string v = ToStr(arr[r, c]);
                    if (!string.IsNullOrEmpty(v)) kolom.Add(v);
                }
                if (kolom.Count > 0) barisList.Add(string.Join(" - ", kolom));
            }
            return string.Join("; ", barisList);
        }
    }
}