using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;
using CKPNLibrary.Styling;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    internal class CKPNIndividu
    {
        private readonly Excel.Application _app;

        public CKPNIndividu(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        public void Hitung(
            string filePath, int topN,
            bool flagNPF, bool flagKol2, bool flagRestru,
            string[] sheetKCList)
        {
            var wsOut = CariSheet(_app.ActiveWorkbook, "A. CKPN - INDV");
            if (wsOut == null)
                throw new InvalidOperationException("Sheet 'A. CKPN - INDV' tidak ditemukan.");

            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                var dictCIF = new Dictionary<string, CIFData>(StringComparer.OrdinalIgnoreCase);
                foreach (var shName in sheetKCList)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); }
                    catch { continue; }
                    ExcelHelper.BacaSheetKC(wbSrc, spec, dictCIF);
                }

                var dictRestru = flagRestru
                    ? ExcelHelper.BacaDaftarRestru(wbSrc, "GB0500", "D")
                    : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                wbSrc.Close(false);
                wbSrc = null;

                if (dictCIF.Count == 0)
                    throw new InvalidOperationException("Tidak ada data piutang yang ditemukan.");

                var listCIF = new List<CIFData>(dictCIF.Values);
                SortHelper.SortDescByTotalOS(listCIF, 0, listCIF.Count - 1);

                if (topN < 1)  topN = 10;
                if (topN > listCIF.Count) topN = listCIF.Count;

                BersihkanOutput(wsOut);

                int baris = 5;
                for (int rCif = 0; rCif < topN; rCif++)
                {
                    var cifData = listCIF[rCif];
                    foreach (var kontrak in cifData.Kontrak.Values)
                    {
                        baris++;
                        string adaPN = TentukanPN(kontrak, dictRestru, flagNPF, flagKol2, flagRestru);
                        TulisBaris(wsOut, baris, rCif + 1, cifData.CIF, cifData.Nama, kontrak, adaPN);
                    }
                }

                int totalRows = baris - 5;
                if (totalRows > 0)
                    StyleHelper.StyleTabelCKPN(wsOut, totalRows);
                
                // ================= Audit Log =================
                // NOA per KC dihitung dari SELURUH ruang lingkup (dictCIF), bukan hanya Top-N,
                // supaya bisa mengecek apakah ada KC yang terlewat (NOA = 0).
                var noaPerKC = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var shName in sheetKCList)          // inisialisasi 0 agar KC kosong tetap tampil
                    noaPerKC[shName] = 0;
                foreach (var cd in dictCIF.Values)
                    foreach (var k in cd.Kontrak.Values)
                    {
                        string kc = (k.SheetKC ?? "").Trim();
                        if (noaPerKC.ContainsKey(kc)) noaPerKC[kc]++;
                        else noaPerKC[kc] = 1;               // KC di luar daftar centang (jaga-jaga)
                    }

                int totalNOA = 0;
                foreach (var v in noaPerKC.Values) totalNOA += v;

                double totalOS = 0;                          // OS Top-N (sesuai kolom S)
                for (int r = 0; r < topN; r++) totalOS += listCIF[r].TotalOS;

                // Jangan sampai kegagalan log menggagalkan proses utama
                try { TulisAuditLog(filePath, topN, totalNOA, totalOS, noaPerKC, sheetKCList); }
                catch { /* abaikan error logging */ }
                // =============================================
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
        }

        // ----------------------------------------------------------------
        private static string TentukanPN(
            KontrakData kontrak, HashSet<string> dictRestru,
            bool flagNPF, bool flagKol2, bool flagRestru)
        {
            if (flagRestru && dictRestru.Contains(kontrak.NoKontrak.Trim()))
                return "Ya";

            switch (kontrak.Kualitas)
            {
                case 3: case 4: case 5: return flagNPF  ? "Ya" : "Tidak";
                case 2:                 return flagKol2 ? "Ya" : "Tidak";
                default:                return "Tidak";
            }
        }

        // ----------------------------------------------------------------
        private static void TulisBaris(
            Excel.Worksheet wsOut, int baris, int nomorUrut,
            string cif, string nama, KontrakData kontrak, string adaPN)
        {
            // Cast eksplisit ke Excel.Range wajib untuk setiap akses .Cells[]
            Excel.Range colD = (Excel.Range)wsOut.Cells[baris, "D"];
            Excel.Range colF = (Excel.Range)wsOut.Cells[baris, "F"];
            colD.NumberFormat = "@";
            colF.NumberFormat = "@";

            ((Excel.Range)wsOut.Cells[baris, "B"]).Value2 = nomorUrut;
            ((Excel.Range)wsOut.Cells[baris, "C"]).Value2 = kontrak.SheetKC;
            ((Excel.Range)wsOut.Cells[baris, "D"]).Value2 = "'" + cif;
            ((Excel.Range)wsOut.Cells[baris, "E"]).Value2 = nama;
            ((Excel.Range)wsOut.Cells[baris, "F"]).Value2 = "'" + kontrak.NoKontrak;
            ((Excel.Range)wsOut.Cells[baris, "G"]).Value2 = kontrak.OS;
            ((Excel.Range)wsOut.Cells[baris, "H"]).Value2 = adaPN;
            ((Excel.Range)wsOut.Cells[baris, "I"]).Value2 = kontrak.Jaminan;

            ((Excel.Range)wsOut.Cells[baris, "K"]).Formula =
            "=IF(H" + baris + "=\"Ya\",MAX(0,G" + baris + "-(I" + baris + "-J" + baris + ")),0)";
        }

        // ----------------------------------------------------------------
        private static void BersihkanOutput(Excel.Worksheet wsOut)
        {
            Excel.Range lastCellB = (Excel.Range)wsOut.Cells[wsOut.Rows.Count, "B"];
            Excel.Range endCell   = (Excel.Range)lastCellB.End[Excel.XlDirection.xlUp];
            int lastOut = (int)endCell.Row;
            if (lastOut < 6) lastOut = 6;

            Excel.Range rng = (Excel.Range)wsOut.Range["B6", "K" + lastOut];
            rng.ClearContents();
            rng.ClearFormats();
            try { rng.UnMerge(); } catch { }
        }

        // ----------------------------------------------------------------
        // Tulis ringkasan eksekusi ke sheet "Audit Log" — SATU BARIS PER KC
        // ----------------------------------------------------------------
        private void TulisAuditLog(
            string filePath, int topN, int totalNOA, double totalOS,
            Dictionary<string, int> noaPerKC, string[] sheetKCList)
        {
            var wsLog = CariSheet(_app.ActiveWorkbook, "Audit Log");
            if (wsLog == null) return;

            // Baris terakhir berdasarkan kolom A (data mulai baris 5, header baris 4)
            Excel.Range lastCell = (Excel.Range)wsLog.Cells[wsLog.Rows.Count, "A"];
            Excel.Range endCell  = (Excel.Range)lastCell.End[Excel.XlDirection.xlUp];
            int rowLog = (int)endCell.Row;
            if (rowLog < 4) rowLog = 4;

            double ts = DateTime.Now.ToOADate();   // timestamp sama untuk satu eksekusi

            // Satu baris untuk tiap KC dalam ruang lingkup (urut sesuai daftar centang)
            foreach (var kc in sheetKCList)
            {
                rowLog++;                          // selalu maju -> tidak menumpuk
                int noa = noaPerKC.ContainsKey(kc) ? noaPerKC[kc] : 0;

                Excel.Range cA = (Excel.Range)wsLog.Cells[rowLog, "A"];
                cA.Value2 = ts;
                cA.NumberFormat = "yyyy-mm-dd hh:mm:ss";

                ((Excel.Range)wsLog.Cells[rowLog, "B"]).Value2 = "CKPN Individu - " + kc;
                ((Excel.Range)wsLog.Cells[rowLog, "F"]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[rowLog, "F"]).Value2 = filePath;
                ((Excel.Range)wsLog.Cells[rowLog, "H"]).Value2 = (noa == 0 ? "Kosong / cek" : "Sukses");
                ((Excel.Range)wsLog.Cells[rowLog, "I"]).Value2 = noa;   // #Rek Awal = NOA KC ini
            }

            // Catatan: topN, totalNOA, totalOS sengaja tidak ditulis lagi karena
            // log kini per-KC. Parameter dibiarkan agar pemanggil di Hitung tak berubah.
        }

        // ----------------------------------------------------------------
        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }
    }
}