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

                if (topN < 1) topN = 10;
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
                var noaPerKC = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                foreach (var shName in sheetKCList)
                    noaPerKC[shName] = 0;
                foreach (var cd in dictCIF.Values)
                    foreach (var k in cd.Kontrak.Values)
                    {
                        string kc = (k.SheetKC ?? "").Trim();
                        if (noaPerKC.ContainsKey(kc)) noaPerKC[kc]++;
                        else noaPerKC[kc] = 1;
                    }

                int totalNOA = 0;
                foreach (var v in noaPerKC.Values) totalNOA += v;

                double totalOS = 0;
                for (int r = 0; r < topN; r++) totalOS += listCIF[r].TotalOS;

                try { TulisAuditLog(filePath, topN, totalNOA, totalOS, noaPerKC, sheetKCList); }
                catch { /* abaikan error logging */ }
                // =============================================

                // ============ MODE TAHUNAN (Master!D17) ============
                // Bulanan  -> tidak ada perubahan (logika di atas).
                // Tahunan  -> isi 'B. CKPN - KOL INDV'!U6:U19 dari file Master!D14,
                //             bucket outstanding per hari tunggakan, Top-N dikecualikan.
                if (CekModeTahunan())
                {
                    try { IsiBucketKolektifTahunan(filePath, topN, sheetKCList); }
                    catch { /* jangan gagalkan proses utama */ }
                }
                // ===================================================
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

            Excel.Range lastCell = (Excel.Range)wsLog.Cells[wsLog.Rows.Count, "A"];
            Excel.Range endCell  = (Excel.Range)lastCell.End[Excel.XlDirection.xlUp];
            int rowLog = (int)endCell.Row;
            if (rowLog < 4) rowLog = 4;

            double ts = DateTime.Now.ToOADate();

            foreach (var kc in sheetKCList)
            {
                rowLog++;
                int noa = noaPerKC.ContainsKey(kc) ? noaPerKC[kc] : 0;

                Excel.Range cA = (Excel.Range)wsLog.Cells[rowLog, "A"];
                cA.Value2 = ts;
                cA.NumberFormat = "yyyy-mm-dd hh:mm:ss";

                ((Excel.Range)wsLog.Cells[rowLog, "B"]).Value2 = "CKPN Individu - " + kc;
                ((Excel.Range)wsLog.Cells[rowLog, "F"]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[rowLog, "F"]).Value2 = filePath;
                ((Excel.Range)wsLog.Cells[rowLog, "H"]).Value2 = (noa == 0 ? "Kosong / cek" : "Sukses");
                ((Excel.Range)wsLog.Cells[rowLog, "I"]).Value2 = noa;
            }
        }

        // ================================================================
        // ============ MODE TAHUNAN: isi B. CKPN - KOL INDV!U6:U19 =======
        // ================================================================

        // -------- Switch Master!D17 --------
        private bool CekModeTahunan()
        {
            var ws = CariSheet(_app.ActiveWorkbook, "Master");
            if (ws == null) return false;
            var v = ((Excel.Range)ws.Range["D17"]).Value2;
            string s = (v ?? "").ToString();
            return s.IndexOf("setahun", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // -------- 14 bucket hari tunggakan -> baris kolom U (blok S5:X20) --------
        private static readonly (int Row, double Lo, double Hi)[] BucketsKol =
        {
            (6,  0,   0),      (7,  1,   30),    (8,  31,  60),    (9,  61,  90),
            (10, 91,  120),    (11, 121, 150),   (12, 151, 180),   (13, 181, 210),
            (14, 211, 240),    (15, 241, 270),   (16, 271, 300),   (17, 301, 330),
            (18, 331, 360),    (19, 361, 999999)
        };

        // -------- Isi U6:U19 (Baki Debet kolektif tahunan) --------
        private void IsiBucketKolektifTahunan(string filePath, int topN, string[] sheetKCList)
        {
            if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
                throw new InvalidOperationException("File sumber (Master!D14) tidak ditemukan: " + filePath);

            var wsKol = CariSheet(_app.ActiveWorkbook, "B. CKPN - KOL INDV");
            if (wsKol == null)
                throw new InvalidOperationException("Sheet 'B. CKPN - KOL INDV' tidak ditemukan.");

            var specs = new List<SheetSpec>();
            foreach (var shName in sheetKCList)
                try { specs.Add(SheetSpec.Buat(shName)); } catch { }
            if (specs.Count == 0)
                throw new InvalidOperationException("Tidak ada sheet KC valid.");

            Excel.Workbook srcWb = null;
            try
            {
                srcWb = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                var allData = BacaSemueSheet(srcWb, specs);
                var setSkip = BangunSetExcludeTopN(allData, specs, topN);   // exclude Top-N

                foreach (var b in BucketsKol)
                {
                    double os = SumBucket(allData, specs, setSkip, b.Lo, b.Hi);
                    ((Excel.Range)wsKol.Cells[b.Row, "U"]).Value2 = os;
                }
            }
            finally
            {
                if (srcWb != null) try { srcWb.Close(false); } catch { }
            }
        }

        // ================================================================
        // Mekanisme baca + exclude Top-N (identik dengan CKPNPDNetFlow)
        // ================================================================
        private class RowRecord
        {
            public string CIF, NoRek, Nama;
            public double Hari1, Nom1, Hari2, Nom2;
            public bool   Masuk;
        }

        private Dictionary<string, List<RowRecord>> BacaSemueSheet(
            Excel.Workbook srcWb, List<SheetSpec> specs)
        {
            var result = new Dictionary<string, List<RowRecord>>(StringComparer.OrdinalIgnoreCase);
            foreach (var spec in specs)
                result[spec.SheetName] = BacaSatuSheet(srcWb, spec);
            return result;
        }

        private List<RowRecord> BacaSatuSheet(Excel.Workbook wb, SheetSpec spec)
        {
            var list = new List<RowRecord>();
            if (!ExcelHelper.SheetAda(wb, spec.SheetName)) return list;

            var ws = (Excel.Worksheet)wb.Worksheets[spec.SheetName];
            int headerRow = spec.HeaderRow;
            int lastRow   = ExcelHelper.CariLastRow(ws, spec.ColCIF, headerRow);
            if (lastRow <= headerRow) return list;

            int dataStart = headerRow + 1;
            string[] cifArr  = ExcelHelper.BacaKolomString(ws, spec.ColCIF,   dataStart, lastRow);
            string[] rekArr  = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, dataStart, lastRow);
            string[] namaArr = ExcelHelper.BacaKolomString(ws, spec.ColNama,  dataStart, lastRow);
            double[] h1Arr, n1Arr;
            double[] h2Arr = null, n2Arr = null;

            switch (spec.Mode)
            {
                case SheetMode.AF:
                    n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColOS,     dataStart, lastRow);
                    h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariAF, dataStart, lastRow);
                    break;
                case SheetMode.Tunggakan1:
                    h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    break;
                default: // Tunggakan2
                    h1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    h2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColHariBasil, dataStart, lastRow);
                    n2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomBasil,  dataStart, lastRow);
                    break;
            }

            for (int i = 0; i < cifArr.Length; i++)
            {
                string cif  = (cifArr[i]  ?? "").Trim();
                string rek  = (rekArr[i]  ?? "").Trim();
                string nama = (namaArr[i] ?? "").Trim();
                if (string.IsNullOrEmpty(cif) || string.IsNullOrEmpty(rek)) continue;

                double h1 = h1Arr[i], n1 = n1Arr[i];
                double h2 = h2Arr != null ? h2Arr[i] : 0;
                double n2 = n2Arr != null ? n2Arr[i] : 0;

                bool masuk;
                switch (spec.Mode)
                {
                    case SheetMode.AF:
                        masuk = true;
                        break;
                    case SheetMode.Tunggakan1:
                        masuk = !string.IsNullOrEmpty(nama);
                        break;
                    default: // Tunggakan2
                        masuk = !string.IsNullOrEmpty(nama);
                        break;
                }
                list.Add(new RowRecord { CIF=cif, NoRek=rek, Nama=nama, Hari1=h1, Nom1=n1, Hari2=h2, Nom2=n2, Masuk=masuk });
            }
            return list;
        }

        private HashSet<string> BangunSetExcludeTopN(
            Dictionary<string, List<RowRecord>> allData,
            List<SheetSpec> specs, int topN)
        {
            var setSkip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (topN <= 0) return setSkip;

            var dictOSGlobal = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                {
                    if (!rec.Masuk) continue;
                    double os = HitungOS(rec, spec.Mode);
                    if (dictOSGlobal.ContainsKey(rec.CIF))
                        dictOSGlobal[rec.CIF] += os;
                    else
                        dictOSGlobal[rec.CIF] = os;
                }
            }

            if (dictOSGlobal.Count == 0) return setSkip;

            var pairs = new List<SortHelper.CifOsPair>();
            foreach (var kv in dictOSGlobal)
                pairs.Add(new SortHelper.CifOsPair { CIF = kv.Key, OS = kv.Value });
            SortHelper.SortDescByOS(pairs, 0, pairs.Count - 1);

            int batas = Math.Min(topN, pairs.Count);
            var topCIF = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < batas; j++)
                topCIF.Add(pairs[j].CIF);

            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                    if (rec.Masuk && topCIF.Contains(rec.CIF))
                        setSkip.Add(spec.SheetName + "|" + rec.NoRek);
            }

            return setSkip;
        }

        private double SumBucket(
            Dictionary<string, List<RowRecord>> allData,
            List<SheetSpec> specs, HashSet<string> setSkip,
            double lo, double hi)
        {
            double total = 0;
            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                {
                    if (!rec.Masuk) continue;
                    if (setSkip.Contains(spec.SheetName + "|" + rec.NoRek)) continue;
                    switch (spec.Mode)
                    {
                        case SheetMode.AF:
                        case SheetMode.Tunggakan1:
                            if (rec.Hari1 >= lo && rec.Hari1 <= hi) total += rec.Nom1;
                            break;
                        case SheetMode.Tunggakan2:
                            if (rec.Hari1 >= lo && rec.Hari1 <= hi) total += rec.Nom1;
                            if (rec.Hari2 >= lo && rec.Hari2 <= hi) total += rec.Nom2;
                            break;
                    }
                }
            }
            return total;
        }

        private static double HitungOS(RowRecord rec, SheetMode mode)
        {
            switch (mode)
            {
                case SheetMode.AF:         return rec.Nom1;
                case SheetMode.Tunggakan1: return rec.Nom1;
                case SheetMode.Tunggakan2: return rec.Nom1 + rec.Nom2;
                default:                   return 0;
            }
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