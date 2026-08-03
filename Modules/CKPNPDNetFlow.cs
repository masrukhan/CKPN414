using System;
using System.Collections.Generic;
using System.Windows.Forms;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    internal class CKPNPDNetFlow
    {
        private readonly Excel.Application _app;

        private static readonly (int Row, double Lo, double Hi)[] Buckets =
        {
            (4,  0,      0),
            (5,  1,      30),
            (6,  31,     60),
            (7,  61,     90),
            (8,  91,     120),
            (9,  121,    150),
            (10, 151,    180),
            (11, 181,    210),
            (12, 211,    240),
            (13, 241,    270),
            (14, 271,    300),
            (15, 301,    330),
            (16, 331,    360),
            (18, 361,    999999)
        };

        private const string WoSheet    = "KC2900";
        private const string WoCol      = "L";
        private const int    WoStartRow = 3;

        public CKPNPDNetFlow(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ----------------------------------------------------------------
        public void IsiBucket(
            string[] filePaths, string[] bulanLabels, string[] refDates,
            int topN, string[] sheetKCList)
        {
            var wb = _app.ActiveWorkbook;

            var wsTarget = CariSheet(wb, "B1.PD-Net Flow");
            if (wsTarget == null)
                throw new InvalidOperationException("Sheet 'B1.PD-Net Flow' tidak ditemukan.");

            var wsLog = CariSheet(wb, "Audit Log");
            if (wsLog == null)
                throw new InvalidOperationException("Sheet 'Audit Log' tidak ditemukan.");

            var specs = new List<SheetSpec>();
            foreach (var shName in sheetKCList)
                try { specs.Add(SheetSpec.Buat(shName)); } catch { }

            if (specs.Count == 0)
                throw new InvalidOperationException("Tidak ada sheet KC valid di sheetKCList.");

            int logRow = NextAuditLogRow(wsLog);
            TulisAuditLogHeader(wsLog, specs);

            double runStamp = DateTime.Now.ToOADate();
            int okCount = 0, skipCount = 0;
            var errLines = new System.Text.StringBuilder();

            for (int mIdx = 0; mIdx < filePaths.Length; mIdx++)
            {
                int    targetColIdx = 4 + mIdx;
                string targetCol    = ColIndexToName(targetColIdx);
                string filePath     = filePaths[mIdx].Trim();
                string bulanLabel   = mIdx < bulanLabels.Length ? bulanLabels[mIdx] : "";
                string refDateStr   = mIdx < refDates.Length    ? refDates[mIdx]    : "";

                double totalOS = 0, woSum = 0, totalEAD = 0, osIndividu = 0, osNonIndividu = 0;
                var    osPerSheet = new double[specs.Count];
                string status     = "";

                if (string.IsNullOrEmpty(filePath) ||
                    filePath.Equals("pilih folder", StringComparison.OrdinalIgnoreCase))
                {
                    status = "Skipped: path kosong"; skipCount++;
                    TulisAuditLog(wsLog, logRow++, runStamp, bulanLabel, refDateStr, filePath,
                                  status, osPerSheet, totalOS, woSum, totalEAD, topN, osIndividu, osNonIndividu);
                    continue;
                }

                if (!System.IO.File.Exists(filePath))
                {
                    status = "File tidak ditemukan";
                    errLines.AppendLine("- File tidak ditemukan: " + filePath);
                    skipCount++;
                    TulisAuditLog(wsLog, logRow++, runStamp, bulanLabel, refDateStr, filePath,
                                  status, osPerSheet, totalOS, woSum, totalEAD, topN, osIndividu, osNonIndividu);
                    continue;
                }

                Excel.Workbook srcWb = null;
                try
                {
                    srcWb = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                    bool sheetMissing = false;
                    foreach (var spec in specs)
                        if (!ExcelHelper.SheetAda(srcWb, spec.SheetName))
                        {
                            errLines.AppendLine("- Sheet '" + spec.SheetName + "' tidak ada di " + srcWb.Name);
                            sheetMissing = true; break;
                        }

                    if (!sheetMissing)
                    {
                        var allData = BacaSemueSheet(srcWb, specs);
                        var setSkip = BangunSetExcludeTopN(allData, specs, topN);

                        foreach (var bucket in Buckets)
                        {
                            Excel.Range cell = (Excel.Range)wsTarget.Cells[bucket.Row, targetCol];
                            cell.Value2 = SumBucket(allData, specs, setSkip, bucket.Lo, bucket.Hi);
                        }

                        if (ExcelHelper.SheetAda(srcWb, WoSheet))
                        {
                            var wsWo = (Excel.Worksheet)srcWb.Worksheets[WoSheet];
                            int woLastRow = ExcelHelper.CariLastRow(wsWo, WoCol, WoStartRow - 1);
                            if (woLastRow >= WoStartRow)
                            {
                                Excel.Range woRng = (Excel.Range)wsWo.Range[WoCol + WoStartRow, WoCol + woLastRow];
                                woSum = (double)wsWo.Application.WorksheetFunction.Sum(woRng);
                            }
                        }
                        ((Excel.Range)wsTarget.Cells[19, targetCol]).Value2 = woSum;
                        ((Excel.Range)wsTarget.Cells[17, targetCol]).Formula = "=" + targetCol + "18+" + targetCol + "19";

                        for (int si = 0; si < specs.Count; si++)
                        {
                            osPerSheet[si] = SumOSSheet(allData, specs[si], setSkip);
                            totalOS += osPerSheet[si];
                        }
                        osIndividu    = SumOSIndividu(allData, specs, setSkip);
                        osNonIndividu = SumOSNonIndividu(allData, specs, setSkip);
                        totalEAD = totalOS + woSum;
                        status = "OK"; okCount++;
                    }
                    else { status = "Sheet KC missing"; }
                }
                catch (Exception ex)
                {
                    status = "Error: " + ex.Message;
                    errLines.AppendLine("- " + filePath + " → " + ex.Message);
                    skipCount++;
                }
                finally
                {
                    if (srcWb != null) try { srcWb.Close(false); } catch { }
                }

                TulisAuditLog(wsLog, logRow++, runStamp, bulanLabel, refDateStr, filePath,
                              status, osPerSheet, totalOS, woSum, totalEAD, topN, osIndividu, osNonIndividu);
            }

            string msg = "Selesai. Bulan terisi: " + okCount + ". Dilewati: " + skipCount + ".";
            if (errLines.Length > 0) msg += "\n\nCatatan:\n" + errLines.ToString();
            MessageBox.Show(msg, "PD Net Flow", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ================================================================
        // Data record per baris
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
            // namaArr dibaca untuk filter KC1000/KC1100: baris tanpa nama = baris agunan
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
                        // KC1100 LOGIKA BARU: masuk jika ada nama debitur
                        // Baris tanpa nama = baris agunan → di-skip
                        // Hari tunggakan (h1) tetap diisi untuk bucket PD Net Flow
                        masuk = !string.IsNullOrEmpty(nama);
                        break;
                    default: // Tunggakan2 = KC1000
                        // KC1000 LOGIKA BARU: masuk jika ada nama debitur
                        // Baris tanpa nama = baris agunan → di-skip
                        // Hari tunggakan tetap diisi untuk bucket PD Net Flow
                        masuk = !string.IsNullOrEmpty(nama);
                        break;
                }
                list.Add(new RowRecord { CIF=cif, NoRek=rek, Nama=nama, Hari1=h1, Nom1=n1, Hari2=h2, Nom2=n2, Masuk=masuk });
            }
            return list;
        }

        // ----------------------------------------------------------------
        // BangunSetExcludeTopN: identifikasi rekening Top-N CIF yang masuk
        // CKPN Individu sehingga bisa dikecualikan dari PD Net Flow.
        //
        // PENTING: logika ranking HARUS sama persis dengan CKPNIndividu.cs:
        //   → Agregasi OS per CIF LINTAS SEMUA SHEET terlebih dulu
        //   → Baru ranking descending, ambil Top-N
        //   → Semua rekening milik CIF Top-N di semua sheet masuk setSkip
        //
        // Sebelumnya (SALAH): ranking per sheet KC secara terpisah
        //   → menghasilkan OS Individu berbeda dengan total CKPN Individu
        //   → CIF yang sama bisa masuk Top-N KC0600 tapi tidak KC0700
        //   → angka di Audit Log "OS Individu" tidak cocok dengan output CKPN Individu
        // ----------------------------------------------------------------
        private HashSet<string> BangunSetExcludeTopN(
            Dictionary<string, List<RowRecord>> allData,
            List<SheetSpec> specs, int topN)
        {
            var setSkip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (topN <= 0) return setSkip;

            // ── Langkah 1: Agregasi OS per CIF LINTAS SEMUA SHEET ──
            // Sama persis dengan CIFData.TambahKontrak di CKPNIndividu:
            // satu CIF bisa punya rekening di KC0600 DAN KC0700 — OS-nya dijumlah.
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

            // ── Langkah 2: Sort descending, ambil Top-N CIF ──
            var pairs = new List<SortHelper.CifOsPair>();
            foreach (var kv in dictOSGlobal)
                pairs.Add(new SortHelper.CifOsPair { CIF = kv.Key, OS = kv.Value });
            SortHelper.SortDescByOS(pairs, 0, pairs.Count - 1);

            int batas = Math.Min(topN, pairs.Count);
            var topCIF = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int j = 0; j < batas; j++)
                topCIF.Add(pairs[j].CIF);

            // ── Langkah 3: Kumpulkan SEMUA rekening milik Top-N CIF ──
            // dari SEMUA sheet KC — bukan hanya sheet tempat CIF pertama ditemukan
            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                    if (rec.Masuk && topCIF.Contains(rec.CIF))
                        setSkip.Add(spec.SheetName + "|" + rec.NoRek);
            }

            return setSkip;
        }

        // ----------------------------------------------------------------
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

        private double SumOSSheet(
            Dictionary<string, List<RowRecord>> allData,
            SheetSpec spec, HashSet<string> setSkip)
        {
            if (!allData.ContainsKey(spec.SheetName)) return 0;
            double total = 0;
            foreach (var rec in allData[spec.SheetName])
            {
                if (!rec.Masuk) continue;
                if (setSkip.Contains(spec.SheetName + "|" + rec.NoRek)) continue;
                total += HitungOS(rec, spec.Mode);
            }
            return total;
        }

        private double SumOSIndividu(
            Dictionary<string, List<RowRecord>> allData,
            List<SheetSpec> specs, HashSet<string> setSkip)
        {
            double total = 0;
            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                {
                    if (!rec.Masuk) continue;
                    if (!setSkip.Contains(spec.SheetName + "|" + rec.NoRek)) continue;
                    total += HitungOS(rec, spec.Mode);
                }
            }
            return total;
        }

        private double SumOSNonIndividu(
            Dictionary<string, List<RowRecord>> allData,
            List<SheetSpec> specs, HashSet<string> setSkip)
        {
            double total = 0;
            foreach (var spec in specs)
            {
                if (!allData.ContainsKey(spec.SheetName)) continue;
                foreach (var rec in allData[spec.SheetName])
                {
                    if (!rec.Masuk) continue;
                    if (setSkip.Contains(spec.SheetName + "|" + rec.NoRek)) continue;
                    total += HitungOS(rec, spec.Mode);
                }
            }
            return total;
        }

        private static double HitungOS(RowRecord rec, SheetMode mode)
        {
            switch (mode)
            {
                case SheetMode.AF:
                    return rec.Nom1;
                case SheetMode.Tunggakan1:
                    // KC1100 LOGIKA BARU: jumlah seluruh tunggakan pokok (AL)
                    // tanpa syarat hari — untuk keperluan OS Individu & BangunSetExcludeTopN
                    return rec.Nom1;
                case SheetMode.Tunggakan2:
                    // KC1000 LOGIKA BARU: jumlah seluruh tunggakan pokok (AL) + basil (AN)
                    // tanpa syarat hari — untuk keperluan OS Individu & BangunSetExcludeTopN
                    return rec.Nom1 + rec.Nom2;
                default:
                    return 0;
            }
        }

        // ----------------------------------------------------------------
        // Audit Log
        // ----------------------------------------------------------------
        private void TulisAuditLogHeader(Excel.Worksheet wsLog, List<SheetSpec> specs)
        {
            const int startCol = 14;
            ((Excel.Range)wsLog.Range[wsLog.Cells[4, 14], wsLog.Cells[4, 29]]).ClearContents();

            for (int i = 0; i < specs.Count; i++)
                ((Excel.Range)wsLog.Cells[4, startCol + i]).Value2 = "OS " + specs[i].SheetName;

            int baseCol = startCol + specs.Count;
            ((Excel.Range)wsLog.Cells[4, baseCol    ]).Value2 = "Total OS";
            ((Excel.Range)wsLog.Cells[4, baseCol + 1]).Value2 = "WO KC2900";
            ((Excel.Range)wsLog.Cells[4, baseCol + 2]).Value2 = "Total EAD";
            ((Excel.Range)wsLog.Cells[4, baseCol + 3]).Value2 = "Top-N Debitur";
            ((Excel.Range)wsLog.Cells[4, baseCol + 4]).Value2 = "OS Individu (Top-N)";
            ((Excel.Range)wsLog.Cells[4, baseCol + 5]).Value2 = "OS Non-Individu (PD Net Flow)";

            Excel.Range hdrRange = (Excel.Range)wsLog.Range[
                wsLog.Cells[4, baseCol + 3], wsLog.Cells[4, baseCol + 5]];
            hdrRange.Font.Bold            = true;
            hdrRange.Font.Color           = 0xFFFFFF;
            hdrRange.Interior.Color       = 0x436E1F;
            hdrRange.HorizontalAlignment  = Excel.XlHAlign.xlHAlignCenter;
        }

        private void TulisAuditLog(
            Excel.Worksheet wsLog, int r, double ts,
            string bulanLabel, string refDate, string filePath, string status,
            double[] osPerSheet, double totalOS, double woSum, double totalEAD,
            int topN, double osIndividu, double osNonIndividu)
        {
            ((Excel.Range)wsLog.Cells[r, 1]).Value2 = ts;
            ((Excel.Range)wsLog.Cells[r, 2]).Value2 = "PD Net Flow";
            ((Excel.Range)wsLog.Cells[r, 3]).Value2 = bulanLabel;
            ((Excel.Range)wsLog.Cells[r, 4]).Value2 = refDate;
            ((Excel.Range)wsLog.Cells[r, 7]).Value2 = filePath;
            ((Excel.Range)wsLog.Cells[r, 8]).Value2 = status;

            const int startCol = 14;
            for (int i = 0; i < osPerSheet.Length; i++)
                ((Excel.Range)wsLog.Cells[r, startCol + i]).Value2 = osPerSheet[i];

            int baseCol = startCol + osPerSheet.Length;
            ((Excel.Range)wsLog.Cells[r, baseCol    ]).Value2 = totalOS;
            ((Excel.Range)wsLog.Cells[r, baseCol + 1]).Value2 = woSum;
            ((Excel.Range)wsLog.Cells[r, baseCol + 2]).Value2 = totalEAD;
            ((Excel.Range)wsLog.Cells[r, baseCol + 3]).Value2 = topN;
            ((Excel.Range)wsLog.Cells[r, baseCol + 4]).Value2 = osIndividu;
            ((Excel.Range)wsLog.Cells[r, baseCol + 5]).Value2 = osNonIndividu;
        }

        private int NextAuditLogRow(Excel.Worksheet wsLog)
        {
            for (int r = 5; r <= 100000; r++)
            {
                string a = (((Excel.Range)wsLog.Cells[r, 1]).Value2 ?? "").ToString().Trim();
                string b = (((Excel.Range)wsLog.Cells[r, 2]).Value2 ?? "").ToString().Trim();
                if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return r;
                if (string.IsNullOrEmpty(a) &&
                    b.IndexOf("DETAIL", StringComparison.OrdinalIgnoreCase) >= 0) return r;
            }
            return 100001;
        }

        // ----------------------------------------------------------------
        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }

        private static string ColIndexToName(int colIdx)
        {
            string name = "";
            while (colIdx > 0)
            {
                colIdx--;
                name = (char)('A' + colIdx % 26) + name;
                colIdx /= 26;
            }
            return name;
        }
    }
}