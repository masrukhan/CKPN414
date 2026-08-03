using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using CKPNLibrary.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA modul "module PD Migration".
    /// Menghitung matriks migrasi kualitas piutang antar dua periode triwulan.
    ///
    /// Output: sheet "B2.PD-Migration" baris baseRow..baseRow+4 (5 kualitas)
    /// Staging: sheet "state PD Migration" (dibersihkan lalu diisi ulang)
    /// Log: sheet "Audit Log"
    /// </summary>
    internal class PDMigration
    {
        private readonly Excel.Application _app;

        // Nama sheet wajib di workbook utama
        private const string SheetState  = "state PD Migration";
        private const string SheetTarget = "B2.PD-Migration";
        private const string SheetLog    = "Audit Log";
        private const string SheetMaster = "Master";

        // Layout KC2900: baris data mulai 3, kolom H=NoRek, I=TglHB, L=BakiDebet
        private const int    WoStartRow  = 3;

        public PDMigration(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ----------------------------------------------------------------
        // Entry point — dipanggil dari CKPNFunctions
        // Parameter dikirim dari VBA sebagai string
        // ----------------------------------------------------------------
        public void Hitung(
            string triwulan,      // "Triwulan I" / "II" / "III" / "IV"
            string pathAwal,      // path file periode awal
            string pathAkhir,     // path file periode akhir
            string tglAwalStr,    // tanggal awal WO "yyyyMMdd"
            string tglAkhirStr,   // tanggal akhir WO "yyyyMMdd"
            int    topN,          // Top-N CIF yang di-skip
            string sheetKCList)   // "KC0600,KC0700,..."
        {
            var wb = _app.ActiveWorkbook;

            // Validasi sheet wajib
            var wsState  = CariSheet(wb, SheetState);
            var wsTarget = CariSheet(wb, SheetTarget);
            var wsLog    = CariSheet(wb, SheetLog);

            if (wsState  == null) throw new InvalidOperationException("Sheet '" + SheetState  + "' tidak ditemukan.");
            if (wsTarget == null) throw new InvalidOperationException("Sheet '" + SheetTarget + "' tidak ditemukan.");
            if (wsLog    == null) throw new InvalidOperationException("Sheet '" + SheetLog    + "' tidak ditemukan.");

            // Tentukan baseRow dari triwulan
            int baseRow = TriwulanKeBaseRow(triwulan);

            // Validasi file
            if (!System.IO.File.Exists(pathAwal))
                throw new System.IO.FileNotFoundException("File Awal tidak ditemukan: " + pathAwal);
            if (!System.IO.File.Exists(pathAkhir))
                throw new System.IO.FileNotFoundException("File Akhir tidak ditemukan: " + pathAkhir);

            // Parse periode WO
            long woMin = 0, woMax = 0;
            long.TryParse(tglAwalStr,  out woMin);
            long.TryParse(tglAkhirStr, out woMax);
            if (woMin == 0 || woMax == 0)
                throw new ArgumentException("Periode WO tidak valid: " + tglAwalStr + " - " + tglAkhirStr);

            // Parse daftar sheet KC
            string[] sheets = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);

            // ---- Bersihkan staging sheet ----
            BersihkanState(wsState);

            // ---- Baca file awal (dengan skip Top-N) ----
            var setSkip = BangunSetSkipTopN(pathAwal, sheets, topN);
            AmbilDariFile(pathAwal,  wsState, "B", "C", "D", sheets, setSkip);
            AmbilDariFile(pathAkhir, wsState, "G", "H", "I", sheets, null);

            // ---- Baca WO dari KC2900 file akhir ----
            var dictWO = BacaKC2900(pathAkhir, woMin, woMax);

            // ---- Bangun dictAkhir dari kolom G:I staging ----
            int lastG = CariLastRow(wsState, "G", 2);
            var dictAkhir = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            if (lastG >= 3)
            {
                Excel.Range rngAkhir = (Excel.Range)wsState.Range["G3", "I" + lastG];
                object[,] arrAkhir = (object[,])rngAkhir.Value2;
                for (int k = 1; k <= arrAkhir.GetLength(0); k++)
                {
                    string key = ToStr(arrAkhir[k, 1]);
                    if (!string.IsNullOrEmpty(key) && !dictAkhir.ContainsKey(key))
                        dictAkhir[key] = new[] { ToDouble(arrAkhir[k, 2]), ToDouble(arrAkhir[k, 3]) };
                }
            }

            // ---- Hitung matriks migrasi ----
            int lastB = CariLastRow(wsState, "B", 2);
            double[] totalSaldoAwal = new double[6];   // index 1..5
            double[,] matriks = new double[6, 8];      // [kualAwal 1..5][kolom 1..7]

            object[,] arrAwal = null;
            object[] eVals = null, fVals = null;

            if (lastB >= 3)
            {
                Excel.Range rngAwal = (Excel.Range)wsState.Range["B3", "D" + lastB];
                arrAwal = (object[,])rngAwal.Value2;
                int nRows = arrAwal.GetLength(0);
                eVals = new object[nRows];
                fVals = new object[nRows];

                for (int k = 1; k <= nRows; k++)
                {
                    string key     = ToStr(arrAwal[k, 1]);
                    int    kualAwal = (int)ToDouble(arrAwal[k, 2]);
                    double saldoAwal = ToDouble(arrAwal[k, 3]);

                    if (kualAwal >= 1 && kualAwal <= 5)
                    {
                        totalSaldoAwal[kualAwal] += saldoAwal;

                        if (dictAkhir.ContainsKey(key))
                        {
                            int    kualAkhir  = (int)dictAkhir[key][0];
                            double saldoAkhir = dictAkhir[key][1];
                            eVals[k - 1] = kualAkhir;
                            fVals[k - 1] = saldoAkhir;

                            if (kualAkhir >= 1 && kualAkhir <= 5)
                                matriks[kualAwal, kualAkhir] += saldoAwal;
                            else if (dictWO.ContainsKey(key))
                                matriks[kualAwal, 6] += saldoAwal;   // WO
                            else
                                matriks[kualAwal, 7] += saldoAwal;   // Lainnya
                        }
                        else
                        {
                            if (dictWO.ContainsKey(key))
                            {
                                eVals[k - 1] = "WO";
                                fVals[k - 1] = dictWO[key][0];
                                matriks[kualAwal, 6] += saldoAwal;
                            }
                            else
                            {
                                eVals[k - 1] = 0;
                                fVals[k - 1] = 0;
                                matriks[kualAwal, 7] += saldoAwal;
                            }
                        }
                    }
                    else { eVals[k - 1] = 0; fVals[k - 1] = 0; }
                }

                // Tulis kolom E & F ke staging
                TulisKolom(wsState, "E", 3, eVals);
                TulisKolom(wsState, "F", 3, fVals);
                ((Excel.Range)wsState.Range["D3", "D" + lastB]).NumberFormat = "#,##0;(#,##0);-";
                ((Excel.Range)wsState.Range["F3", "F" + lastB]).NumberFormat = "#,##0;(#,##0);-";
            }
            if (lastG >= 3)
                ((Excel.Range)wsState.Range["I3", "I" + lastG]).NumberFormat = "#,##0;(#,##0);-";

            // ---- Tulis matriks ke B2.PD-Migration ----
            for (int rowK = 1; rowK <= 5; rowK++)
            {
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "B"]).Value2 = totalSaldoAwal[rowK];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "C"]).Value2 = matriks[rowK, 1];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "D"]).Value2 = matriks[rowK, 2];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "E"]).Value2 = matriks[rowK, 3];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "F"]).Value2 = matriks[rowK, 4];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "G"]).Value2 = matriks[rowK, 5];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "H"]).Value2 = matriks[rowK, 6];
                ((Excel.Range)wsTarget.Cells[baseRow + rowK - 1, "I"]).Value2 = matriks[rowK, 7];
            }

            // ---- Hitung OS Individu (rekening yang di-skip) ----
            // Buka kembali file awal untuk hitung total OS dari setSkip
            double osIndividu = HitungOSSetSkip(pathAwal, sheets, setSkip);

            // ---- Audit Log ----
            TulisAuditLog(wsLog, triwulan, tglAwalStr, tglAkhirStr,
                          pathAwal, pathAkhir,
                          lastB - 2, lastG - 2,
                          dictWO, totalSaldoAwal, matriks,
                          topN, setSkip.Count, osIndividu, sheetKCList);

            // ---- Pesan selesai ----
            System.Windows.Forms.MessageBox.Show(
                "Selesai.\n" +
                triwulan + " → B2.PD-Migration baris " + baseRow + ":" + (baseRow + 4) + "\n" +
                "Skip Top-N  : " + topN + " CIF (" + setSkip.Count + " rekening)\n" +
                "OS Individu : " + osIndividu.ToString("N0") + "\n" +
                "Awal        : " + (lastB - 2) + " baris\n" +
                "Akhir       : " + (lastG - 2) + " baris\n" +
                "Rekening WO : " + dictWO.Count,
                "PD Migration", System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        // ================================================================
        // Helpers
        // ================================================================

        private static int TriwulanKeBaseRow(string triwulan)
        {
            switch (triwulan.Trim())
            {
                case "Triwulan I":   return 4;
                case "Triwulan II":  return 23;
                case "Triwulan III": return 42;
                case "Triwulan IV":  return 61;
                default: throw new ArgumentException("Triwulan tidak dikenali: " + triwulan);
            }
        }

        private void BersihkanState(Excel.Worksheet ws)
        {
            BersihkanKolom(ws, "B", "D");
            BersihkanKolom(ws, "G", "I");
            BersihkanKolom(ws, "E", "F");

            // Tulis header ulang
            ((Excel.Range)ws.Range["B1"]).Value2 = "Triwulan Awal";
            ((Excel.Range)ws.Range["G1"]).Value2 = "Triwulan Akhir";
            ((Excel.Range)ws.Range["E1"]).Value2 = "VLOOKUP DATA";
            ((Excel.Range)ws.Range["B2"]).Value2 = "Nomor Rekening";
            ((Excel.Range)ws.Range["C2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["D2"]).Value2 = "Saldo Harga Pokok";
            ((Excel.Range)ws.Range["E2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["F2"]).Value2 = "Saldo Harga Pokok";
            ((Excel.Range)ws.Range["G2"]).Value2 = "Nomor Rekening";
            ((Excel.Range)ws.Range["H2"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["I2"]).Value2 = "Saldo Harga Pokok";
        }

        private void BersihkanKolom(Excel.Worksheet ws, string kolAwal, string kolAkhir)
        {
            int lastA = CariLastRow(ws, kolAwal,  2);
            int lastB = CariLastRow(ws, kolAkhir, 2);
            int last  = Math.Max(lastA, lastB);
            if (last >= 3)
                ((Excel.Range)ws.Range[kolAwal + "3", kolAkhir + last]).ClearContents();
        }

        // Baca file sumber → tulis ke kolom staging (setara AmbilDariFile_Dinamis VBA)
        private void AmbilDariFile(
            string filePath, Excel.Worksheet wsOut,
            string kolRek, string kolKual, string kolPok,
            string[] sheets, HashSet<string> setSkip)
        {
            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                int outRow = 3;

                foreach (var shName in sheets)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); }
                    catch { continue; }

                    if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;
                    var ws = (Excel.Worksheet)wbSrc.Worksheets[shName];

                    int headerRow = spec.HeaderRow;
                    // Gunakan ColCIF (kolom C) untuk cari lastRow — lebih konsisten
                    // karena kolom C (CIF) selalu terisi untuk setiap baris debitur
                    // dan tidak tergantung format per sheet KC
                    int lastRow   = ExcelHelper.CariLastRow(ws, spec.ColCIF, headerRow);
                    if (lastRow <= headerRow) continue;

                    int dataStart = headerRow + 1;
                    string[] namaArr  = ExcelHelper.BacaKolomString(ws, spec.ColNama,     dataStart, lastRow);
                    string[] rekArr   = ExcelHelper.BacaKolomString(ws, spec.ColNoRek,    dataStart, lastRow);
                    string[] kualArr = ExcelHelper.BacaKolomString(ws, spec.ColKualitas, dataStart, lastRow);                    

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

                    for (int i = 0; i < namaArr.Length; i++)
                    {
                        // Filter utama: baris harus punya nama (bukan baris agunan)
                        if (string.IsNullOrEmpty(namaArr[i])) continue;

                        bool masuk; double osVal;
                        switch (spec.Mode)
                        {
                            case SheetMode.AF:
                                // KC0600–KC0900: semua baris berisi nama masuk
                                masuk = true;
                                osVal = n1Arr[i];
                                break;
                            case SheetMode.Tunggakan1:
                                // KC1100 LOGIKA BARU: masuk jika ada nama (sudah dicek di atas)
                                // OS = seluruh AL tanpa syarat hari tunggakan
                                masuk = true;
                                osVal = n1Arr[i];
                                break;
                            default: // Tunggakan2 = KC1000
                                // KC1000 LOGIKA BARU: masuk jika ada nama (sudah dicek di atas)
                                // OS = AL + AN tanpa syarat hari tunggakan
                                masuk = true;
                                osVal = n1Arr[i] + (n2Arr != null ? n2Arr[i] : 0);
                                break;
                        }
                        if (!masuk) continue;

                        string rek = rekArr[i].Trim();
                        if (rek.StartsWith("'")) rek = rek.Substring(1);

                        // Skip Top-N jika berlaku
                        if (setSkip != null && setSkip.Contains(shName + "|" + rek)) continue;
                    
                        int kual = NormKualitas(kualArr[i]);

                        ((Excel.Range)wsOut.Cells[outRow, kolRek ]).Value2 = rek;
                        ((Excel.Range)wsOut.Cells[outRow, kolKual]).Value2 = kual > 0 ? (object)kual : (object)kualArr[i];
                        ((Excel.Range)wsOut.Cells[outRow, kolPok ]).Value2 = osVal;
                        outRow++;
                    }
                }
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
        }

        // Baca KC2900 file akhir → dictWO (setara AmbilDataKC2900 VBA)
        private Dictionary<string, double[]> BacaKC2900(string filePath, long woMin, long woMax)
        {
            var dict = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            if (!System.IO.File.Exists(filePath)) return dict;

            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);
                if (!ExcelHelper.SheetAda(wbSrc, "KC2900")) return dict;

                var ws      = (Excel.Worksheet)wbSrc.Worksheets["KC2900"];
                int lastRow = ExcelHelper.CariLastRow(ws, "H", WoStartRow - 1);
                if (lastRow < WoStartRow) return dict;

                Excel.Range rng  = (Excel.Range)ws.Range["H" + WoStartRow, "L" + lastRow];
                object[,]   arr  = (object[,])rng.Value2;

                for (int i = 1; i <= arr.GetLength(0); i++)
                {
                    string noRek = ToStr(arr[i, 1]);
                    if (string.IsNullOrEmpty(noRek)) continue;

                    long   tglWO = (long)ToDouble(arr[i, 2]);
                    if (tglWO < woMin || tglWO > woMax) continue;

                    double osVal = ToDouble(arr[i, 5]);   // kolom L = indeks 5 dari H
                    if (!dict.ContainsKey(noRek))
                        dict[noRek] = new[] { osVal, tglWO };
                    else
                        dict[noRek][0] += osVal;
                }
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
            return dict;
        }

        // ----------------------------------------------------------------
        // BangunSetSkipTopN: identifikasi rekening Top-N CIF yang dikecualikan
        // (masuk CKPN Individu) dari PD Migration.
        //
        // PERBAIKAN dari versi lama:
        //   1. Ranking LINTAS semua sheet KC (bukan per sheet terpisah)
        //      → konsisten dengan logika CKPNIndividu.cs
        //   2. Filter KC1000/KC1100 menggunakan logika baru:
        //      masuk jika nama debitur ada, OS tanpa syarat hari tunggakan
        //   3. Rekening milik Top-N CIF dikumpulkan dari SEMUA sheet KC
        // ----------------------------------------------------------------
        private HashSet<string> BangunSetSkipTopN(string filePath, string[] sheets, int topN)
        {
            var setSkip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (topN <= 0 || !System.IO.File.Exists(filePath)) return setSkip;

            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                // ── Langkah 1: Agregasi OS per CIF LINTAS SEMUA SHEET ──
                // dictOS[cif] = total OS gabungan dari semua sheet KC
                // recByCIF[cif] = semua "sheetName|noRek" milik CIF ini
                var dictOS   = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
                var recByCIF = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                foreach (var shName in sheets)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); } catch { continue; }
                    if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;

                    var ws      = (Excel.Worksheet)wbSrc.Worksheets[shName];
                    // Pakai ColCIF (kolom C) untuk CariLastRow — konsisten lintas KC
                    int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
                    if (lastRow <= spec.HeaderRow) continue;

                    int dataStart5 = spec.HeaderRow + 1;
                    string[] namaArr = ExcelHelper.BacaKolomString(ws, spec.ColNama,  dataStart5, lastRow);
                    string[] cifArr  = ExcelHelper.BacaKolomString(ws, spec.ColCIF,   dataStart5, lastRow);
                    string[] rekArr  = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, dataStart5, lastRow);
                    double[] h1Arr = new double[namaArr.Length];
                    double[] n1Arr, n2Arr = null;

                    switch (spec.Mode)
                    {
                        case SheetMode.AF:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColOS,        dataStart5, lastRow);
                            break;
                        case SheetMode.Tunggakan1:
                            // KC1100: baca nominal pokok (tidak perlu hari)
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart5, lastRow);
                            break;
                        default: // Tunggakan2 = KC1000
                            // KC1000: baca nominal pokok + basil (tidak perlu hari)
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dataStart5, lastRow);
                            n2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomBasil,  dataStart5, lastRow);
                            break;
                    }

                    for (int i = 0; i < namaArr.Length; i++)
                    {
                        // Filter: harus ada nama debitur (bukan baris agunan)
                        if (string.IsNullOrEmpty(namaArr[i])) continue;

                        double osVal;
                        switch (spec.Mode)
                        {
                            case SheetMode.AF:
                                osVal = n1Arr[i]; break;
                            case SheetMode.Tunggakan1:
                                // KC1100 LOGIKA BARU: OS = seluruh AL tanpa syarat hari
                                osVal = n1Arr[i]; break;
                            default: // KC1000
                                // KC1000 LOGIKA BARU: OS = AL + AN tanpa syarat hari
                                osVal = n1Arr[i] + (n2Arr != null ? n2Arr[i] : 0); break;
                        }

                        string cif = cifArr[i].Trim();
                        string rek = rekArr[i].Trim();
                        if (rek.StartsWith("'")) rek = rek.Substring(1);
                        if (string.IsNullOrEmpty(cif)) continue;

                        if (!dictOS.ContainsKey(cif))
                        {
                            dictOS[cif]   = 0;
                            recByCIF[cif] = new List<string>();
                        }
                        dictOS[cif] += osVal;
                        recByCIF[cif].Add(shName + "|" + rek);
                    }
                }

                if (dictOS.Count == 0)
                {
                    if (wbSrc != null) try { wbSrc.Close(false); } catch { }
                    wbSrc = null;
                    return setSkip;
                }

                // ── Langkah 2: Sort descending lintas sheet, ambil Top-N CIF ──
                var pairs = new List<SortHelper.CifOsPair>();
                foreach (var kv in dictOS)
                    pairs.Add(new SortHelper.CifOsPair { CIF = kv.Key, OS = kv.Value });
                SortHelper.SortDescByOS(pairs, 0, pairs.Count - 1);

                int ambil = Math.Min(topN, pairs.Count);

                // ── Langkah 3: Kumpulkan semua rekening Top-N CIF dari semua sheet ──
                for (int t = 0; t < ambil; t++)
                    foreach (var keyItem in recByCIF[pairs[t].CIF])
                        setSkip.Add(keyItem);
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
            return setSkip;
        }

        // ----------------------------------------------------------------
        // TulisAuditLog: catat ringkasan perhitungan ke sheet Audit Log
        // Parameter tambahan: topN, osIndividu, sheetKCList
        // ----------------------------------------------------------------
        private void TulisAuditLog(
            Excel.Worksheet wsLog,
            string triwulan, string tglAwalStr, string tglAkhirStr,
            string pathAwal, string pathAkhir,
            int nAwal, int nAkhir,
            Dictionary<string, double[]> dictWO,
            double[] totalSaldoAwal, double[,] matriks,
            int topN, int nSkipRek, double osIndividu, string sheetKCList)
        {
            const int summaryStart = 5;
            const int summaryMax   = 200;

            int nextRow = summaryStart;
            while (nextRow <= summaryMax)
            {
                string a = ToStr(((Excel.Range)wsLog.Cells[nextRow, 1]).Value2);
                string b = ToStr(((Excel.Range)wsLog.Cells[nextRow, 2]).Value2);
                if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) break;
                nextRow++;
            }

            double ts = DateTime.Now.ToOADate();
            double totalSA = 0, totalWO = 0;
            for (int i = 1; i <= 5; i++) { totalSA += totalSaldoAwal[i]; totalWO += matriks[i, 6]; }

            // Konversi tglAwalStr / tglAkhirStr dari "yyyyMMdd" ke OADate
            // agar kolom D dan E tampil sebagai tanggal (mis. "1/1/2025")
            // bukan teks angka "20250101"
            double oaAwal  = ParseTglToOADate(tglAwalStr);
            double oaAkhir = ParseTglToOADate(tglAkhirStr);

            ((Excel.Range)wsLog.Cells[nextRow,  1]).Value2 = ts;
            ((Excel.Range)wsLog.Cells[nextRow,  1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow,  2]).Value2 = "PD Migration";
            ((Excel.Range)wsLog.Cells[nextRow,  3]).Value2 = triwulan;
            ((Excel.Range)wsLog.Cells[nextRow,  4]).Value2 = oaAwal  > 0 ? oaAwal  : (object)tglAwalStr;
            ((Excel.Range)wsLog.Cells[nextRow,  4]).NumberFormat = oaAwal  > 0 ? "m/d/yyyy" : "@";
            ((Excel.Range)wsLog.Cells[nextRow,  5]).Value2 = oaAkhir > 0 ? oaAkhir : (object)tglAkhirStr;
            ((Excel.Range)wsLog.Cells[nextRow,  5]).NumberFormat = oaAkhir > 0 ? "m/d/yyyy" : "@";
            ((Excel.Range)wsLog.Cells[nextRow,  6]).Value2 = pathAwal;
            ((Excel.Range)wsLog.Cells[nextRow,  7]).Value2 = pathAkhir;
            ((Excel.Range)wsLog.Cells[nextRow,  8]).Value2 = "OK";
            ((Excel.Range)wsLog.Cells[nextRow,  9]).Value2 = nAwal;
            ((Excel.Range)wsLog.Cells[nextRow, 10]).Value2 = nAkhir;
            ((Excel.Range)wsLog.Cells[nextRow, 11]).Value2 = dictWO.Count;
            ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = totalSA;
            ((Excel.Range)wsLog.Cells[nextRow, 13]).Value2 = topN;
            ((Excel.Range)wsLog.Cells[nextRow, 14]).Value2 = nSkipRek;
            ((Excel.Range)wsLog.Cells[nextRow, 15]).Value2 = osIndividu;
            ((Excel.Range)wsLog.Cells[nextRow, 15]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 16]).Value2 = sheetKCList;
            ((Excel.Range)wsLog.Cells[nextRow, 19]).Value2 = totalWO;

            // Header kolom baru jika baris pertama di area summary
            if (nextRow == summaryStart)
            {
                ((Excel.Range)wsLog.Cells[summaryStart - 1, 13]).Value2 = "Top-N";
                ((Excel.Range)wsLog.Cells[summaryStart - 1, 14]).Value2 = "Rek Skip (Individu)";
                ((Excel.Range)wsLog.Cells[summaryStart - 1, 15]).Value2 = "OS Individu";
                ((Excel.Range)wsLog.Cells[summaryStart - 1, 16]).Value2 = "Ruang Lingkup KC";
            }
        }

        // ----------------------------------------------------------------
        // Utility methods
        // ----------------------------------------------------------------
        private static int CariLastRow(Excel.Worksheet ws, string col, int minRow)
        {
            Excel.Range lastCell = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            Excel.Range endCell  = (Excel.Range)lastCell.End[Excel.XlDirection.xlUp];
            int row = (int)endCell.Row;
            return row < minRow ? minRow : row;
        }

        private static void TulisKolom(Excel.Worksheet ws, string col, int startRow, object[] vals)
        {
            for (int i = 0; i < vals.Length; i++)
                ((Excel.Range)ws.Cells[startRow + i, col]).Value2 = vals[i];
        }

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

        // ----------------------------------------------------------------
        // ParseTglToOADate: konversi string "yyyyMMdd" ke OADate Excel
        // Contoh: "20250101" → DateTime(2025,1,1).ToOADate()
        // Return 0 jika parse gagal (akan fallback ke teks asli)
        // ----------------------------------------------------------------
        private static double ParseTglToOADate(string tglStr)
        {
            if (string.IsNullOrEmpty(tglStr) || tglStr.Length != 8) return 0;
            int yr, mo, dy;
            if (!int.TryParse(tglStr.Substring(0, 4), out yr)) return 0;
            if (!int.TryParse(tglStr.Substring(4, 2), out mo)) return 0;
            if (!int.TryParse(tglStr.Substring(6, 2), out dy)) return 0;
            try
            {
                return new DateTime(yr, mo, dy).ToOADate();
            }
            catch { return 0; }
        }

        private static string ToStr(object val)
        {
            if (val == null) return "";
            return val.ToString().Trim();
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

        // ----------------------------------------------------------------
        // HitungOSSetSkip: hitung total OS rekening yang masuk CKPN Individu
        // Dipanggil setelah BangunSetSkipTopN untuk keperluan Audit Log
        // ----------------------------------------------------------------
        private double HitungOSSetSkip(string filePath, string[] sheets, HashSet<string> setSkip)
        {
            if (setSkip.Count == 0 || !System.IO.File.Exists(filePath)) return 0;

            double totalOS = 0;
            Excel.Workbook wbSrc = null;
            try
            {
                wbSrc = _app.Workbooks.Open(filePath, UpdateLinks: 0, ReadOnly: true);

                foreach (var shName in sheets)
                {
                    SheetSpec spec;
                    try { spec = SheetSpec.Buat(shName); } catch { continue; }
                    if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;

                    var ws      = (Excel.Worksheet)wbSrc.Worksheets[shName];
                    // Pakai ColCIF untuk CariLastRow — konsisten
                    int lastRow = ExcelHelper.CariLastRow(ws, spec.ColCIF, spec.HeaderRow);
                    if (lastRow <= spec.HeaderRow) continue;

                    int dsOS = spec.HeaderRow + 1;
                    string[] namaArr = ExcelHelper.BacaKolomString(ws, spec.ColNama,  dsOS, lastRow);
                    string[] rekArr  = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, dsOS, lastRow);
                    double[] n1Arr, n2Arr = null;

                    switch (spec.Mode)
                    {
                        case SheetMode.AF:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColOS,        dsOS, lastRow);
                            break;
                        case SheetMode.Tunggakan1:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dsOS, lastRow);
                            break;
                        default:
                            n1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomPokok,  dsOS, lastRow);
                            n2Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColNomBasil,  dsOS, lastRow);
                            break;
                    }

                    for (int i = 0; i < namaArr.Length; i++)
                    {
                        if (string.IsNullOrEmpty(namaArr[i])) continue;
                        string rek = rekArr[i].Trim();
                        if (rek.StartsWith("'")) rek = rek.Substring(1);

                        string key = shName + "|" + rek;
                        if (!setSkip.Contains(key)) continue;

                        double osVal;
                        switch (spec.Mode)
                        {
                            case SheetMode.AF:
                            case SheetMode.Tunggakan1:
                                osVal = n1Arr[i]; break;
                            default:
                                osVal = n1Arr[i] + (n2Arr != null ? n2Arr[i] : 0); break;
                        }
                        totalOS += osVal;
                    }
                }
            }
            finally
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
            }
            return totalOS;
        }
    }
}