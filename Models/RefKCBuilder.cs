using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA modul "BangunTabelReferensiKC".
    ///
    /// Fungsi: membangun tabel referensi kode KC di sheet "RefKC" dan
    /// mengisi kolom M (Kode KC) di sheet KC2900 pada setiap file sumber.
    ///
    /// Kolom M KC2900 wajib terisi agar LGD Expected Recoveries bisa
    /// memfilter rekening berdasarkan ruang lingkup KC.
    ///
    /// Alur:
    ///   PASS 1 — buka tiap file READ-ONLY, agregasi 3-digit kode produk
    ///            dari semua sheet KC0600..KC1100. Tentukan kode KC
    ///            mayoritas untuk setiap kode produk.
    ///   PASS 2 — buka tiap file READ-WRITE, isi kolom M di KC2900
    ///            berdasarkan hasil PASS 1, lalu SAVE.
    /// </summary>
    internal class RefKCBuilder
    {
        private readonly Excel.Application _app;

        private const string KC2900_SHEET = "KC2900";
        private const string KC2900_REK   = "H";   // kolom No. Rekening
        private const string KC2900_FLAG  = "M";   // kolom Kode KC (output)
        private const int    KC2900_START = 3;      // baris data mulai

        private const string REF_SHEET    = "RefKC";
        private const string REF_NAMA_COL = "D";   // kolom Nama di KC
        private const string REF_REK_COL  = "J";   // kolom No. Rek di KC
        private const int    REF_DATA_START = 5;   // baris data KC mulai

        private static readonly string[] KC_LIST =
        {
            "KC0600","KC0700","KC0800","KC0900","KC1000","KC1100"
        };

        public RefKCBuilder(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ================================================================
        // Entry point — dipanggil dari LGDExpectedRecoveries sebelum BacaKC2900
        //
        // Parameter:
        //   filePaths  : Dictionary<tahun, path> dari file sumber tahunan
        //   wsLog      : sheet Audit Log untuk mencatat hasil
        //
        // Return:
        //   winner     : Dictionary<prodCode3digit, kodeKC>
        //                Dipakai untuk filter di BacaKC2900
        //
        // Catatan: PASS 2 menulis dan menyimpan file sumber (read-write).
        // ================================================================
        public Dictionary<string, string> BangunRefKC(
            Dictionary<int, string> filePaths,
            Excel.Worksheet wsLog)
        {
            var wb = _app.ActiveWorkbook;

            // ── PASS 1: Agregasi kode produk dari semua file ──────────
            // agg[prodCode][kodeKC] = jumlah kemunculan
            var agg = new Dictionary<string, Dictionary<string, int>>(
                StringComparer.OrdinalIgnoreCase);

            int jmlAda = 0, jmlHilang = 0;
            foreach (var kv in filePaths)
            {
                string path = kv.Value;
                if (!System.IO.File.Exists(path)) { jmlHilang++; continue; }

                jmlAda++;
                Excel.Workbook wbSrc = null;
                try
                {
                    wbSrc = _app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true);
                    AgregasiSatuWorkbook(wbSrc, agg);
                    wbSrc.Close(false); wbSrc = null;
                }
                finally { if (wbSrc != null) try { wbSrc.Close(false); } catch { } }
            }

            // ── Pilih kode KC mayoritas per kode produk ───────────────
            var winner = PilihKodeKCMayoritas(agg);

            // ── Tulis tabel RefKC di workbook utama ───────────────────
            TulisTabelReferensi(wb, winner);

            // ── PASS 2: Isi kolom M di KC2900 setiap file ────────────
            int totRows = 0, totHit = 0, fileOk = 0, fileGagal = 0;

            foreach (var kv in filePaths)
            {
                string path = kv.Value;
                if (!System.IO.File.Exists(path)) { fileGagal++; continue; }

                int rowsThis, hitThis;
                bool okThis = FlagSatuFileKC2900(path, winner, out rowsThis, out hitThis);

                if (okThis)
                {
                    fileOk++;
                    totRows += rowsThis;
                    totHit  += hitThis;
                }
                else fileGagal++;
            }

            // ── Catat ke Audit Log ────────────────────────────────────
            if (wsLog != null)
            {
                try
                {
                    CatatAuditLog(wsLog, filePaths, jmlAda, jmlHilang,
                                  winner.Count, totRows, totHit,
                                  fileOk, fileGagal);
                }
                catch { /* jangan gagalkan proses utama */ }
            }

            return winner;
        }

        // ================================================================
        // CekKolomMTerisi: cek apakah kolom M KC2900 sudah berisi data
        // di semua file sumber. Dipakai untuk putuskan apakah perlu re-run.
        // Return: true jika semua file OK (semua kolom M terisi)
        // ================================================================
        public bool CekKolomMTerisi(Dictionary<int, string> filePaths)
        {
            foreach (var kv in filePaths)
            {
                string path = kv.Value;
                if (!System.IO.File.Exists(path)) continue;

                Excel.Workbook wbSrc = null;
                try
                {
                    wbSrc = _app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true);

                    if (!ExcelHelper.SheetAda(wbSrc, KC2900_SHEET))
                    {
                        wbSrc.Close(false); wbSrc = null;
                        return false;
                    }

                    var wsKC = (Excel.Worksheet)wbSrc.Worksheets[KC2900_SHEET];
                    int lastRow = ExcelHelper.CariLastRow(wsKC, KC2900_REK, KC2900_START - 1);
                    if (lastRow < KC2900_START)
                    {
                        wbSrc.Close(false); wbSrc = null;
                        continue;   // file kosong, skip
                    }

                    // Cek beberapa baris kolom M — jika semua kosong, perlu di-build
                    int cekMax = Math.Min(lastRow, KC2900_START + 9);
                    string[] mArr = ExcelHelper.BacaKolomString(wsKC, KC2900_FLAG, KC2900_START, cekMax);
                    bool adaIsi = false;
                    foreach (var m in mArr)
                        if (!string.IsNullOrEmpty(m)) { adaIsi = true; break; }

                    wbSrc.Close(false); wbSrc = null;

                    if (!adaIsi) return false;
                }
                finally { if (wbSrc != null) try { wbSrc.Close(false); } catch { } }
            }
            return true;
        }

        // ================================================================
        // PASS 1: Agregasi satu workbook
        // agg[prodCode3digit][kodeKC] = jumlah kemunculan
        // ================================================================
        private void AgregasiSatuWorkbook(
            Excel.Workbook wbSrc,
            Dictionary<string, Dictionary<string, int>> agg)
        {
            foreach (var shName in KC_LIST)
            {
                if (!ExcelHelper.SheetAda(wbSrc, shName)) continue;

                var ws = (Excel.Worksheet)wbSrc.Worksheets[shName];
                int lastRowD = ExcelHelper.CariLastRow(ws, REF_NAMA_COL, REF_DATA_START - 1);
                int lastRowJ = ExcelHelper.CariLastRow(ws, REF_REK_COL,  REF_DATA_START - 1);
                int lastRow  = Math.Max(lastRowD, lastRowJ);
                if (lastRow < REF_DATA_START) continue;

                string[] namaArr = ExcelHelper.BacaKolomString(ws, REF_NAMA_COL, REF_DATA_START, lastRow);
                string[] rekArr  = ExcelHelper.BacaKolomString(ws, REF_REK_COL,  REF_DATA_START, lastRow);

                for (int i = 0; i < namaArr.Length; i++)
                {
                    string nama  = (namaArr[i] ?? "").Trim();
                    string noRek = (rekArr[i]  ?? "").Trim();

                    // Sama dengan VBA: skip jika nama kosong atau noRek < 3 char
                    if (string.IsNullOrEmpty(nama) || noRek.Length < 3) continue;

                    string prodCode = noRek.Substring(0, 3);

                    if (!agg.ContainsKey(prodCode))
                        agg[prodCode] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                    if (agg[prodCode].ContainsKey(shName))
                        agg[prodCode][shName]++;
                    else
                        agg[prodCode][shName] = 1;
                }
            }
        }

        // ================================================================
        // Pilih kode KC mayoritas per kode produk
        // Jika jumlah sama, pilih kode KC yang lebih kecil (alfabetis)
        // ================================================================
        private Dictionary<string, string> PilihKodeKCMayoritas(
            Dictionary<string, Dictionary<string, int>> agg)
        {
            var winner = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var kvProd in agg)
            {
                string bestKode = "";
                int    bestCnt  = -1;

                foreach (var kvKC in kvProd.Value)
                {
                    if (kvKC.Value > bestCnt ||
                        (kvKC.Value == bestCnt &&
                         string.Compare(kvKC.Key, bestKode, StringComparison.OrdinalIgnoreCase) < 0))
                    {
                        bestCnt  = kvKC.Value;
                        bestKode = kvKC.Key;
                    }
                }
                if (!string.IsNullOrEmpty(bestKode))
                    winner[kvProd.Key] = bestKode;
            }
            return winner;
        }

        // ================================================================
        // Tulis tabel RefKC di sheet "RefKC" workbook utama
        // ================================================================
        private void TulisTabelReferensi(
            Excel.Workbook wb,
            Dictionary<string, string> winner)
        {
            // Buat/clear sheet RefKC
            Excel.Worksheet ws = null;
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (sh.Name.Equals(REF_SHEET, StringComparison.OrdinalIgnoreCase))
                { ws = sh; break; }

            if (ws == null)
            {
                ws = (Excel.Worksheet)wb.Worksheets.Add(
                    After: wb.Worksheets[wb.Worksheets.Count]);
                ws.Name = REF_SHEET;
            }

            ws.Cells.Clear();

            // Header
            ((Excel.Range)ws.Range["A1"]).Value2 = "Kode Produk (3 digit)";
            ((Excel.Range)ws.Range["B1"]).Value2 = "Kode KC";
            Excel.Range hdrRng = (Excel.Range)ws.Range["A1", "B1"];
            hdrRng.Font.Bold          = true;
            hdrRng.Font.Color         = 0xFFFFFF;
            hdrRng.Interior.Color     = 0x436E1F;
            hdrRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            if (winner.Count == 0)
            {
                ((Excel.Range)ws.Range["A2"]).Value2 = "(tidak ada data)";
                return;
            }

            // Sort kode produk ascending
            var keys = new List<string>(winner.Keys);
            keys.Sort(StringComparer.OrdinalIgnoreCase);

            int n = keys.Count;
            var outArr = new object[n, 2];
            for (int i = 0; i < n; i++)
            {
                outArr[i, 0] = keys[i];
                outArr[i, 1] = winner[keys[i]];
            }

            ((Excel.Range)ws.Range["A2"]).NumberFormat = "@";
            Excel.Range dataRng = (Excel.Range)ws.Range["A2", "B" + (n + 1)];
            dataRng.Value2 = outArr;
            dataRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ((Excel.Range)ws.Columns["A"]).AutoFit();
            ((Excel.Range)ws.Columns["B"]).AutoFit();
        }

        // ================================================================
        // PASS 2: Isi kolom M di KC2900 untuk satu file
        // File dibuka READ-WRITE, di-save, lalu ditutup
        // ================================================================
        private bool FlagSatuFileKC2900(
            string fullPath,
            Dictionary<string, string> winner,
            out int rowsThis, out int hitThis)
        {
            rowsThis = 0; hitThis = 0;

            Excel.Workbook wbSrc = null;
            try
            {
                // Buka READ-WRITE (bukan ReadOnly)
                wbSrc = _app.Workbooks.Open(fullPath, UpdateLinks: 0, ReadOnly: false);

                if (!ExcelHelper.SheetAda(wbSrc, KC2900_SHEET))
                {
                    wbSrc.Close(false); wbSrc = null;
                    return false;
                }

                var wsKC    = (Excel.Worksheet)wbSrc.Worksheets[KC2900_SHEET];
                int lastRow = ExcelHelper.CariLastRow(wsKC, KC2900_REK, KC2900_START - 1);

                if (lastRow < KC2900_START)
                {
                    wbSrc.Close(false); wbSrc = null;
                    return true;  // file kosong, tapi bukan error
                }

                // Tulis header kolom M
                ((Excel.Range)wsKC.Range[KC2900_FLAG + "1"]).Value2 = "Referensi";
                ((Excel.Range)wsKC.Range[KC2900_FLAG + "2"]).Value2 = "Kode KC";
                Excel.Range mHdr = (Excel.Range)wsKC.Range[KC2900_FLAG + "1", KC2900_FLAG + "2"];
                mHdr.Font.Bold          = true;
                mHdr.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Baca kolom H (No. Rekening)
                int n = lastRow - KC2900_START + 1;
                string[] rekArr = ExcelHelper.BacaKolomString(
                    wsKC, KC2900_REK, KC2900_START, lastRow);

                // Bangun array output kolom M
                var outArr = new object[n, 1];
                int cnt = 0;
                for (int i = 0; i < rekArr.Length; i++)
                {
                    string noRek = (rekArr[i] ?? "").Trim();
                    if (noRek.Length >= 3)
                    {
                        string prod = noRek.Substring(0, 3);
                        if (winner.ContainsKey(prod))
                        {
                            outArr[i, 0] = winner[prod];
                            cnt++;
                        }
                        else outArr[i, 0] = "";
                    }
                    else outArr[i, 0] = "";
                }

                // Tulis bulk ke kolom M
                Excel.Range mRng = (Excel.Range)wsKC.Range[
                    KC2900_FLAG + KC2900_START,
                    KC2900_FLAG + lastRow];
                mRng.Value2 = outArr;

                // Save file
                wbSrc.Save();
                wbSrc.Close(false); wbSrc = null;

                rowsThis = n;
                hitThis  = cnt;
                return true;
            }
            catch
            {
                if (wbSrc != null) try { wbSrc.Close(false); } catch { }
                wbSrc = null;
                return false;
            }
        }

        // ================================================================
        // Catat ringkasan ke Audit Log
        // ================================================================
        private void CatatAuditLog(
            Excel.Worksheet wsLog,
            Dictionary<int, string> filePaths,
            int jmlAda, int jmlHilang, int jmlKodeProduk,
            int totRows, int totHit, int fileOk, int fileGagal)
        {
            // Cari baris kosong berikutnya
            int r = 5;
            for (int row = 5; row <= 1000000; row++)
            {
                string a = (((Excel.Range)wsLog.Cells[row, 1]).Value2 ?? "").ToString().Trim();
                string b = (((Excel.Range)wsLog.Cells[row, 2]).Value2 ?? "").ToString().Trim();
                if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) { r = row; break; }
            }

            string status = fileGagal == 0 && totHit > 0 ? "SUKSES"
                          : fileOk   == 0               ? "GAGAL"
                          :                               "SEBAGIAN";

            // Ambil path file pertama untuk kolom G
            string pathPertama = "";
            foreach (var kv in filePaths) { pathPertama = kv.Value; break; }

            ((Excel.Range)wsLog.Cells[r, 1]).Value2 = DateTime.Now.ToOADate();
            ((Excel.Range)wsLog.Cells[r, 1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[r, 2]).Value2 = "Bangun RefKC";
            ((Excel.Range)wsLog.Cells[r, 3]).Value2 = "KC0600,KC0700,KC0800,KC0900,KC1000,KC1100";
            ((Excel.Range)wsLog.Cells[r, 6]).Value2 = "Master!D60:D65";
            ((Excel.Range)wsLog.Cells[r, 7]).Value2 = pathPertama;
            ((Excel.Range)wsLog.Cells[r, 8]).Value2 = status;
            ((Excel.Range)wsLog.Cells[r, 9]).Value2 = filePaths.Count;     // jumlah path
            ((Excel.Range)wsLog.Cells[r,10]).Value2 = fileOk;              // file OK
            ((Excel.Range)wsLog.Cells[r,11]).Value2 = jmlHilang;           // file hilang
            ((Excel.Range)wsLog.Cells[r,18]).Value2 = jmlKodeProduk;       // kolom R = kode produk unik
        }
    }
}