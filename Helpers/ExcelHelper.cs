using System;
using System.Collections.Generic;
using CKPNLibrary.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    internal static class ExcelHelper
    {
        // ----------------------------------------------------------------
        // Konversi nama kolom Excel ("A"–"ZZ") ke indeks 1-based
        // ----------------------------------------------------------------
        public static int KolIndex(string colName)
        {
            colName = colName.ToUpper().Trim();
            int result = 0;
            foreach (char c in colName)
                result = result * 26 + (c - 'A' + 1);
            return result;
        }

        // ----------------------------------------------------------------
        // BacaRangeAman: baca Range.Value2 ke object[,] secara aman
        //
        // Masalah: Excel mengembalikan tipe berbeda tergantung ukuran range:
        //   Range multi-baris  → object[,]   (index base 1)
        //   Range 1 baris/cell → nilai tunggal: string, double, int, null, dll
        //
        // Solusi: normalisasi semua kasus ke object[,] 1-based.
        // ----------------------------------------------------------------
        public static object[,] BacaRangeAman(Excel.Range rng, int nRows)
        {
            var result = new object[nRows + 1, 2];   // index 1..nRows, kolom 1

            if (nRows <= 0) return result;

            object raw = rng.Value2;

            if (raw == null)
                return result;   // semua null/kosong

            // Kasus multi-baris: sudah object[,] dengan base-index 1
            if (raw is object[,])
            {
                return (object[,])raw;
            }

            // Kasus 1 baris/1 cell: nilai tunggal — bungkus ke array
            // Tipe yang mungkin: string, double, int, bool, DateTime (OADate)
            result[1, 1] = raw;
            return result;
        }

        // ----------------------------------------------------------------
        // Baca satu kolom sebagai array string (bulk read, aman 1 baris)
        // ----------------------------------------------------------------
        public static string[] BacaKolomString(Excel.Worksheet ws,
            string col, int startRow, int lastRow)
        {
            if (lastRow < startRow) return new string[0];
            int n = lastRow - startRow + 1;

            Excel.Range rng = (Excel.Range)ws.Range[col + startRow, col + lastRow];
            object[,] raw   = BacaRangeAman(rng, n);

            var result = new string[n];
            for (int i = 0; i < n; i++)
                result[i] = ValToString(raw[i + 1, 1]);
            return result;
        }

        // ----------------------------------------------------------------
        // Baca satu kolom sebagai array double (bulk read, aman 1 baris)
        // ----------------------------------------------------------------
        public static double[] BacaKolomDouble(Excel.Worksheet ws,
            string col, int startRow, int lastRow)
        {
            if (lastRow < startRow) return new double[0];
            int n = lastRow - startRow + 1;

            Excel.Range rng = (Excel.Range)ws.Range[col + startRow, col + lastRow];
            object[,] raw   = BacaRangeAman(rng, n);

            var result = new double[n];
            for (int i = 0; i < n; i++)
                result[i] = ValToDouble(raw[i + 1, 1]);
            return result;
        }

        // ----------------------------------------------------------------
        // Cari lastRow berdasarkan kolom tertentu
        // ----------------------------------------------------------------
        public static int CariLastRow(Excel.Worksheet ws, string col, int headerRow)
        {
            Excel.Range lastCell = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            Excel.Range endCell  = (Excel.Range)lastCell.End[Excel.XlDirection.xlUp];
            return (int)endCell.Row;
        }

        // ----------------------------------------------------------------
        // Cek apakah sheet ada di workbook
        // ----------------------------------------------------------------
        public static bool SheetAda(Excel.Workbook wb, string sheetName)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        // ----------------------------------------------------------------
        // Baca daftar nomor rekening restrukturisasi dari sheet GB0500
        // ----------------------------------------------------------------
        public static HashSet<string> BacaDaftarRestru(Excel.Workbook wb,
            string sheetName, string col)
        {
            var hasil = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!SheetAda(wb, sheetName)) return hasil;

            var ws      = (Excel.Worksheet)wb.Worksheets[sheetName];
            const int headerRow = 4;
            int lastRow = CariLastRow(ws, col, headerRow);
            if (lastRow <= headerRow) return hasil;

            string[] noRekArr = BacaKolomString(ws, col, headerRow + 1, lastRow);
            foreach (var nr in noRekArr)
                if (!string.IsNullOrEmpty(nr) && !hasil.Contains(nr))
                    hasil.Add(nr);

            return hasil;
        }

        // ----------------------------------------------------------------
        // Baca satu sheet KC → isi dictCIF
        // ----------------------------------------------------------------
        public static Dictionary<string, CIFData> BacaSheetKC(
            Excel.Workbook wb,
            SheetSpec spec,
            Dictionary<string, CIFData> dictCIF = null)
        {
            if (dictCIF == null)
                dictCIF = new Dictionary<string, CIFData>(StringComparer.OrdinalIgnoreCase);

            if (!SheetAda(wb, spec.SheetName)) return dictCIF;

            var ws      = (Excel.Worksheet)wb.Worksheets[spec.SheetName];
            int headerRow = spec.HeaderRow;
            int lastRow   = CariLastRow(ws, spec.ColCIF, headerRow);
            if (lastRow <= headerRow) return dictCIF;

            int dataStart = headerRow + 1;
            int n = lastRow - dataStart + 1;

            string[] cifArr   = BacaKolomString(ws, spec.ColCIF,      dataStart, lastRow);
            string[] namaArr  = BacaKolomString(ws, spec.ColNama,     dataStart, lastRow);
            string[] noRekArr = BacaKolomString(ws, spec.ColNoRek,    dataStart, lastRow);
            double[] jamArr   = BacaKolomDouble(ws, spec.ColJaminan,  dataStart, lastRow);
            double[] kualArr  = BacaKolomDouble(ws, spec.ColKualitas, dataStart, lastRow);

            double[] hariPokok = null, nomPokok = null, hariBasil = null, nomBasil = null;
            double[] osArr = null, hariAF = null;

            switch (spec.Mode)
            {
                case SheetMode.AF:
                    osArr  = BacaKolomDouble(ws, spec.ColOS,     dataStart, lastRow);
                    hariAF = BacaKolomDouble(ws, spec.ColHariAF, dataStart, lastRow);
                    break;
                case SheetMode.Tunggakan1:
                    hariPokok = BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    nomPokok  = BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    break;
                case SheetMode.Tunggakan2:
                    hariPokok = BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    nomPokok  = BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    hariBasil = BacaKolomDouble(ws, spec.ColHariBasil, dataStart, lastRow);
                    nomBasil  = BacaKolomDouble(ws, spec.ColNomBasil,  dataStart, lastRow);
                    break;
            }

            for (int i = 0; i < n; i++)
            {
                string cif       = (cifArr[i] ?? "").Trim();
                string noKontrak = (noRekArr[i] ?? "").Trim();
                if (string.IsNullOrEmpty(cif) || string.IsNullOrEmpty(noKontrak)) continue;

                string nama     = (namaArr[i] ?? "").Trim();
                double jaminan  = jamArr.Length  > i ? jamArr[i]  : 0;
                int    kualitas = kualArr.Length > i ? (int)kualArr[i] : 0;

                double osVal;
                bool   masuk;

                switch (spec.Mode)
                {
                    case SheetMode.AF:
                        masuk = true;
                        osVal = osArr[i];
                        break;

                    case SheetMode.Tunggakan1:
                        // KC1100 — LOGIKA BARU:
                        // Filter: baris masuk jika nama debitur (kolom D) tidak kosong.
                        // Baris tanpa nama = baris agunan/subtotal → di-skip.
                        // OS = seluruh nomPokok (AL), tidak peduli hari tunggakan.
                        // Alasan: ada debitur dengan hari tunggakan = 0 tapi nominal > 0.
                        masuk = !string.IsNullOrEmpty(nama);
                        osVal = nomPokok[i];
                        break;

                    case SheetMode.Tunggakan2:
                        // KC1000 — LOGIKA BARU:
                        // Filter: baris masuk jika nama debitur (kolom D) tidak kosong.
                        // Baris tanpa nama = baris agunan/subtotal → di-skip.
                        // OS = nomPokok (AL) + nomBasil (AN), tidak peduli hari tunggakan.
                        // Alasan: ada debitur dengan hari tunggakan = 0 tapi nominal > 0.
                        masuk = !string.IsNullOrEmpty(nama);
                        osVal = nomPokok[i] + nomBasil[i];
                        break;

                    default:
                        continue;
                }

                if (!masuk) continue;

                if (!dictCIF.ContainsKey(cif))
                    dictCIF[cif] = new CIFData(cif, nama, spec.SheetName);
                else if (string.IsNullOrEmpty(dictCIF[cif].Nama) && !string.IsNullOrEmpty(nama))
                    dictCIF[cif].Nama = nama;

                dictCIF[cif].TambahKontrak(noKontrak, osVal, jaminan, kualitas, spec.SheetName);
            }

            return dictCIF;
        }

        // ================================================================
        // Private helpers
        // ================================================================

        /// <summary>
        /// Konversi nilai cell ke string dengan aman.
        /// Menangani: null, string, double, int, bool, error code.
        /// </summary>
        private static string ValToString(object val)
        {
            if (val == null)                return "";
            if (val is System.DBNull)       return "";
            // Formula error (mis. #VALUE!, #REF!) → int/uint di interop
            if (val is int || val is uint)  return "";
            return val.ToString().Trim();
        }

        /// <summary>
        /// Konversi nilai cell ke double dengan aman.
        /// Menangani: null, string numerik, double, int, bool, error code.
        /// </summary>
        private static double ValToDouble(object val)
        {
            if (val == null)                return 0;
            if (val is System.DBNull)       return 0;
            if (val is int || val is uint)  return 0;   // formula error
            if (val is double  d)           return d;
            if (val is float   f)           return f;
            if (val is int     i)           return i;
            if (val is long    l)           return l;
            if (val is bool    b)           return b ? 1 : 0;
            double parsed;
            string s = val.ToString().Trim().Replace(",", ".");
            return double.TryParse(s,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out parsed) ? parsed : 0;
        }
    }
}