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
        // Tahap 7e: jenis pengikatan 90 = tidak ada pengikatan.
        // Menerima "90", "090", 90, "90 - Tidak diikat", dsb. (blok digit pertama).
        // Kosong / kode lain → false (nilai agunan tetap dipakai).
        // ----------------------------------------------------------------
        public static bool TanpaIkatan(string kode)
        {
            if (string.IsNullOrEmpty(kode)) return false;
            string digit = "";
            foreach (char c in kode.Trim())
            {
                if (char.IsDigit(c)) digit += c;
                else if (digit.Length > 0) break;
            }
            int v;
            return digit.Length > 0 && int.TryParse(digit, out v) && v == 90;
        }

        // ----------------------------------------------------------------
        // Baca satu sheet KC → isi dictCIF
        //
        // Satu nomor kontrak dapat tercatat di beberapa baris: setiap baris
        // mewakili satu agunan. Nilai agunan kontrak = jumlah seluruh baris
        // dengan nomor kontrak (kolom J) yang sama.
        //
        // Nilai agunan per baris (Tahap 7e): kolom Nilai Agunan/Jaminan, KECUALI
        // jenis pengikatan = 90 (tidak ada pengikatan) → 0.
        //
        //   KC0600–KC0900 (mode AF): setiap baris terisi lengkap → semua baris
        //                            masuk, jaminan diakumulasi di TambahKontrak.
        //   KC1000 / KC1100        : hanya baris pertama yang bernama (kolom D);
        //                            baris agunan berikutnya tanpa nama. Baris
        //                            tanpa nama TIDAK menambah OS/kontrak baru,
        //                            tetapi nilai agunannya ditambahkan ke kontrak
        //                            yang nomornya sama (Tahap 7d).
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
            // Baris agunan bisa tanpa CIF → cari baris terakhir dari kolom CIF
            // dan kolom nomor kontrak, ambil yang terbesar.
            int lastRow   = Math.Max(CariLastRow(ws, spec.ColCIF,   headerRow),
                                     CariLastRow(ws, spec.ColNoRek, headerRow));
            if (lastRow <= headerRow) return dictCIF;

            int dataStart = headerRow + 1;
            int n = lastRow - dataStart + 1;

            string[] cifArr   = BacaKolomString(ws, spec.ColCIF,      dataStart, lastRow);
            string[] namaArr  = BacaKolomString(ws, spec.ColNama,     dataStart, lastRow);
            string[] noRekArr = BacaKolomString(ws, spec.ColNoRek,    dataStart, lastRow);
            double[] jamArr   = BacaKolomDouble(ws, spec.ColJaminan,  dataStart, lastRow);
            // Tahap 7e: jenis pengikatan — kode 90 (tidak diikat) → nilai agunan = 0
            string[] ikatArr  = string.IsNullOrEmpty(spec.ColIkatan)
                                ? new string[n]
                                : BacaKolomString(ws, spec.ColIkatan, dataStart, lastRow);
            double[] kualArr  = BacaKolomDouble(ws, spec.ColKualitas, dataStart, lastRow);

            double[] hariPokok = null, nomPokok = null, nomUjroh = null, hariBasil = null, nomBasil = null;
            double[] osArr = null, hariAF = null;

            switch (spec.Mode)
            {
                case SheetMode.AF:
                    osArr  = BacaKolomDouble(ws, spec.ColOS,     dataStart, lastRow);
                    hariAF = BacaKolomDouble(ws, spec.ColHariAF, dataStart, lastRow);
                    break;
                case SheetMode.Tunggakan1:
                    // KC1100 (ijarah): pokok (AL) + ujroh/imbalan (AM).
                    hariPokok = BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    nomPokok  = BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    nomUjroh  = BacaKolomDouble(ws, spec.ColNomUjroh,  dataStart, lastRow);
                    // Ruang lingkup KC1100: imbalan hanya kualitas 1–2; pokok semua kualitas
                    RuangLingkup.TerapkanKC1100(spec, ws, dataStart, lastRow, nomUjroh);
                    break;
                case SheetMode.Tunggakan2:
                    hariPokok = BacaKolomDouble(ws, spec.ColHariPokok, dataStart, lastRow);
                    nomPokok  = BacaKolomDouble(ws, spec.ColNomPokok,  dataStart, lastRow);
                    hariBasil = BacaKolomDouble(ws, spec.ColHariBasil, dataStart, lastRow);
                    nomBasil  = BacaKolomDouble(ws, spec.ColNomBasil,  dataStart, lastRow);
                    // Ruang lingkup KC1000: pokok hanya kualitas Macet; bagi hasil sesuai ⚙ Pengaturan
                    RuangLingkup.TerapkanKC1000(spec, ws, dataStart, lastRow, nomPokok, nomBasil);
                    break;
            }

            // Tahap 7d — kontrak yang sudah terbaca di sheet ini (kunci = no. kontrak),
            // dan agunan dari baris tanpa nama yang muncul SEBELUM baris utamanya.
            var kontrakSheet  = new Dictionary<string, KontrakData>(StringComparer.OrdinalIgnoreCase);
            var agunanTertunda = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < n; i++)
            {
                string cif       = (cifArr[i] ?? "").Trim();
                string noKontrak = (noRekArr[i] ?? "").Trim();
                if (string.IsNullOrEmpty(noKontrak)) continue;

                string nama     = (namaArr[i] ?? "").Trim();
                double jaminan  = jamArr.Length  > i ? jamArr[i]  : 0;
                if (ikatArr.Length > i && TanpaIkatan(ikatArr[i])) jaminan = 0;
                int    kualitas = kualArr.Length > i ? (int)kualArr[i] : 0;

                // ---- Baris agunan tambahan ----
                // KC1000 / KC1100: baris tanpa nama; semua sheet: baris tanpa CIF.
                // Nilai agunannya ditambahkan ke kontrak bernomor sama.
                bool barisAgunan = string.IsNullOrEmpty(cif) ||
                                   (spec.Mode != SheetMode.AF && string.IsNullOrEmpty(nama));
                if (barisAgunan)
                {
                    if (jaminan == 0) continue;
                    KontrakData kAda;
                    if (kontrakSheet.TryGetValue(noKontrak, out kAda))
                        kAda.Jaminan += jaminan;
                    else
                    {
                        double t;
                        agunanTertunda.TryGetValue(noKontrak, out t);
                        agunanTertunda[noKontrak] = t + jaminan;
                    }
                    continue;
                }

                double osVal;
                double hariVal;      // hari tunggakan untuk kriteria penurunan nilai 7–30 hari

                switch (spec.Mode)
                {
                    case SheetMode.AF:
                        osVal   = osArr[i];
                        hariVal = hariAF[i];
                        break;

                    case SheetMode.Tunggakan1:
                        // KC1100: OS/EAD = tunggakan pokok (AL) + tunggakan ujroh (AM); aging gabungan di AK.
                        osVal   = nomPokok[i] + nomUjroh[i];
                        hariVal = hariPokok[i];
                        break;

                    case SheetMode.Tunggakan2:
                        // KC1000: OS = tunggakan pokok (AL) + tunggakan basil (AN); aging terlama.
                        osVal   = nomPokok[i] + nomBasil[i];
                        hariVal = Math.Max(hariPokok[i], hariBasil[i]);
                        break;

                    default:
                        continue;
                }

                if (!dictCIF.ContainsKey(cif))
                    dictCIF[cif] = new CIFData(cif, nama, spec.SheetName);
                else if (string.IsNullOrEmpty(dictCIF[cif].Nama) && !string.IsNullOrEmpty(nama))
                    dictCIF[cif].Nama = nama;

                dictCIF[cif].TambahKontrak(noKontrak, osVal, jaminan, kualitas, spec.SheetName, hariVal);

                KontrakData k = dictCIF[cif].Kontrak[noKontrak];
                if (!kontrakSheet.ContainsKey(noKontrak))
                {
                    kontrakSheet[noKontrak] = k;
                    double tunda;
                    if (agunanTertunda.TryGetValue(noKontrak, out tunda))
                    {
                        k.Jaminan += tunda;
                        agunanTertunda.Remove(noKontrak);
                    }
                }
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