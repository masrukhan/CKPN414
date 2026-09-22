using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA modul "module lgd chortfall pby macet".
    ///
    /// Logika inti:
    ///   1. Baca file referensi periodik (Juni/Desember tiap tahun)
    ///   2. Filter rekening kualitas Macet (kol 5) + jaminan fisik
    ///   3. Rekening "selesai" = terakhir muncul macet di periode X,
    ///      lalu di periode X+1 benar-benar hilang (bukan turun kol)
    ///   4. Prefill kolom G = nilai agunan net haircut (atau OS jika tidak ada agunan)
    ///   5. Tulis ke sheet "B4.LGD-CS MACET" + ringkasan LGD Weighted
    ///
    /// Parameter dikirim dari VBA:
    ///   fileConfigStr : "labelPer1=tahun1=bulan1=path1|labelPer2=..."
    ///   haircutAgn    : haircut biaya penjualan agunan (dari Master!C70)
    ///   sheetKCList   : "KC0600,KC0700,..."
    /// </summary>
    internal class LGDCollateralShortfall
    {
        private readonly Excel.Application _app;

        private const string SheetOutput = "B4.LGD-CS MACET";
        private const string SheetLog    = "Audit Log";
        private const int    ROW_FIRST   = 5;

        // Warna tema (RGB format untuk Excel .Color property)
        private const int CLR_GREEN_DARK   = 0x436E1F;  // sesuai VBA: RGB(31,110,67) -> &H436E1F
        private const int CLR_GREEN_MED    = 0x2E8B2E;  // RGB(46,139,87)
        private const int CLR_GRAY_LIGHT   = 0x969696;  // RGB(150,150,150)
        private const int CLR_BLUE_DARK    = 0x990033;  // RGB(0,51,153) BGR
        private const int CLR_RED_DARK     = 0x0000C0;  // RGB(192,0,0) BGR
        private const int CLR_WHITE        = 0xFFFFFF;
        private const int CLR_YELLOW_SOFT  = 0xE7FDFF;  // RGB(255,253,231) BGR

        private const string FMT_NUM = "_(* #,##0_);_(* (#,##0);_(* \"-\"??_);_(@_)";

        // ----------------------------------------------------------------
        // Spec kolom per jenis KC untuk modul LGD-CS
        // Berbeda dengan SheetSpec di Models/ karena kolom berbeda
        // ----------------------------------------------------------------
        private class KCSpec
        {
            public string ColNama   = "D";
            public string ColNoRek  = "J";
            public string ColKual;
            public string ColBaki1;
            public string ColBaki2;   // kosong jika tidak ada
            public string ColAgn;
            public string ColJenis;   // jenis jaminan
        }

        private static KCSpec BuatSpec(string kcName)
        {
            switch (kcName.ToUpper())
            {
                case "KC0600": case "KC0700": case "KC0800":
                    return new KCSpec { ColKual="Y", ColBaki1="AC", ColBaki2="", ColAgn="AS", ColJenis="AJ" };
                case "KC0900":
                    return new KCSpec { ColKual="Y", ColBaki1="AC", ColBaki2="", ColAgn="AQ", ColJenis="AH" };
                case "KC1000":
                    return new KCSpec { ColKual="AE", ColBaki1="AG", ColBaki2="", ColAgn="AY", ColJenis="AP" };
                case "KC1100":
                    return new KCSpec { ColKual="AF", ColBaki1="AB", ColBaki2="AI", ColAgn="AS", ColJenis="AO" };
                default:
                    return null;
            }
        }

        // ----------------------------------------------------------------
        // Model data rekening macet per periode
        // ----------------------------------------------------------------
        private class RekMacet
        {
            public double Baki;
            public double Agunan;
            public string Nama;
        }

        // Model data lintas periode (kemunculan paling awal)
        private class RekAll
        {
            public int    ThnPertamaMacet;
            public double NilaiAgunanAwal;
            public double BakiDebetAwal;
            public string Nama;
        }

        // Model konfigurasi file per periode
        private class PeriodeCfg
        {
            public string Label;
            public int    RefYear;
            public int    RefMonth;   // 6=Juni, 12=Desember
            public string RefKey;     // "yyyyMM" mis. "202406"
            public string Path;
        }

        public LGDCollateralShortfall(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ================================================================
        // Entry point
        // fileConfigStr: "label=tahun=bulan=path|label=tahun=bulan=path|..."
        // haircutAgn   : nilai 0.0-1.0, mis. 0.2 untuk 20%
        // sheetKCList  : "KC0600,KC0700,..."
        // ================================================================
        public void Hitung(string fileConfigStr, double haircutAgn, string sheetKCList)
        {
            var wb  = _app.ActiveWorkbook;
            var ws  = CariSheet(wb, SheetOutput);
            var wsLog = CariSheet(wb, SheetLog);

            if (ws    == null) throw new InvalidOperationException("Sheet '" + SheetOutput + "' tidak ditemukan.");
            if (wsLog == null) throw new InvalidOperationException("Sheet '" + SheetLog    + "' tidak ditemukan.");

            // Parse konfigurasi file periode
            var periodes = ParsePeriodeCfg(fileConfigStr);
            if (periodes.Count == 0)
                throw new ArgumentException("fileConfigStr tidak valid: " + fileConfigStr);

            // Validasi file
            foreach (var p in periodes)
                if (!string.IsNullOrEmpty(p.Path) && !System.IO.File.Exists(p.Path))
                    throw new System.IO.FileNotFoundException(
                        "File periode " + p.Label + " tidak ditemukan:\n" + p.Path);

            string[] sheets = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
            string   lastPerKey = periodes[periodes.Count - 1].RefKey;

            // ---- Bersihkan output ----
            BersihkanSheet(ws);

            // ---- Baca semua file periode ----
            // dictByPer[refKey]    = Dictionary<noRek, RekMacet>   (hanya macet+jaminanFisik)
            // dictAllRekByPer[refKey] = HashSet<noRek>             (semua rekening, apapun kualitas)
            // dictAll[noRek]       = RekAll                        (kemunculan paling awal)
            var dictByPer       = new Dictionary<string, Dictionary<string, RekMacet>>(StringComparer.OrdinalIgnoreCase);
            var dictAllRekByPer = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var dictAll         = new Dictionary<string, RekAll>(StringComparer.OrdinalIgnoreCase);

            int logRow     = NextAuditLogRow(wsLog);
            int nOK = 0, nSkip = 0;
            int totRekSemua = 0, totRekMacet = 0;
            double totBakiMacet = 0;

            foreach (var periode in periodes)
            {
                var dictY       = new Dictionary<string, RekMacet>(StringComparer.OrdinalIgnoreCase);
                var dictAllRekY = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (string.IsNullOrEmpty(periode.Path) || !System.IO.File.Exists(periode.Path))
                {
                    dictByPer[periode.RefKey]       = dictY;
                    dictAllRekByPer[periode.RefKey] = dictAllRekY;
                    TulisLog(wsLog, logRow++, periode.Label, periode.RefYear, 0,
                             periode.Path, "Skipped - Path tidak ditemukan", 0, 0, 0);
                    nSkip++;
                    continue;
                }

                int nRekTotal = 0, nRekMacetFile = 0;
                double sumBakiFile = 0;

                Excel.Workbook wbSrc = null;
                try
                {
                    wbSrc = _app.Workbooks.Open(periode.Path, UpdateLinks: 0, ReadOnly: true);

                    foreach (var shName in sheets)
                        BacaSheetKC(wbSrc, shName, dictY, dictAllRekY, dictAll,
                                    periode.RefYear,
                                    ref nRekTotal, ref nRekMacetFile, ref sumBakiFile);

                    wbSrc.Close(false); wbSrc = null;
                }
                finally { if (wbSrc != null) try { wbSrc.Close(false); } catch { } }

                dictByPer[periode.RefKey]       = dictY;
                dictAllRekByPer[periode.RefKey] = dictAllRekY;

                totRekSemua  += nRekTotal;
                totRekMacet  += nRekMacetFile;
                totBakiMacet += sumBakiFile;
                nOK++;

                TulisLog(wsLog, logRow++, periode.Label, periode.RefYear, 0,
                         periode.Path, "OK", nRekTotal, nRekMacetFile, sumBakiFile);
            }

            // ---- Sort rekening: asc tahun macet, lalu noRek ----
            var sortedKeys = new List<string>(dictAll.Keys);
            sortedKeys.Sort((a, b) =>
            {
                int ya = dictAll[a].ThnPertamaMacet;
                int yb = dictAll[b].ThnPertamaMacet;
                if (ya != yb) return ya.CompareTo(yb);
                return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
            });

            // ---- Header kolom I ----
            TulisHeaderKolomI(ws);

            // ---- Tulis baris detail ----
            int outRow = ROW_FIRST;

            foreach (string noRek in sortedKeys)
            {
                RekAll info = dictAll[noRek];

                // Cari periode terakhir debitur masih MACET
                int lastIdxFound = -1;
                for (int yi = periodes.Count - 1; yi >= 0; yi--)
                {
                    if (dictByPer.ContainsKey(periodes[yi].RefKey) &&
                        dictByPer[periodes[yi].RefKey].ContainsKey(noRek))
                    {
                        lastIdxFound = yi; break;
                    }
                }
                if (lastIdxFound < 0) continue;

                // Masih macet di periode terakhir → belum selesai, SKIP
                if (periodes[lastIdxFound].RefKey == lastPerKey) continue;

                // Cek periode sesudah → apakah masih ada di file (turun kol)?
                int periodeSesudahIdx = lastIdxFound + 1;
                if (dictAllRekByPer.ContainsKey(periodes[periodeSesudahIdx].RefKey) &&
                    dictAllRekByPer[periodes[periodeSesudahIdx].RefKey].Contains(noRek))
                    continue;   // cuma turun kolektibilitas, bukan eksekusi

                // Rekening selesai/eksekusi
                int    thnEksekusi = periodes[lastIdxFound].RefYear;
                double bakiLast    = dictByPer[periodes[lastIdxFound].RefKey][noRek].Baki;
                double agnLast     = dictByPer[periodes[lastIdxFound].RefKey][noRek].Agunan;

                string sumberPrefill = agnLast > 0 ? "AGN" : "OS";

                // Tulis baris
                TulisBaris(ws, outRow, noRek,
                           info.BakiDebetAwal, info.NilaiAgunanAwal,
                           info.ThnPertamaMacet, thnEksekusi,
                           bakiLast, agnLast, haircutAgn, sumberPrefill,
                           info.Nama);
                outRow++;
            }

            int lastOut = outRow - 1;

            // ---- Catatan workflow di B3 ----
            TulisCatatan(ws, haircutAgn);

            // ---- Baris total ----
            int totalRow = lastOut + 2;
            TulisBarisTotalDanBorder(ws, totalRow, lastOut);

            // ---- Blok ringkasan LGD Weighted ----
            int sumStart = totalRow + 3;
            TulisRingkasanLGD(ws, sumStart, totalRow);

            // ---- Log ringkasan ----
            double totRecovery = 0;
            for (int r = ROW_FIRST; r <= lastOut; r++)
            {
                object v = ((Excel.Range)ws.Cells[r, "G"]).Value2;
                if (v != null) { double d; if (double.TryParse(v.ToString(), out d)) totRecovery += d; }
            }
            TulisLog(wsLog, logRow, "RINGKASAN",
                     periodes[0].RefYear, periodes[periodes.Count - 1].RefYear,
                     "Files: " + nOK + " OK, " + nSkip + " skipped",
                     "Ringkasan", totRekMacet, lastOut - ROW_FIRST + 1, totBakiMacet);
            // Kolom M: total recovery, N: scope KC
            ((Excel.Range)wsLog.Cells[logRow, "M"]).Value2 = totRecovery;
            ((Excel.Range)wsLog.Cells[logRow, "M"]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[logRow, "N"]).Value2 = "KC: " + sheetKCList;
            Excel.Range hlRow = (Excel.Range)wsLog.Range["A" + logRow, "T" + logRow];
            hlRow.Interior.Color = CLR_YELLOW_SOFT;
            hlRow.Font.Bold = true;

            System.Windows.Forms.MessageBox.Show(
                "Perhitungan LGD Collateral Shortfall (MACET) selesai.\n" +
                "Ruang lingkup KC       : " + sheetKCList + "\n" +
                "File referensi terbaca : " + periodes.Count + "\n" +
                "Rekening macet selesai : " + (lastOut - ROW_FIRST + 1),
                "LGD Collateral Shortfall",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        // ================================================================
        // BacaSheetKC: baca satu sheet KC dari workbook sumber
        // Setara BacaSheetKC VBA — filter macet + jaminan fisik
        // ================================================================
        private void BacaSheetKC(
            Excel.Workbook wbSrc, string kcName,
            Dictionary<string, RekMacet> dictY,
            HashSet<string> dictAllRekY,
            Dictionary<string, RekAll> dictAll,
            int refYear,
            ref int nRekTotal, ref int nRekMacet, ref double sumBaki)
        {
            if (!ExcelHelper.SheetAda(wbSrc, kcName)) return;
            KCSpec spec = BuatSpec(kcName);
            if (spec == null) return;

            var ws      = (Excel.Worksheet)wbSrc.Worksheets[kcName];
            int lastRow = ExcelHelper.CariLastRow(ws, spec.ColNoRek, 4);
            if (lastRow < 5) return;

            // Bulk read
            string[] namaArr  = ExcelHelper.BacaKolomString(ws, spec.ColNama,  5, lastRow);
            string[] rekArr   = ExcelHelper.BacaKolomString(ws, spec.ColNoRek, 5, lastRow);
            string[] kualArr  = ExcelHelper.BacaKolomString(ws, spec.ColKual,  5, lastRow);
            double[] baki1Arr = ExcelHelper.BacaKolomDouble(ws, spec.ColBaki1, 5, lastRow);
            double[] baki2Arr = !string.IsNullOrEmpty(spec.ColBaki2)
                                ? ExcelHelper.BacaKolomDouble(ws, spec.ColBaki2, 5, lastRow)
                                : new double[namaArr.Length];
            double[] agnArr   = ExcelHelper.BacaKolomDouble(ws, spec.ColAgn,   5, lastRow);
            string[] jenisArr = ExcelHelper.BacaKolomString(ws, spec.ColJenis, 5, lastRow);

            for (int i = 0; i < rekArr.Length; i++)
            {
                string noRek = rekArr[i].Trim();
                if (string.IsNullOrEmpty(noRek)) continue;

                nRekTotal++;
                dictAllRekY.Add(noRek);   // HashSet auto-deduplicate

                // Filter kualitas macet
                if (!IsKualitasMacet(kualArr[i])) continue;

                // Filter jaminan fisik
                if (!IsJaminanFisik(jenisArr[i])) continue;

                // Hitung baki debet
                double baki = !string.IsNullOrEmpty(spec.ColBaki2)
                    ? baki1Arr[i] - baki2Arr[i]   // KC1100: AB - AI
                    : baki1Arr[i];

                double agn  = agnArr[i];
                string nama = namaArr[i].Trim();

                nRekMacet++;
                sumBaki += baki;

                // dictY per periode: simpan kemunculan pertama
                if (!dictY.ContainsKey(noRek))
                    dictY[noRek] = new RekMacet { Baki = baki, Agunan = agn, Nama = nama };

                // dictAll lintas periode: simpan kemunculan paling awal
                if (!dictAll.ContainsKey(noRek))
                    dictAll[noRek] = new RekAll
                    {
                        ThnPertamaMacet = refYear,
                        NilaiAgunanAwal = agn,
                        BakiDebetAwal   = baki,
                        Nama            = nama
                    };
                else if (refYear < dictAll[noRek].ThnPertamaMacet)
                    dictAll[noRek] = new RekAll
                    {
                        ThnPertamaMacet = refYear,
                        NilaiAgunanAwal = agn,
                        BakiDebetAwal   = baki,
                        Nama            = nama
                    };
            }
        }

        // ================================================================
        // Tulis satu baris rekening ke output sheet
        // ================================================================
        private void TulisBaris(
            Excel.Worksheet ws, int row, string noRek,
            double pokokAwal, double nilaiAwal,
            int thnDiserahkan, int thnEksekusi,
            double bakiLast, double agnLast,
            double haircutAgn, string sumberPrefill,
            string nama)
        {
            ((Excel.Range)ws.Cells[row, "B"]).NumberFormat = "@";
            ((Excel.Range)ws.Cells[row, "B"]).Value2 = noRek;
            ((Excel.Range)ws.Cells[row, "C"]).Value2 = pokokAwal;
            ((Excel.Range)ws.Cells[row, "D"]).Value2 = nilaiAwal;

            ((Excel.Range)ws.Cells[row, "E"]).NumberFormat = "@";
            if (thnDiserahkan > 0)
                ((Excel.Range)ws.Cells[row, "E"]).Value2 = thnDiserahkan.ToString();

            ((Excel.Range)ws.Cells[row, "F"]).NumberFormat = "@";
            if (thnEksekusi > 0)
                ((Excel.Range)ws.Cells[row, "F"]).Value2 = thnEksekusi.ToString();

            // Kolom G: prefill formula nilai eksekusi
            Excel.Range cellG = (Excel.Range)ws.Cells[row, "G"];
            if (thnEksekusi > 0)
            {
                if (agnLast > 0)
                    cellG.Formula = "=MIN(D" + row + "*(1-Master!$C$70),C" + row + ")";
                else
                    cellG.Formula = "=C" + row;

                // Warna font sesuai sumber prefill
                cellG.Font.Color = sumberPrefill == "AGN" ? CLR_BLUE_DARK : CLR_RED_DARK;

                // Comment penjelasan prefill
                try { cellG.Comment.Delete(); } catch { }
                string pctStr = (haircutAgn * 100).ToString("0.0") + "%";
                string cmtText = sumberPrefill == "AGN"
                    ? "Prefill otomatis dari sistem.\n" +
                      "Sumber: Nilai Agunan x (1 - " + pctStr + ") biaya penjualan, cap di OS\n" +
                      "WAJIB diganti dengan REALISASI hasil eksekusi\n" +
                      "(net biaya notaris/lelang/pajak) dari Berita Acara Lelang."
                    : "Prefill otomatis dari sistem.\n" +
                      "Sumber: OS terakhir (baki debet) — tidak ada data agunan\n" +
                      "WAJIB diganti dengan REALISASI hasil eksekusi.";

                Excel.Comment cmt = cellG.AddComment(cmtText);
                cmt.Shape.TextFrame.AutoSize = true;
            }

            // Kolom H: shortfall = MAX(0, OS - Recovery)
            ((Excel.Range)ws.Cells[row, "H"]).Formula =
                "=IF(G" + row + ">C" + row + ",0,C" + row + "-G" + row + ")";

            // Kolom I: nama debitur (informatif)
            Excel.Range cellI = (Excel.Range)ws.Cells[row, "I"];
            cellI.Value2 = nama;
            cellI.Font.Italic = true;
            cellI.Font.Color  = CLR_GRAY_LIGHT;
        }

        // ================================================================
        // Tulis catatan di B3
        // ================================================================
        private void TulisCatatan(Excel.Worksheet ws, double haircutAgn)
        {
            try { ((Excel.Range)ws.Range["B3", "I3"]).UnMerge(); } catch { }

            string pctStr = (haircutAgn * 100).ToString("0.0") + "%";
            ((Excel.Range)ws.Range["B3"]).Value2 =
                "Catatan: Kolom G (nilai eksekusi) di-prefill otomatis. " +
                "WAJIB di-override dengan realisasi penjualan agunan (net biaya). " +
                "Verifikasi: jika debitur lunas BUKAN karena eksekusi agunan, hapus baris secara manual. " +
                "Warna prefill: MERAH = dari OS terakhir, BIRU = dari nilai agunan (haircut " + pctStr + "). " +
                "Kolom I (Nama Debitur) bersifat informatif — dapat dihapus setelah verifikasi.";

            ((Excel.Range)ws.Range["B3", "I3"]).Merge();
            Excel.Range r3 = (Excel.Range)ws.Range["B3"];
            r3.Font.Bold  = true; r3.Font.Italic = true;
            r3.Font.Color = CLR_WHITE;
            r3.Interior.Color = CLR_GREEN_MED;
            r3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            r3.VerticalAlignment   = Excel.XlVAlign.xlVAlignCenter;
            r3.WrapText = true;

            ApplyBorder((Excel.Range)ws.Range["B3", "I3"],
                        Excel.XlLineStyle.xlContinuous, CLR_GREEN_DARK);
        }

        // ================================================================
        // Tulis header kolom I baris 4
        // ================================================================
        private void TulisHeaderKolomI(Excel.Worksheet ws)
        {
            Excel.Range cell = (Excel.Range)ws.Cells[4, "I"];
            cell.Value2 = "Nama Debitur (informatif)";
            cell.Font.Bold   = true; cell.Font.Italic = true;
            cell.Font.Color  = CLR_WHITE;
            cell.Interior.Color = CLR_GREEN_DARK;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        }

        // ================================================================
        // Tulis baris total dan border
        // ================================================================
        private void TulisBarisTotalDanBorder(Excel.Worksheet ws, int totalRow, int lastOut)
        {
            ((Excel.Range)ws.Cells[totalRow, "B"]).Value2 = "Total";
            ((Excel.Range)ws.Cells[totalRow, "C"]).Formula =
                "=IFERROR(SUM(C" + ROW_FIRST + ":C" + lastOut + "),0)";
            ((Excel.Range)ws.Cells[totalRow, "G"]).Formula =
                "=IFERROR(SUM(G" + ROW_FIRST + ":G" + lastOut + "),0)";
            ((Excel.Range)ws.Cells[totalRow, "H"]).Formula =
                "=IFERROR(SUM(H" + ROW_FIRST + ":H" + lastOut + "),0)";

            // Format angka data
            ((Excel.Range)ws.Range["C" + ROW_FIRST, "D" + lastOut]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Range["G" + ROW_FIRST, "H" + lastOut]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Range["C" + totalRow, "H" + totalRow]).NumberFormat = FMT_NUM;

            // Border data
            ApplyGreenBorderMacet((Excel.Range)ws.Range["B4", "H" + lastOut]);
            ApplyGreenBorderMacet((Excel.Range)ws.Range["B" + totalRow, "H" + totalRow]);

            // Border abu-abu kolom I (informatif)
            ApplyBorderDash((Excel.Range)ws.Range["I4", "I" + lastOut]);

            // Warna baris total
            Excel.Range totalRng = (Excel.Range)ws.Range["B" + totalRow, "H" + totalRow];
            totalRng.Interior.Color = CLR_GREEN_MED;
            totalRng.Font.Color     = CLR_WHITE;
            totalRng.Font.Bold      = true;
        }

        // ================================================================
        // Tulis blok ringkasan LGD Weighted
        // ================================================================
        private void TulisRingkasanLGD(Excel.Worksheet ws, int sumStart, int totalRow)
        {
            ((Excel.Range)ws.Cells[sumStart,     "B"]).Value2 = "Total WO";
            ((Excel.Range)ws.Cells[sumStart + 1, "B"]).Value2 = "Total Recovery";
            ((Excel.Range)ws.Cells[sumStart + 2, "B"]).Value2 = "LGD Weighted";

            // Kolom C: Collateral Shortfall
            ((Excel.Range)ws.Cells[sumStart - 1, "C"]).Value2 = "Collateral Shortfall";
            ((Excel.Range)ws.Cells[sumStart,     "C"]).Formula =
                "=INDEX($C:$C,MATCH(\"Total\",$B:$B,0))";   // C9 — Total WO
            ((Excel.Range)ws.Cells[sumStart + 1, "C"]).Formula =
                "=INDEX($G:$G,MATCH(\"Total\",$B:$B,0))";   // C10 — Total Recovery
            ((Excel.Range)ws.Cells[sumStart + 2, "C"]).Formula =
                "=IFERROR(1-C" + (sumStart + 1) + "/C" + sumStart + ",0)";

            // Kolom D: Expected Recoveries (link ke sheet B3.LGD-ER)
            ((Excel.Range)ws.Cells[sumStart - 1, "D"]).Value2 = "Expected Recoveries";
            ((Excel.Range)ws.Cells[sumStart,     "D"]).Formula =
                "=INDEX('B3.LGD-ER'!$D:$D,MATCH(\"Total WO\",'B3.LGD-ER'!$C:$C,0))";
            ((Excel.Range)ws.Cells[sumStart + 1, "D"]).Formula =
                "=INDEX('B3.LGD-ER'!$D:$D,MATCH(\"Total Recovery\",'B3.LGD-ER'!$C:$C,0))";
            ((Excel.Range)ws.Cells[sumStart + 2, "D"]).Formula =
                "=IFERROR(1-D" + (sumStart + 1) + "/D" + sumStart + ",0)";

            // Kolom E: Combined
            ((Excel.Range)ws.Cells[sumStart - 1, "E"]).Value2 = "Combined";
            ((Excel.Range)ws.Cells[sumStart,     "E"]).Formula =
                "=C" + sumStart + "+D" + sumStart;
            ((Excel.Range)ws.Cells[sumStart + 1, "E"]).Formula =
                "=C" + (sumStart + 1) + "+D" + (sumStart + 1);
            ((Excel.Range)ws.Cells[sumStart + 2, "E"]).Formula =
                "=IFERROR(1-E" + (sumStart + 1) + "/E" + sumStart + ",0)";

            // Format
            ((Excel.Range)ws.Range["C" + sumStart, "E" + (sumStart + 1)]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Range["C" + (sumStart + 2), "E" + (sumStart + 2)]).NumberFormat = "0.00%";

            // Header row warna hijau
            Excel.Range hdrRng = (Excel.Range)ws.Range["B" + (sumStart - 1), "E" + (sumStart - 1)];
            hdrRng.Font.Bold = true; hdrRng.Font.Color = CLR_WHITE;
            hdrRng.Interior.Color = CLR_GREEN_DARK;

            // Baris LGD Weighted: warna kuning
            Excel.Range lgdRng = (Excel.Range)ws.Range["B" + (sumStart + 2), "E" + (sumStart + 2)];
            lgdRng.Interior.Color = CLR_YELLOW_SOFT;
            lgdRng.Font.Bold = true;

            ApplyGreenBorderMacet((Excel.Range)ws.Range["B" + (sumStart - 1), "E" + (sumStart + 2)]);
        }

        // ================================================================
        // Bersihkan sheet output
        // ================================================================
        private void BersihkanSheet(Excel.Worksheet ws)
        {
            // ---- Tentukan batas bawah dari beberapa sumber ----

            // 1. Dari End(xlUp) tiap kolom data
            int lastB = CariLastRow(ws, "B", ROW_FIRST);
            int lastC = CariLastRow(ws, "C", ROW_FIRST);
            int lastH = CariLastRow(ws, "H", ROW_FIRST);
            int lastI = CariLastRow(ws, "I", ROW_FIRST);
            int last  = Math.Max(Math.Max(lastB, lastC), Math.Max(lastH, lastI));

            // 2. Iterasi Comments — comment "yatim" tidak terdeteksi End(xlUp)
            //    karena cell-nya kosong (value sudah dihapus tapi comment tertinggal).
            //    Setara loop "For Each cmt In ws.Comments" di VBA.
            try
            {
                Excel.Comments comments = ws.Comments;
                foreach (Excel.Comment cmt in comments)
                {
                    // .Parent mengembalikan object — cast eksplisit wajib
                    Excel.Range parent = (Excel.Range)cmt.Parent;
                    // Hanya comment di kolom B:I (kolom 2..9)
                    if (parent.Column >= 2 && parent.Column <= 9)
                        if (parent.Row > last) last = parent.Row;
                }
            }
            catch { /* abaikan jika Comments tidak bisa diakses */ }

            // 3. UsedRange sebagai jaring pengaman terakhir
            //    Menangkap format/border yang tersisa walau value sudah kosong.
            try
            {
                int urLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
                if (urLast > last) last = urLast;
            }
            catch { }

            if (last < ROW_FIRST) return;

            Excel.Range rng = (Excel.Range)ws.Range["B" + ROW_FIRST, "I" + last];

            // ---- Hapus comment SECARA EKSPLISIT ----
            // ClearContents dan ClearFormats TIDAK menghapus comment.
            // Harus dipanggil tersendiri sebelum Clear lainnya.
            try { rng.ClearComments(); } catch { }

            rng.ClearContents();
            rng.ClearFormats();

            int[] sides = new[]
            {
                (int)Excel.XlBordersIndex.xlEdgeTop,
                (int)Excel.XlBordersIndex.xlEdgeBottom,
                (int)Excel.XlBordersIndex.xlEdgeLeft,
                (int)Excel.XlBordersIndex.xlEdgeRight,
                (int)Excel.XlBordersIndex.xlInsideHorizontal,
                (int)Excel.XlBordersIndex.xlInsideVertical
            };
            foreach (int s in sides)
                try { rng.Borders[(Excel.XlBordersIndex)s].LineStyle = Excel.XlLineStyle.xlLineStyleNone; }
                catch { }

            rng.Interior.Pattern = Excel.XlPattern.xlPatternNone;
        }

        // ================================================================
        // Filter helpers
        // ================================================================

        /// <summary>Cek apakah nilai kualitas = 5 (Macet), berbagai format</summary>
        private static bool IsKualitasMacet(string val)
        {
            if (string.IsNullOrEmpty(val)) return false;
            val = val.Trim();
            double d;
            if (double.TryParse(val, out d)) return (int)d == 5;

            // Cari digit pertama
            foreach (char c in val)
                if (char.IsDigit(c)) return c == '5';

            return false;
        }

        /// <summary>
        /// Cek jenis jaminan fisik: 092,161,162,163,176,177,187.
        /// Ekstrak blok digit pertama, normalisasi ke 3 digit.
        /// </summary>
        private static bool IsJaminanFisik(string val)
        {
            if (string.IsNullOrEmpty(val)) return false;
            val = val.Trim();

            string digits = "";
            bool started = false;
            foreach (char c in val)
            {
                if (char.IsDigit(c)) { digits += c; started = true; }
                else if (started)    break;
            }
            if (digits.Length == 0) return false;

            string code = digits.PadLeft(3, '0');
            if (code.Length > 3) code = code.Substring(code.Length - 3);

            switch (code)
            {
                case "092": case "161": case "162": case "163":
                case "176": case "177": case "187":
                    return true;
                default:
                    return false;
            }
        }

        // ================================================================
        // Audit Log
        // ================================================================
        private void TulisLog(
            Excel.Worksheet wsLog, int row,
            string periode, int tglAwal, int tglAkhir,
            string pathFile, string status,
            int nRekAwal, int nRekAkhir, double totalBaki)
        {
            ((Excel.Range)wsLog.Cells[row, "A"]).Value2 = DateTime.Now.ToOADate();
            ((Excel.Range)wsLog.Cells[row, "A"]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[row, "B"]).Value2 = "LGD-CS MACET";
            ((Excel.Range)wsLog.Cells[row, "C"]).Value2 = periode;
            if (tglAwal  > 0) { ((Excel.Range)wsLog.Cells[row, "D"]).Value2 = tglAwal;  ((Excel.Range)wsLog.Cells[row, "D"]).NumberFormat = "0"; }
            if (tglAkhir > 0) { ((Excel.Range)wsLog.Cells[row, "E"]).Value2 = tglAkhir; ((Excel.Range)wsLog.Cells[row, "E"]).NumberFormat = "0"; }
            ((Excel.Range)wsLog.Cells[row, "G"]).Value2 = pathFile;
            ((Excel.Range)wsLog.Cells[row, "H"]).Value2 = status;
            if (nRekAwal  > 0) ((Excel.Range)wsLog.Cells[row, "I"]).Value2 = nRekAwal;
            if (nRekAkhir > 0) ((Excel.Range)wsLog.Cells[row, "J"]).Value2 = nRekAkhir;
            if (totalBaki != 0)
            {
                ((Excel.Range)wsLog.Cells[row, "L"]).Value2 = totalBaki;
                ((Excel.Range)wsLog.Cells[row, "L"]).NumberFormat = "#,##0;(#,##0);-";
            }
        }

        private static int NextAuditLogRow(Excel.Worksheet wsLog)
        {
            Excel.Range lc = (Excel.Range)wsLog.Cells[wsLog.Rows.Count, "A"];
            Excel.Range ec = (Excel.Range)lc.End[Excel.XlDirection.xlUp];
            int last = (int)ec.Row;
            return last < 4 ? 5 : last + 1;
        }

        // ================================================================
        // Parsing dan utility
        // ================================================================

        /// <summary>
        /// Parse fileConfigStr: "label=tahun=bulan=path|label=tahun=bulan=path|..."
        /// bulan: 6 atau 12
        /// </summary>
        private static List<PeriodeCfg> ParsePeriodeCfg(string input)
        {
            var list = new List<PeriodeCfg>();
            foreach (var part in input.Split(new[]{'|'}, StringSplitOptions.RemoveEmptyEntries))
            {
                // Format: "label=tahun=bulan=path"
                // path mungkin mengandung '=' (drive letter), jadi split max 4
                string[] seg = part.Split(new[]{'='}, 4);
                if (seg.Length < 4) continue;
                int year, month;
                if (!int.TryParse(seg[1].Trim(), out year))  continue;
                if (!int.TryParse(seg[2].Trim(), out month)) continue;

                list.Add(new PeriodeCfg
                {
                    Label    = seg[0].Trim(),
                    RefYear  = year,
                    RefMonth = month,
                    RefKey   = year.ToString("0000") + month.ToString("00"),
                    Path     = seg[3].Trim()
                });
            }
            return list;
        }

        private static int CariLastRow(Excel.Worksheet ws, string col, int minRow)
        {
            Excel.Range lc = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            Excel.Range ec = (Excel.Range)lc.End[Excel.XlDirection.xlUp];
            int row = (int)ec.Row;
            return row < minRow ? minRow : row;
        }

        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }

        private void ApplyGreenBorderMacet(Excel.Range rng)
        {
            int[] sides = new[]
            {
                (int)Excel.XlBordersIndex.xlEdgeTop,
                (int)Excel.XlBordersIndex.xlEdgeBottom,
                (int)Excel.XlBordersIndex.xlEdgeLeft,
                (int)Excel.XlBordersIndex.xlEdgeRight,
                (int)Excel.XlBordersIndex.xlInsideHorizontal,
                (int)Excel.XlBordersIndex.xlInsideVertical
            };
            foreach (int s in sides)
                try
                {
                    Excel.Border b = rng.Borders[(Excel.XlBordersIndex)s];
                    b.LineStyle = Excel.XlLineStyle.xlContinuous;
                    b.Weight    = Excel.XlBorderWeight.xlThin;
                    b.Color     = CLR_GREEN_DARK;
                }
                catch { }
        }

        private void ApplyBorder(Excel.Range rng, Excel.XlLineStyle style, int color)
        {
            int[] sides = new[]
            {
                (int)Excel.XlBordersIndex.xlEdgeTop,
                (int)Excel.XlBordersIndex.xlEdgeBottom,
                (int)Excel.XlBordersIndex.xlEdgeLeft,
                (int)Excel.XlBordersIndex.xlEdgeRight
            };
            foreach (int s in sides)
                try
                {
                    Excel.Border b = rng.Borders[(Excel.XlBordersIndex)s];
                    b.LineStyle = style;
                    b.Weight    = Excel.XlBorderWeight.xlThin;
                    b.Color     = color;
                }
                catch { }
        }

        private void ApplyBorderDash(Excel.Range rng)
        {
            int[] sides = new[]
            {
                (int)Excel.XlBordersIndex.xlEdgeTop,
                (int)Excel.XlBordersIndex.xlEdgeBottom,
                (int)Excel.XlBordersIndex.xlEdgeLeft,
                (int)Excel.XlBordersIndex.xlEdgeRight,
                (int)Excel.XlBordersIndex.xlInsideHorizontal
            };
            foreach (int s in sides)
                try
                {
                    Excel.Border b = rng.Borders[(Excel.XlBordersIndex)s];
                    b.LineStyle = Excel.XlLineStyle.xlDash;
                    b.Weight    = Excel.XlBorderWeight.xlThin;
                    b.Color     = CLR_GRAY_LIGHT;
                }
                catch { }
        }
    }
}