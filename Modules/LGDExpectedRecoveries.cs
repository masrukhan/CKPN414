using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA modul "module lgd expected recoveries".
    /// Menghitung LGD berdasarkan recovery rate per cohort tahun hapus buku.
    ///
    /// Window: currYear-5 .. currYear (6 cohort)
    /// Output: sheet "B3.LGD-ER" dibangun ulang sepenuhnya
    /// </summary>
    internal class LGDExpectedRecoveries
    {
        private readonly Excel.Application _app;

        // Warna tema — sama persis dengan konstanta VBA
        private const int CLR_GREEN_DARK      = 0x1F6E43;  // &H436E1F  BGR→RGB
        private const int CLR_GREEN_MED       = 0x2E8B57;  // &H578B2E
        private const int CLR_GREEN_LIGHT     = 0xD4EBD4;  // &HDDEBD4
        private const int CLR_GREEN_SOFT      = 0xF1F8F4;  // &HF4F8F1
        private const int CLR_HIGHLIGHT_YELLOW= 0xFFFDE7;  // &HE7FDFF
        private const int CLR_WHITE           = 0xFFFFFF;
        private const int CLR_BLACK           = 0x000000;

        private const string FMT_NUM  = "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"_-;_-@";
        private const string FMT_PCT  = "0.00%";
        private const string FMT_YEAR = "0";
        private const int    N_YEARS  = 6;

        private const string SheetLGD    = "B3.LGD-ER";
        private const string SheetLog    = "Audit Log";
        private const string SheetMaster = "Master";
        private const string SheetKC2900 = "KC2900";
        private const int    KC2900Start = 3;

        public LGDExpectedRecoveries(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ----------------------------------------------------------------
        // Entry point
        // Parameter:
        //   currYear    : tahun berjalan (dari Master!C4)
        //   filePathsStr: "tahun1=path1|tahun2=path2|..." (6 pasang)
        //   sheetKCList : "KC0600,KC0700,..."
        // ----------------------------------------------------------------
        public void Hitung(int currYear, string filePathsStr, string sheetKCList)
        {
            var wb  = _app.ActiveWorkbook;
            var wsL = CariSheet(wb, SheetLGD);
            if (wsL == null)
                throw new InvalidOperationException("Sheet '" + SheetLGD + "' tidak ditemukan.");

            var wsLog = CariSheet(wb, SheetLog);
            if (wsLog == null)
                throw new InvalidOperationException("Sheet '" + SheetLog + "' tidak ditemukan.");

            // Parse file paths: "2020=F:\path\file.xlsx|2021=..."
            var filePaths = ParseFilePaths(filePathsStr);
            if (filePaths.Count == 0)
                throw new ArgumentException("filePathsStr tidak valid: " + filePathsStr);

            // Validasi semua file ada
            foreach (var kv in filePaths)
                if (!System.IO.File.Exists(kv.Value))
                    throw new System.IO.FileNotFoundException(
                        "File tahun " + kv.Key + " tidak ditemukan: " + kv.Value);

            // Parse scope KC
            var scopeKC = new HashSet<string>(
                sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);

            // ---- Auto-build RefKC dan isi kolom M KC2900 ----
            // Kolom M (Kode KC) di KC2900 wajib terisi agar filter KC bekerja.
            // Sistem cek otomatis: jika kolom M belum terisi di salah satu file,
            // jalankan RefKCBuilder secara otomatis sebelum membaca data.
            // Ini menggantikan keharusan user menjalankan "Bangun Tabel Referensi KC"
            // secara manual terlebih dahulu.
            var refBuilder = new RefKCBuilder(_app);
            bool kolomMOK  = refBuilder.CekKolomMTerisi(filePaths);

            if (!kolomMOK)
            {
                // Kolom M belum terisi — bangun otomatis
                // Tampilkan info ke user bahwa proses tambahan sedang berjalan
                System.Windows.Forms.MessageBox.Show(
                    "Kolom referensi KC (M) di KC2900 belum terisi." +
                    "Sistem akan membangun tabel referensi KC secara otomatis" +
                    "sebelum melanjutkan perhitungan LGD Expected Recoveries." +
                    "File sumber akan dibuka READ-WRITE dan disimpan." +
                    "Proses ini hanya perlu dilakukan sekali per set file.",
                    "Auto-Build Referensi KC",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                refBuilder.BangunRefKC(filePaths, wsLog);
            }

            // ---- Baca KC2900 dari setiap file tahunan ----
            // dataByYear[year]    = Dictionary<noRek, double[]{ baki, tglHB }>
            // rekPerTahun[year]   = jumlah rekening KC2900 per file
            // bakiPerTahun[year]  = total baki KC2900 per file
            // statRefKC: rekap Jenis Instrumen x Kode KC lintas file, dikumpulkan
            // SEBELUM filter scopeKC agar rekening tanpa Kode KC tetap terlihat
            var statRefKC = new StatistikRefKC();

            var dataByYear   = new Dictionary<int, Dictionary<string, double[]>>();
            var rekPerTahun  = new Dictionary<int, int>();
            var bakiPerTahun = new Dictionary<int, double>();
            var pathPerTahun = new Dictionary<int, string>();

            foreach (var kv in filePaths)
            {
                pathPerTahun[kv.Key] = kv.Value;
                Excel.Workbook srcWb = null;
                try
                {
                    srcWb = _app.Workbooks.Open(kv.Value, UpdateLinks: 0, ReadOnly: true);
                    var dataYear = BacaKC2900(srcWb, scopeKC, kv.Key, statRefKC);
                    dataByYear[kv.Key]   = dataYear;
                    rekPerTahun[kv.Key]  = dataYear.Count;

                    // Hitung total baki per tahun untuk log
                    double totalBakiYear = 0;
                    foreach (var d in dataYear.Values) totalBakiYear += d[0];
                    bakiPerTahun[kv.Key] = totalBakiYear;

                    srcWb.Close(false); srcWb = null;
                }
                finally { if (srcWb != null) try { srcWb.Close(false); } catch { } }
            }

            // ---- Klasifikasi rekening ke cohort ----
            int minYear = currYear - (N_YEARS - 1);
            var accountMeta = new Dictionary<string, int[]>(StringComparer.OrdinalIgnoreCase);
            // cohortRek[cohortYear] = HashSet<noRek>
            var cohortRek = new Dictionary<int, HashSet<string>>();

            // Key di dataByYear sudah dinormalisasi oleh BacaKC2900 via NormNoRek.
            // Iterasi langsung aman — tidak perlu normalisasi ulang di sini.
            foreach (var yk in dataByYear.Keys)
            {
                foreach (var kv in dataByYear[yk])
                {
                    string noRek = kv.Key;   // sudah dinormalisasi
                    if (accountMeta.ContainsKey(noRek)) continue;

                    double[] meta   = kv.Value;
                    long     tglHB  = (long)meta[1];
                    int      hbYear = tglHB >= 19000101 ? (int)(tglHB / 10000) : yk;

                    if (hbYear >= minYear && hbYear <= currYear)
                    {
                        accountMeta[noRek] = new[]{ hbYear, (int)tglHB };
                        if (!cohortRek.ContainsKey(hbYear))
                            cohortRek[hbYear] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        cohortRek[hbYear].Add(noRek);
                    }
                }
            }

            // ---- Reset sheet output ----
            int lastUsed = CariLastRow(wsL, "C", 4);
            if (lastUsed < 200) lastUsed = 200;
            ((Excel.Range)wsL.Range["B4", "O" + lastUsed]).Clear();

            // ---- Header utama ----
            TulisHeaderUtama(wsL, currYear, minYear);

            // ---- Bangun tiap cohort block ----
            int    curRow       = 5;
            string subtotalRows = "";

            for (int ci = 0; ci < N_YEARS; ci++)
            {
                int cohortYear = minYear + ci;
                curRow = BangunCohortBlock(
                    wsL, cohortYear, currYear, ci,
                    cohortRek, dataByYear,
                    curRow, ref subtotalRows);
                curRow++;
            }

            // ---- Grand total ----
            TulisGrandTotal(wsL, curRow, subtotalRows);

            // ---- Lebar kolom ----
            ((Excel.Range)wsL.Columns["B"]).ColumnWidth = 5;
            ((Excel.Range)wsL.Columns["C"]).ColumnWidth = 28;
            ((Excel.Range)wsL.Columns["D"]).ColumnWidth = 12;
            ((Excel.Range)wsL.Columns["E"]).ColumnWidth = 16;
            ((Excel.Range)wsL.Range["F:K", wsL.Range["F:K"]]).ColumnWidth = 14;
            ((Excel.Range)wsL.Columns["L"]).ColumnWidth = 10;
            ((Excel.Range)wsL.Columns["M"]).ColumnWidth = 10;
            ((Excel.Range)wsL.Columns["N"]).ColumnWidth = 16;
            ((Excel.Range)wsL.Columns["O"]).ColumnWidth = 12;

            _app.Calculate();

            // ---- Audit Log ----
            // Satu baris per file tahunan + satu baris ringkasan
            try
            {
                TulisAuditLog(wsLog, currYear, minYear,
                              filePaths, pathPerTahun,
                              rekPerTahun, bakiPerTahun,
                              accountMeta, cohortRek, sheetKCList);
            }
            catch { /* abaikan error logging — jangan gagalkan proses utama */ }

            // ---- Audit Log: rekap Jenis Instrumen x Kode KC ----
            try { TulisLogRefKC(wsLog, statRefKC); }
            catch { /* abaikan error logging — jangan gagalkan proses utama */ }

            // ---- Pesan selesai ----
            string pesan = "Selesai.\nTotal rekening: " + accountMeta.Count +
                           " dalam " + cohortRek.Count + " cohort.";

            if (statRefKC.TotalTanpaKC > 0)
            {
                pesan += "\n\nPERHATIAN: " + statRefKC.TotalTanpaKC +
                         " rekening tidak punya Kode KC (kolom M) senilai " +
                         statRefKC.BakiTanpaKC.ToString("N0") + "." +
                         "\nRekening ini TIDAK ikut dihitung." +
                         "\nJenis instrumen: " + statRefKC.RingkasProdukTanpaKC() +
                         "\nRincian per tahun ada di sheet 'Audit Log' " +
                         "(baris \"LGD ER - Referensi KC\").";
            }

            System.Windows.Forms.MessageBox.Show(
                pesan, "LGD Expected Recoveries",
                System.Windows.Forms.MessageBoxButtons.OK,
                statRefKC.TotalTanpaKC > 0
                    ? System.Windows.Forms.MessageBoxIcon.Warning
                    : System.Windows.Forms.MessageBoxIcon.Information);
        }

        // ================================================================
        // Baca KC2900 dari satu workbook tahunan
        // Return: Dictionary<noRek, double[]{ baki, tglHB }>
        // ================================================================
        private Dictionary<string, double[]> BacaKC2900(
            Excel.Workbook wb, HashSet<string> scopeKC, int tahun, StatistikRefKC stat)
        {
            var d = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            if (!ExcelHelper.SheetAda(wb, SheetKC2900)) return d;

            var ws      = (Excel.Worksheet)wb.Worksheets[SheetKC2900];
            int lastRow = ExcelHelper.CariLastRow(ws, "H", KC2900Start - 1);
            if (lastRow < KC2900Start) return d;

            // Bulk read: G=JenisInstrumen, H=NoRek, I=TglHB, L=Baki, M=KodeKC
            string[] prodArr = ExcelHelper.BacaKolomString(ws, "G", KC2900Start, lastRow);
            string[] rekArr  = ExcelHelper.BacaKolomString(ws, "H", KC2900Start, lastRow);
            double[] tglArr  = ExcelHelper.BacaKolomDouble(ws, "I", KC2900Start, lastRow);
            double[] bakiArr = ExcelHelper.BacaKolomDouble(ws, "L", KC2900Start, lastRow);
            string[] refArr  = ExcelHelper.BacaKolomString(ws, "M", KC2900Start, lastRow);

            for (int i = 0; i < rekArr.Length; i++)
            {
                // NormNoRek: hapus apostrof prefix, trim spasi, strip leading-zero
                // agar matching konsisten antar file tahun yang berbeda format
                string noRek = NormNoRek(rekArr[i]);
                if (string.IsNullOrEmpty(noRek)) continue;

                string refKC = NormKodeKC(refArr[i]);

                // Rekap dicatat SEBELUM filter scope, sehingga rekening yang
                // gagal dipetakan (kode konversi CBS lama / WO lama) tetap
                // muncul di Audit Log dan bisa dilengkapi manual oleh user.
                if (stat != null)
                    stat.Catat(tahun, prodArr[i], refKC, noRek, bakiArr[i]);

                if (!scopeKC.Contains(refKC)) continue;

                // Jika noRek sudah ada (duplikat), akumulasi baki
                if (d.ContainsKey(noRek))
                    d[noRek][0] += bakiArr[i];
                else
                    d[noRek] = new[]{ bakiArr[i], tglArr[i] };
            }
            return d;
        }

        private double GetBaki(Dictionary<int, Dictionary<string, double[]>> dataByYear,
                               int yr, string noRek)
        {
            if (!dataByYear.ContainsKey(yr)) return 0;
            // Normalisasi noRek sebelum lookup — key di dictionary sudah dinormalisasi
            // saat BacaKC2900 dipanggil. Normalisasi di sini memastikan konsistensi
            // meskipun noRek berasal dari sumber yang berbeda format.
            string normRek = NormNoRek(noRek);
            double[] v;
            if (dataByYear[yr].TryGetValue(normRek, out v)) return v[0];
            return 0;
        }

        // ================================================================
        // Tulis header utama baris 3-4
        // ================================================================
        private void TulisHeaderUtama(Excel.Worksheet ws, int currYear, int minYear)
        {
            Excel.Range r3 = (Excel.Range)ws.Range["B3", "O3"];
            r3.Merge();
            r3.Value2 = "B3. LGD — Expected Recoveries";
            r3.Font.Bold = true; r3.Font.Size = 13;
            r3.Font.Color = CLR_WHITE;
            r3.Interior.Color = CLR_GREEN_DARK;
            r3.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            r3.RowHeight = 22;

            Excel.Range c4 = (Excel.Range)ws.Range["C4", "D4"];
            c4.Merge();
            c4.Value2 = "Tahun hapus buku berjalan";
            c4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            c4.Interior.Color = CLR_GREEN_MED;
            c4.Font.Color = CLR_WHITE; c4.Font.Bold = true;

            // E4 = oldest year (referensi ke Master!F60)
            Excel.Range e4 = (Excel.Range)ws.Range["E4"];
            e4.Formula = "=Master!$F$60";
            e4.NumberFormat = FMT_YEAR; e4.Font.Bold = true;
            e4.Font.Color = CLR_GREEN_DARK;
            e4.Interior.Color = CLR_HIGHLIGHT_YELLOW;
            e4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            for (int j = 0; j <= 5; j++)
            {
                Excel.Range cell = (Excel.Range)ws.Cells[4, 6 + j];
                cell.Formula = "=Master!$F$" + (60 + j);
                cell.NumberFormat = FMT_YEAR;
            }
            Excel.Range f4k4 = (Excel.Range)ws.Range["F4", "K4"];
            f4k4.Interior.Color = CLR_GREEN_LIGHT;
            f4k4.Font.Bold = true; f4k4.Font.Color = CLR_GREEN_DARK;
            f4k4.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ApplyGreenBorder((Excel.Range)ws.Range["C4", "K4"]);
        }

        // ================================================================
        // Bangun satu blok cohort — setara BuildCohortBlock VBA
        // ================================================================
        private int BangunCohortBlock(
            Excel.Worksheet ws,
            int cohortYear, int currYear, int cohortIdx,
            Dictionary<int, HashSet<string>> cohortRek,
            Dictionary<int, Dictionary<string, double[]>> dataByYear,
            int startRow, ref string subtotalRows)
        {
            bool isCurrent   = (cohortYear == currYear);
            int  headerRow   = startRow;
            int  yearRow     = startRow + 1;
            int  labelRow    = startRow + 2;
            int  firstDataRow= startRow + 3;
            int  masterRowRef= 60 + cohortIdx;

            HashSet<string> listRek;
            if (!cohortRek.TryGetValue(cohortYear, out listRek))
                listRek = new HashSet<string>();

            // ---- Header baris 1 ----
            Excel.Range hdrCD = (Excel.Range)ws.Range["C" + headerRow, "D" + headerRow];
            hdrCD.Merge();
            hdrCD.Formula = "=\"Tahun hapus buku \" & Master!$F$" + masterRowRef;
            hdrCD.Font.Bold = true; hdrCD.Font.Color = CLR_WHITE;
            hdrCD.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            hdrCD.Interior.Color = CLR_GREEN_DARK;

            Excel.Range hdrE = (Excel.Range)ws.Cells[headerRow, "E"];
            hdrE.Formula = "=Master!$F$" + masterRowRef;
            hdrE.NumberFormat = FMT_YEAR; hdrE.Font.Bold = true;
            hdrE.Font.Color = CLR_WHITE; hdrE.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            hdrE.Interior.Color = CLR_GREEN_MED;
            ((Excel.Range)ws.Range["F" + headerRow, "O" + headerRow]).Interior.Color = CLR_GREEN_MED;

            // ---- Year row ----
            Excel.Range yearRng = (Excel.Range)ws.Range["C" + yearRow, "O" + yearRow];
            yearRng.Interior.Color = CLR_GREEN_LIGHT;
            yearRng.Font.Bold = true; yearRng.Font.Color = CLR_GREEN_DARK;
            yearRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            for (int jj = 0; jj <= 5; jj++)
            {
                Excel.Range cell = (Excel.Range)ws.Cells[yearRow, 6 + jj];
                if (cohortYear + jj <= currYear) cell.Value2 = cohortYear + jj;
                cell.NumberFormat = FMT_YEAR;
            }

            // ---- Label row ----
            Excel.Range lblRng = (Excel.Range)ws.Range["C" + labelRow, "O" + labelRow];
            lblRng.Interior.Color = CLR_GREEN_LIGHT;
            lblRng.Font.Bold = true; lblRng.Font.Color = CLR_GREEN_DARK;
            lblRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ((Excel.Range)ws.Cells[labelRow, "C"]).Value2 = "Nama Nasabah";
            ((Excel.Range)ws.Cells[labelRow, "D"]).Value2 = "Tahun Hapus Buku";
            ((Excel.Range)ws.Cells[labelRow, "E"]).Value2 = "Balance hapus buku";
            for (int j = 0; j <= 5; j++)
                ((Excel.Range)ws.Cells[labelRow, 6 + j]).Value2 = "year " + (j + 1);
            ((Excel.Range)ws.Cells[labelRow, "N"]).Value2 = "Jumlah recovery";
            ((Excel.Range)ws.Cells[labelRow, "O"]).Value2 = "Recovery Rate";

            // ---- Data rows ----
            int outRow = firstDataRow;

            if (listRek.Count > 0)
            {
                foreach (string noRek in listRek)
                {
                    double[] baki  = new double[7];   // index 1..6
                    double[] recov = new double[7];   // index 2..6
                    bool[]   hasYr = new bool[7];

                    for (int j = 1; j <= 6; j++)
                        baki[j] = cohortYear + j - 1 <= currYear
                            ? GetBaki(dataByYear, cohortYear + j - 1, noRek)
                            : 0;

                    for (int j = 2; j <= 6; j++)
                    {
                        if (cohortYear + j - 1 <= currYear)
                        {
                            hasYr[j] = true;
                            recov[j] = baki[j - 1] > baki[j] ? baki[j - 1] - baki[j] : 0;
                        }
                    }

                    // Baris saldo
                    ((Excel.Range)ws.Cells[outRow, "C"]).NumberFormat = "@";
                    ((Excel.Range)ws.Cells[outRow, "C"]).Value2 = noRek;
                    ((Excel.Range)ws.Cells[outRow, "D"]).Formula = "=Master!$F$" + masterRowRef;
                    ((Excel.Range)ws.Cells[outRow, "D"]).NumberFormat = FMT_YEAR;
                    ((Excel.Range)ws.Cells[outRow, "D"]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ((Excel.Range)ws.Cells[outRow, "E"]).Value2 = baki[1];
                    ((Excel.Range)ws.Cells[outRow, "E"]).NumberFormat = FMT_NUM;

                    for (int j = 1; j <= 6; j++)
                    {
                        Excel.Range cell = (Excel.Range)ws.Cells[outRow, 5 + j];
                        if (cohortYear + j - 1 <= currYear) cell.Value2 = baki[j];
                        cell.NumberFormat = FMT_NUM;
                    }

                    // Baris recovery
                    ((Excel.Range)ws.Cells[outRow + 1, "C"]).Value2 = "-";
                    ((Excel.Range)ws.Cells[outRow + 1, "C"]).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ((Excel.Range)ws.Cells[outRow + 1, "E"]).Value2 = "recovery";
                    ((Excel.Range)ws.Cells[outRow + 1, "E"]).Font.Italic = true;

                    if (!isCurrent)
                    {
                        for (int j = 2; j <= 6; j++)
                        {
                            Excel.Range cell = (Excel.Range)ws.Cells[outRow + 1, 4 + j];
                            if (hasYr[j]) cell.Value2 = recov[j];
                            cell.NumberFormat = FMT_NUM;
                        }
                        ((Excel.Range)ws.Cells[outRow + 1, "N"]).Formula =
                            "=SUM(F" + (outRow + 1) + ":K" + (outRow + 1) + ")";
                        ((Excel.Range)ws.Cells[outRow + 1, "N"]).NumberFormat = FMT_NUM;
                        ((Excel.Range)ws.Cells[outRow + 1, "O"]).Formula =
                            "=IFERROR(N" + (outRow + 1) + "/E" + outRow + ",0)";
                        ((Excel.Range)ws.Cells[outRow + 1, "O"]).NumberFormat = FMT_PCT;
                    }
                    outRow += 2;
                }
            }
            else
            {
                ((Excel.Range)ws.Cells[outRow,     "C"]).Value2 = "(tidak ada data)";
                ((Excel.Range)ws.Cells[outRow,     "D"]).Formula = "=Master!$F$" + masterRowRef;
                ((Excel.Range)ws.Cells[outRow,     "D"]).NumberFormat = FMT_YEAR;
                ((Excel.Range)ws.Cells[outRow + 1, "C"]).Value2 = "-";
                ((Excel.Range)ws.Cells[outRow + 1, "E"]).Value2 = "recovery";
                outRow += 2;
            }

            // ---- Subtotal combine ----
            int subRow        = outRow;
            int firstSaldoRow = firstDataRow;
            int lastSaldoRow  = subRow - 1;

            ((Excel.Range)ws.Cells[subRow, "C"]).Value2 = "combine";
            ((Excel.Range)ws.Cells[subRow, "C"]).Font.Bold = true;
            ((Excel.Range)ws.Cells[subRow, "E"]).Formula =
                "=SUM(E" + firstSaldoRow + ":E" + lastSaldoRow + ")";
            ((Excel.Range)ws.Cells[subRow, "E"]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Cells[subRow, "E"]).Font.Bold = true;

            for (int j = 6; j <= 11; j++)  // F..K
            {
                string colLet = ((char)(64 + j)).ToString();
                ((Excel.Range)ws.Cells[subRow, j]).Formula =
                    "=SUMPRODUCT((MOD(ROW(" + colLet + firstSaldoRow + ":" + colLet + lastSaldoRow +
                    ")-" + firstSaldoRow + ",2)=0)*" + colLet + firstSaldoRow + ":" + colLet + lastSaldoRow + ")";
                ((Excel.Range)ws.Cells[subRow, j]).NumberFormat = FMT_NUM;
                ((Excel.Range)ws.Cells[subRow, j]).Font.Bold = true;
            }
            ((Excel.Range)ws.Cells[subRow, "N"]).Formula =
                "=SUM(N" + firstSaldoRow + ":N" + lastSaldoRow + ")";
            ((Excel.Range)ws.Cells[subRow, "N"]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Cells[subRow, "N"]).Font.Bold = true;
            ((Excel.Range)ws.Cells[subRow, "O"]).Formula =
                "=IFERROR(N" + subRow + "/E" + subRow + ",0)";
            ((Excel.Range)ws.Cells[subRow, "O"]).NumberFormat = FMT_PCT;
            ((Excel.Range)ws.Cells[subRow, "O"]).Font.Bold = true;

            Excel.Range subRng = (Excel.Range)ws.Range["C" + subRow, "O" + subRow];
            subRng.Interior.Color = CLR_GREEN_SOFT;
            subRng.Font.Bold = true; subRng.Font.Color = CLR_GREEN_DARK;

            ApplyGreenBorder((Excel.Range)ws.Range["C" + headerRow, "O" + subRow]);

            // Block warna untuk year > currYear
            for (int j = 0; j <= 5; j++)
            {
                if (cohortYear + j > currYear)
                {
                    Excel.Range blocked = (Excel.Range)ws.Range[
                        ws.Cells[firstDataRow, 6 + j],
                        ws.Cells[subRow, 6 + j]];
                    blocked.Interior.Color = CLR_GREEN_DARK;
                    blocked.Font.Color     = CLR_WHITE;
                }
            }

            if (subtotalRows.Length > 0) subtotalRows += ",";
            subtotalRows += subRow;

            return subRow;
        }

        // ================================================================
        // Grand total dan ringkasan LGD
        // ================================================================
        private void TulisGrandTotal(Excel.Worksheet ws, int startRow, string subtotalRows)
        {
            int r = startRow + 1;
            string[] parts = subtotalRows.Split(',');
            string sumE = "", sumN = "";
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0) { sumE += "+"; sumN += "+"; }
                sumE += "E" + parts[i];
                sumN += "N" + parts[i];
            }

            ((Excel.Range)ws.Cells[r, "C"]).Value2 = "TOTAL COMBINE";
            ((Excel.Range)ws.Cells[r, "E"]).Formula = "=" + sumE;
            ((Excel.Range)ws.Cells[r, "E"]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Cells[r, "N"]).Formula = "=" + sumN;
            ((Excel.Range)ws.Cells[r, "N"]).NumberFormat = FMT_NUM;
            ((Excel.Range)ws.Cells[r, "O"]).Formula = "=IFERROR(N" + r + "/E" + r + ",0)";
            ((Excel.Range)ws.Cells[r, "O"]).NumberFormat = FMT_PCT;

            Excel.Range totalRng = (Excel.Range)ws.Range["C" + r, "O" + r];
            totalRng.Interior.Color = CLR_GREEN_DARK;
            totalRng.Font.Color = CLR_WHITE; totalRng.Font.Bold = true;
            totalRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            totalRng.RowHeight = 20;
            ((Excel.Range)ws.Cells[r, "C"]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            ApplyGreenBorder(totalRng, Excel.XlBorderWeight.xlMedium);

            // Ringkasan LGD Weighted
            int r2 = r + 2;
            Excel.Range lblRng = (Excel.Range)ws.Range["C" + r2, "D" + r2];
            lblRng.Merge();
            lblRng.Value2 = "Ringkasan LGD Weighted";
            lblRng.Font.Bold = true; lblRng.Font.Color = CLR_WHITE;
            lblRng.Interior.Color = CLR_GREEN_MED;
            lblRng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            ((Excel.Range)ws.Cells[r2 + 1, "C"]).Value2 = "Total WO";
            ((Excel.Range)ws.Cells[r2 + 1, "D"]).Formula = "=E" + r;
            ((Excel.Range)ws.Cells[r2 + 1, "D"]).NumberFormat = FMT_NUM;

            ((Excel.Range)ws.Cells[r2 + 2, "C"]).Value2 = "Total Recovery";
            ((Excel.Range)ws.Cells[r2 + 2, "D"]).Formula = "=N" + r;
            ((Excel.Range)ws.Cells[r2 + 2, "D"]).NumberFormat = FMT_NUM;

            ((Excel.Range)ws.Cells[r2 + 3, "C"]).Value2 = "LGD Weighted";
            ((Excel.Range)ws.Cells[r2 + 3, "D"]).Formula =
                "=IFERROR(1-D" + (r2 + 2) + "/D" + (r2 + 1) + ",0)";
            ((Excel.Range)ws.Cells[r2 + 3, "D"]).NumberFormat = FMT_PCT;

            Excel.Range lblCols = (Excel.Range)ws.Range["C" + (r2 + 1), "C" + (r2 + 3)];
            lblCols.Interior.Color = CLR_GREEN_LIGHT;
            lblCols.Font.Color = CLR_GREEN_DARK; lblCols.Font.Bold = true;
            lblCols.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            lblCols.IndentLevel = 1;

            Excel.Range valCols = (Excel.Range)ws.Range["D" + (r2 + 1), "D" + (r2 + 3)];
            valCols.Interior.Color = CLR_GREEN_SOFT;
            valCols.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            Excel.Range lgdCell = (Excel.Range)ws.Cells[r2 + 3, "D"];
            lgdCell.Interior.Color = CLR_HIGHLIGHT_YELLOW;
            lgdCell.Font.Color = CLR_GREEN_DARK;
            ((Excel.Range)ws.Range["C" + (r2 + 3), "D" + (r2 + 3)]).Font.Bold = true;
            ((Excel.Range)ws.Range["C" + (r2 + 3), "D" + (r2 + 3)]).Font.Size = 12;

            ApplyGreenBorder((Excel.Range)ws.Range["C" + r2, "D" + (r2 + 3)],
                             Excel.XlBorderWeight.xlMedium);
        }

        // ================================================================
        // Utility
        // ================================================================

        /// <summary>Parse "2020=path1|2021=path2|..." ke Dictionary</summary>
        private static Dictionary<int, string> ParseFilePaths(string input)
        {
            var dict = new Dictionary<int, string>();
            foreach (var part in input.Split(new[]{'|'}, StringSplitOptions.RemoveEmptyEntries))
            {
                int idx = part.IndexOf('=');
                if (idx < 1) continue;
                int year;
                if (int.TryParse(part.Substring(0, idx).Trim(), out year))
                    dict[year] = part.Substring(idx + 1).Trim();
            }
            return dict;
        }

        private void ApplyGreenBorder(Excel.Range rng,
            Excel.XlBorderWeight weight = Excel.XlBorderWeight.xlThin)
        {
            int[] sides = new[]
            {
                (int)Excel.XlBordersIndex.xlEdgeLeft,
                (int)Excel.XlBordersIndex.xlEdgeRight,
                (int)Excel.XlBordersIndex.xlEdgeTop,
                (int)Excel.XlBordersIndex.xlEdgeBottom,
                (int)Excel.XlBordersIndex.xlInsideHorizontal,
                (int)Excel.XlBordersIndex.xlInsideVertical
            };
            foreach (int s in sides)
            {
                try
                {
                    Excel.Border b = rng.Borders[(Excel.XlBordersIndex)s];
                    b.LineStyle = Excel.XlLineStyle.xlContinuous;
                    b.Weight    = weight;
                    b.Color     = CLR_GREEN_DARK;
                }
                catch { }
            }
        }

        // ================================================================
        // NormNoRek: normalisasi nomor rekening untuk matching konsisten
        //
        // Masalah yang diatasi:
        //   1. Apostrof prefix: "'001234" → "001234"
        //      Excel menyimpan angka sebagai teks dengan prefix apostrof
        //   2. Spasi di awal/akhir: " 001234 " → "001234"
        //   3. Leading-zero tidak konsisten antar file tahun:
        //      "1234" di file lama vs "001234" di file baru
        //      → keduanya dinormalisasi ke representasi numerik "1234"
        //      KECUALI jika seluruh string bukan numerik murni
        //      (mis. "KC-001234") → tetap pakai string asli setelah trim
        //
        // Catatan: leading-zero dihapus HANYA jika noRek adalah angka murni.
        // Format alfanumerik (mis. "KP0001234") tidak diubah.
        // ================================================================
        private static string NormNoRek(string raw)
        {
            if (raw == null) return "";

            // Hapus apostrof prefix dan spasi
            string s = raw.Trim();
            if (s.StartsWith("'")) s = s.Substring(1).Trim();
            if (s.Length == 0) return "";

            // Jika murni angka → strip leading-zero untuk keseragaman
            // "001234" dan "1234" akan match ke "1234"
            bool allDigit = true;
            foreach (char c in s)
                if (!char.IsDigit(c)) { allDigit = false; break; }

            if (allDigit)
            {
                // TrimStart('0') — tapi jaga jika semua nol ("000" → "0")
                string stripped = s.TrimStart('0');
                return stripped.Length == 0 ? "0" : stripped;
            }

            // Alfanumerik: kembalikan apa adanya (sudah trim + no apostrof)
            return s;
        }

        private static int CariLastRow(Excel.Worksheet ws, string col, int minRow)
        {
            Excel.Range lc = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            Excel.Range ec = (Excel.Range)lc.End[Excel.XlDirection.xlUp];
            int row = (int)ec.Row;
            return row < minRow ? minRow : row;
        }

        // ================================================================
        // TulisAuditLog: catat ringkasan LGD ER ke sheet Audit Log
        //
        // Pola penulisan: SATU BARIS per file tahunan (detail per tahun)
        // + SATU BARIS RINGKASAN di akhir (total keseluruhan)
        //
        // Kolom A-Y sesuai standar Audit Log:
        //   A  = Timestamp
        //   B  = Jenis Proses ("LGD Expected Recoveries")
        //   C  = Periode/Label ("Des YYYY")
        //   D  = Tgl Ref / Tahun file (angka)
        //   E  = Tgl Akhir / currYear (di baris ringkasan)
        //   F  = Path Awal (kosong untuk LGD-ER, satu file per tahun)
        //   G  = Path File tahunan
        //   H  = Status ("OK" / "File kosong")
        //   I  = #Rek Awal / #Rek KC2900 di file ini
        //   J  = #Rek Akhir (di baris ringkasan = total rekening masuk cohort)
        //   K  = #Rek WO (di baris ringkasan = total cohort)
        //   L  = Total Saldo Awal (total baki KC2900 file ini)
        //   S  = Ruang Lingkup KC (di baris ringkasan)
        // ================================================================
        private void TulisAuditLog(
            Excel.Worksheet wsLog,
            int currYear, int minYear,
            Dictionary<int, string> filePaths,
            Dictionary<int, string> pathPerTahun,
            Dictionary<int, int>    rekPerTahun,
            Dictionary<int, double> bakiPerTahun,
            Dictionary<string, int[]> accountMeta,
            Dictionary<int, HashSet<string>> cohortRek,
            string sheetKCList)
        {
            // Cari baris kosong berikutnya di Audit Log
            int nextRow = NextAuditLogRow(wsLog);
            double ts   = DateTime.Now.ToOADate();

            // ── Detail per file tahunan ──────────────────────────────
            // Urutkan tahun ascending (tua ke baru)
            var sortedYears = new List<int>(filePaths.Keys);
            sortedYears.Sort();

            int totalRekLog   = 0;
            double totalBakiLog = 0;

            foreach (int yr in sortedYears)
            {
                string path   = filePaths.ContainsKey(yr) ? filePaths[yr] : "";
                int    nRek   = rekPerTahun.ContainsKey(yr)  ? rekPerTahun[yr]  : 0;
                double nBaki  = bakiPerTahun.ContainsKey(yr) ? bakiPerTahun[yr] : 0;
                string status = nRek > 0 ? "OK" : "File kosong / tidak ada data KC2900";

                // Hitung jumlah rekening di cohort tahun ini
                int nCohort = cohortRek.ContainsKey(yr) ? cohortRek[yr].Count : 0;

                // A: Timestamp
                ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = ts;
                ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
                // B: Jenis Proses
                ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "LGD Expected Recoveries";
                // C: Periode/Label
                ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = "Des " + yr;
                // D: Tahun file (Tgl Ref)
                ((Excel.Range)wsLog.Cells[nextRow, 4]).Value2 = yr;
                // G: Path File
                ((Excel.Range)wsLog.Cells[nextRow, 7]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[nextRow, 7]).Value2 = path;
                // H: Status
                ((Excel.Range)wsLog.Cells[nextRow, 8]).Value2 = status;
                // I: #Rek KC2900 di file ini
                ((Excel.Range)wsLog.Cells[nextRow, 9]).Value2 = nRek;
                // J: #Rek masuk cohort tahun ini
                ((Excel.Range)wsLog.Cells[nextRow, 10]).Value2 = nCohort;
                // L: Total Baki KC2900
                ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = nBaki;
                ((Excel.Range)wsLog.Cells[nextRow, 12]).NumberFormat = "#,##0;(#,##0);-";

                totalRekLog   += nRek;
                totalBakiLog  += nBaki;
                nextRow++;
            }

            // ── Baris ringkasan keseluruhan ──────────────────────────
            // Total cohort dan rekening lintas semua tahun
            int totalRekCohort = accountMeta.Count;
            int totalCohort    = cohortRek.Count;

            ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = ts;
            ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "LGD Expected Recoveries";
            // C: Label ringkasan
            ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = "RINGKASAN " + minYear + "-" + currYear;
            // D: Tahun tertua (minYear)
            ((Excel.Range)wsLog.Cells[nextRow, 4]).Value2 = minYear;
            // E: currYear
            ((Excel.Range)wsLog.Cells[nextRow, 5]).Value2 = currYear;
            // H: Status
            ((Excel.Range)wsLog.Cells[nextRow, 8]).Value2 = "OK - " + totalCohort + " cohort";
            // I: Total #Rek KC2900 dari semua file
            ((Excel.Range)wsLog.Cells[nextRow, 9]).Value2 = totalRekLog;
            // J: Total rekening masuk cohort (lintas semua tahun)
            ((Excel.Range)wsLog.Cells[nextRow, 10]).Value2 = totalRekCohort;
            // K: Jumlah cohort aktif
            ((Excel.Range)wsLog.Cells[nextRow, 11]).Value2 = totalCohort;
            // L: Total baki KC2900 lintas semua file
            ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = totalBakiLog;
            ((Excel.Range)wsLog.Cells[nextRow, 12]).NumberFormat = "#,##0;(#,##0);-";
            // S: Ruang lingkup KC (kolom 19)
            ((Excel.Range)wsLog.Cells[nextRow, 19]).Value2 = sheetKCList;

            // Warna baris ringkasan agar mudah dibedakan
            Excel.Range ringkasanRng = (Excel.Range)wsLog.Range[
                wsLog.Cells[nextRow, 1], wsLog.Cells[nextRow, 19]];
            ringkasanRng.Interior.Color = 0xD4EBD4;  // hijau muda
            ringkasanRng.Font.Bold      = true;
        }

        private static int NextAuditLogRow(Excel.Worksheet wsLog)
        {
            for (int r = 5; r <= 1000000; r++)
            {
                string a = (((Excel.Range)wsLog.Cells[r, 1]).Value2 ?? "").ToString().Trim();
                string b = (((Excel.Range)wsLog.Cells[r, 2]).Value2 ?? "").ToString().Trim();
                if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return r;
            }
            return 1000001;
        }

        // ================================================================
        // NormKodeKC: normalisasi isi kolom M (Kode KC) di KC2900
        //
        // Kolom M diisi otomatis oleh RefKCBuilder dengan cara memetakan
        // nomor kontrak ke sheet KC di file sumber. Pemetaan bisa GAGAL untuk:
        //   - nomor kontrak yang sudah lama dihapus buku sehingga tidak lagi
        //     muncul di sheet KC mana pun
        //   - nomor kontrak hasil konversi dari CBS lama dengan format berbeda
        //
        // Saat gagal, sel bisa berisi kosong, spasi, "FALSE" (hasil formula
        // boolean), "#N/A", atau "0". Semuanya dianggap TIDAK TERPETAKAN dan
        // dinormalisasi menjadi string kosong.
        // ================================================================
        private static string NormKodeKC(string raw)
        {
            if (raw == null) return "";
            string s = raw.Trim().ToUpper();
            if (s.Length == 0)      return "";
            if (s == "FALSE")       return "";
            if (s == "TRUE")        return "";
            if (s == "0")           return "";
            if (s.StartsWith("#"))  return "";   // #N/A, #REF!, dll
            return s;
        }

        // ================================================================
        // StatistikRefKC: rekap distinct Jenis Instrumen x Kode KC
        //
        // Dikumpulkan saat membaca KC2900 setiap file tahunan, SEBELUM filter
        // ruang lingkup KC diterapkan. Tujuannya agar user bisa melihat:
        //   1. kombinasi produk dan KC apa saja yang ada di data
        //   2. produk mana yang rekeningnya gagal dipetakan ke KC
        //   3. produk yang terpetakan ke lebih dari satu KC (indikasi salah map)
        // sehingga kolom M bisa dilengkapi/dikoreksi manual.
        // ================================================================
        internal class StatistikRefKC
        {
            internal class Baris
            {
                public int    Tahun;
                public string Produk;
                public string KodeKC;      // "" = tidak terpetakan
                public int    Jumlah;
                public double Baki;
                public readonly List<string> Contoh = new List<string>();
            }

            private readonly Dictionary<string, Baris> _map =
                new Dictionary<string, Baris>(StringComparer.OrdinalIgnoreCase);

            private const int MaksContoh = 5;

            public void Catat(int tahun, string produk, string kodeKC,
                              string noRek, double baki)
            {
                string p = (produk ?? "").Trim();
                if (p.Length == 0) p = "(kosong)";
                string k = kodeKC ?? "";

                string key = tahun + "\u0001" + p + "\u0001" + k;
                Baris b;
                if (!_map.TryGetValue(key, out b))
                {
                    b = new Baris { Tahun = tahun, Produk = p, KodeKC = k };
                    _map[key] = b;
                }
                b.Jumlah++;
                b.Baki += baki;

                // Contoh nomor rekening hanya disimpan untuk yang gagal
                // dipetakan — itulah yang perlu ditelusuri user
                if (k.Length == 0 && b.Contoh.Count < MaksContoh)
                    b.Contoh.Add(noRek);
            }

            /// <summary>Semua baris, diurutkan: tahun, produk, lalu kode KC (kosong terakhir).</summary>
            public List<Baris> Semua()
            {
                var list = new List<Baris>(_map.Values);
                list.Sort(delegate (Baris a, Baris b)
                {
                    if (a.Tahun != b.Tahun) return a.Tahun.CompareTo(b.Tahun);
                    int c = string.Compare(a.Produk, b.Produk, StringComparison.OrdinalIgnoreCase);
                    if (c != 0) return c;
                    bool ax = a.KodeKC.Length == 0, bx = b.KodeKC.Length == 0;
                    if (ax != bx) return ax ? 1 : -1;   // tanpa KC ditaruh paling bawah
                    return string.Compare(a.KodeKC, b.KodeKC, StringComparison.OrdinalIgnoreCase);
                });
                return list;
            }

            public int TotalTanpaKC
            {
                get
                {
                    int n = 0;
                    foreach (var b in _map.Values) if (b.KodeKC.Length == 0) n += b.Jumlah;
                    return n;
                }
            }

            public double BakiTanpaKC
            {
                get
                {
                    double v = 0;
                    foreach (var b in _map.Values) if (b.KodeKC.Length == 0) v += b.Baki;
                    return v;
                }
            }

            /// <summary>Daftar produk yang punya minimal satu rekening tanpa Kode KC.</summary>
            public string RingkasProdukTanpaKC()
            {
                var set = new List<string>();
                foreach (var b in Semua())
                    if (b.KodeKC.Length == 0 && !set.Contains(b.Produk)) set.Add(b.Produk);
                return set.Count == 0 ? "-" : string.Join("; ", set.ToArray());
            }

            /// <summary>
            /// Kumpulan "tahun|produk" yang terpetakan ke lebih dari satu Kode KC.
            /// Biasanya indikasi salah pemetaan yang perlu dicek user.
            /// </summary>
            public HashSet<string> ProdukGanda()
            {
                var kcPerProduk = new Dictionary<string, HashSet<string>>();
                foreach (var b in _map.Values)
                {
                    if (b.KodeKC.Length == 0) continue;
                    string key = b.Tahun + "|" + b.Produk;
                    if (!kcPerProduk.ContainsKey(key))
                        kcPerProduk[key] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    kcPerProduk[key].Add(b.KodeKC);
                }
                var hasil = new HashSet<string>();
                foreach (var kv in kcPerProduk)
                    if (kv.Value.Count > 1) hasil.Add(kv.Key);
                return hasil;
            }
        }

        // ================================================================
        // TulisLogRefKC: tulis rekap Jenis Instrumen x Kode KC ke Audit Log
        //
        // Satu baris per kombinasi (tahun, jenis instrumen, kode KC), plus
        // satu baris ringkasan rekening tanpa Kode KC di akhir.
        //
        // Pemetaan kolom (mengikuti konvensi Audit Log yang sudah ada):
        //   A  Timestamp
        //   B  Jenis Proses  = "LGD ER - Referensi KC"
        //   C  Periode       = "Des yyyy"
        //   D  Tahun file
        //   H  Status        = "Terpetakan" / "TIDAK ADA KC" / "Terpetakan (ganda)"
        //   I  Jumlah rekening
        //   L  Total baki debet
        //   T  Rek tanpa Kode KC (hanya di baris ringkasan)
        //   U  Jenis Instrumen
        //   V  Kode KC
        //   W  Contoh No Rekening (maksimal 5, hanya untuk yang tanpa KC)
        // ================================================================
        private void TulisLogRefKC(Excel.Worksheet wsLog, StatistikRefKC stat)
        {
            var baris = stat.Semua();
            if (baris.Count == 0) return;

            var ganda   = stat.ProdukGanda();
            int nextRow = NextAuditLogRow(wsLog);
            double ts   = DateTime.Now.ToOADate();

            // Header kolom tambahan — ditulis ulang agar selalu berlabel
            ((Excel.Range)wsLog.Cells[4, 20]).Value2 = "Rek tanpa Kode KC";
            ((Excel.Range)wsLog.Cells[4, 21]).Value2 = "Jenis Instrumen";
            ((Excel.Range)wsLog.Cells[4, 22]).Value2 = "Kode KC";
            ((Excel.Range)wsLog.Cells[4, 23]).Value2 = "Contoh No Rekening";

            foreach (var b in baris)
            {
                bool tanpaKC = b.KodeKC.Length == 0;
                string status;
                if (tanpaKC)
                    status = "TIDAK ADA KC - lengkapi kolom M manual";
                else if (ganda.Contains(b.Tahun + "|" + b.Produk))
                    status = "Terpetakan (produk ini juga ke KC lain)";
                else
                    status = "Terpetakan";

                ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = ts;
                ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
                ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "LGD ER - Referensi KC";
                ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = "Des " + b.Tahun;
                ((Excel.Range)wsLog.Cells[nextRow, 4]).Value2 = b.Tahun;
                ((Excel.Range)wsLog.Cells[nextRow, 8]).Value2 = status;
                ((Excel.Range)wsLog.Cells[nextRow, 9]).Value2 = b.Jumlah;
                ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = b.Baki;
                ((Excel.Range)wsLog.Cells[nextRow, 12]).NumberFormat = "#,##0;(#,##0);-";
                ((Excel.Range)wsLog.Cells[nextRow, 21]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[nextRow, 21]).Value2 = b.Produk;
                ((Excel.Range)wsLog.Cells[nextRow, 22]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[nextRow, 22]).Value2 =
                    tanpaKC ? "(kosong)" : b.KodeKC;
                ((Excel.Range)wsLog.Cells[nextRow, 23]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[nextRow, 23]).Value2 =
                    tanpaKC ? string.Join(", ", b.Contoh.ToArray()) : "";

                // Baris tanpa Kode KC diberi warna agar langsung terlihat
                if (tanpaKC)
                {
                    Excel.Range rng = (Excel.Range)wsLog.Range[
                        wsLog.Cells[nextRow, 1], wsLog.Cells[nextRow, 23]];
                    rng.Interior.Color = CLR_HIGHLIGHT_YELLOW;
                }
                nextRow++;
            }

            // ---- Baris ringkasan ----
            ((Excel.Range)wsLog.Cells[nextRow, 1]).Value2 = ts;
            ((Excel.Range)wsLog.Cells[nextRow, 1]).NumberFormat = "m/d/yyyy h:mm";
            ((Excel.Range)wsLog.Cells[nextRow, 2]).Value2 = "LGD ER - Referensi KC";
            ((Excel.Range)wsLog.Cells[nextRow, 3]).Value2 = "RINGKASAN";
            ((Excel.Range)wsLog.Cells[nextRow, 8]).Value2 =
                stat.TotalTanpaKC == 0
                    ? "OK - semua rekening terpetakan"
                    : "PERLU TINDAKAN - ada rekening tanpa Kode KC";
            ((Excel.Range)wsLog.Cells[nextRow, 12]).Value2 = stat.BakiTanpaKC;
            ((Excel.Range)wsLog.Cells[nextRow, 12]).NumberFormat = "#,##0;(#,##0);-";
            ((Excel.Range)wsLog.Cells[nextRow, 20]).Value2 = stat.TotalTanpaKC;
            ((Excel.Range)wsLog.Cells[nextRow, 21]).NumberFormat = "@";
            ((Excel.Range)wsLog.Cells[nextRow, 21]).Value2 = stat.RingkasProdukTanpaKC();

            Excel.Range ringkasanRng = (Excel.Range)wsLog.Range[
                wsLog.Cells[nextRow, 1], wsLog.Cells[nextRow, 23]];
            ringkasanRng.Interior.Color = CLR_GREEN_LIGHT;
            ringkasanRng.Font.Bold      = true;
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