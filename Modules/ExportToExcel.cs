using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Export sheet-sheet hasil perhitungan CKPN ke file Excel baru.
    ///
    /// Yang di-export (values only, tanpa formula):
    ///   - Summary
    ///   - A. CKPN - INDV
    ///   - B. CKPN - KOL INDV
    ///   - B1.PD-Net Flow
    ///   - B2.PD-Migration
    ///   - B3.LGD-ER
    ///   - B4.LGD-CS MACET
    ///
    /// Format output:
    ///   - Semua formula diganti dengan nilai (PasteSpecial Values)
    ///   - Format angka, warna, border, lebar kolom dipertahankan
    ///   - Nama file: [NamaFileAsli]_Export_[yyyyMMdd_HHmm].xlsx
    ///   - Lokasi: dipilih user via SaveFileDialog
    /// </summary>
    internal class ExportToExcel
    {
        private readonly Excel.Application _app;

        // Sheet yang akan di-export (urutan sesuai tampilan di file asli)
        private static readonly string[] SheetYangDiExport =
        {
            "Master",
            "Summary",
            "A. CKPN - INDV",
            "B. CKPN - KOL INDV",
            "B1.PD-Net Flow",
            "B2.PD-Migration",
            "B3.LGD-ER",
            "B4.LGD-CS MACET"
        };

        public ExportToExcel(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        // ================================================================
        // Entry point — dipanggil dari CKPNFunctions via ExcelCommand
        //
        // Pendekatan: TIDAK menggunakan wsSrc.Copy() karena metode Copy
        // selalu membawa VBA project ikut ke workbook baru sehingga
        // muncul dialog "macro-free workbook" saat SaveAs ke .xlsx.
        //
        // Pendekatan baru (bersih tanpa VBA ikut):
        //   1. Buat workbook baru kosong
        //   2. Per sheet: tambah sheet baru di wbDst, copy UsedRange
        //      dari wsSrc, PasteSpecial(Values+Formats) ke wsDst
        //   3. SaveAs sebagai .xlsx — tidak ada VBA, tidak ada dialog
        // ================================================================
        public void Export(string savePath, string sheetKCList)
        {
            if (string.IsNullOrEmpty(savePath))
                return;

            var wbSrc = _app.ActiveWorkbook;
            if (wbSrc == null)
                throw new InvalidOperationException("Tidak ada workbook yang aktif.");

            // Kumpulkan sheet yang ada (skip yang tidak ditemukan)
            var sheetAda      = new List<string>();
            var sheetTidakAda = new List<string>();

            foreach (var nama in SheetYangDiExport)
            {
                bool ada = false;
                foreach (Excel.Worksheet sh in wbSrc.Worksheets)
                    if (sh.Name.Equals(nama, StringComparison.OrdinalIgnoreCase))
                    { ada = true; break; }

                if (ada) sheetAda.Add(nama);
                else     sheetTidakAda.Add(nama);
            }

            if (sheetAda.Count == 0)
                throw new InvalidOperationException(
                    "Tidak ada sheet yang bisa di-export.\n" +
                    "Sheet yang dicari:\n  " +
                    string.Join("\n  ", SheetYangDiExport));

            // Buat workbook baru — kosong, tanpa VBA sama sekali
            Excel.Workbook wbDst = _app.Workbooks.Add();
            _app.DisplayAlerts   = false;

            try
            {
                // Hapus semua sheet default (Sheet1, Sheet2, dst)
                // Workbook baru minimal harus punya 1 sheet, jadi hapus
                // setelah kita menambahkan sheet pertama kita sendiri
                var sheetDefaultList = new List<string>();
                foreach (Excel.Worksheet sh in wbDst.Worksheets)
                    sheetDefaultList.Add(sh.Name);

                bool sheetDefaultDihapus = false;

                foreach (var namaSheet in sheetAda)
                {
                    // Cari sheet sumber
                    Excel.Worksheet wsSrc = null;
                    foreach (Excel.Worksheet sh in wbSrc.Worksheets)
                        if (sh.Name.Equals(namaSheet, StringComparison.OrdinalIgnoreCase))
                        { wsSrc = sh; break; }
                    if (wsSrc == null) continue;

                    // Tambah sheet baru di wbDst — BUKAN copy, jadi VBA tidak ikut
                    Excel.Worksheet wsDst = (Excel.Worksheet)wbDst.Worksheets.Add(
                        After: wbDst.Worksheets[wbDst.Worksheets.Count]);
                    wsDst.Name = namaSheet;

                    // Hapus sheet default setelah sheet pertama kita ada
                    if (!sheetDefaultDihapus)
                    {
                        foreach (var defName in sheetDefaultList)
                        {
                            try
                            {
                                Excel.Worksheet defSh = null;
                                foreach (Excel.Worksheet s in wbDst.Worksheets)
                                    if (s.Name.Equals(defName, StringComparison.OrdinalIgnoreCase))
                                    { defSh = s; break; }
                                if (defSh != null) defSh.Delete();
                            }
                            catch { }
                        }
                        sheetDefaultDihapus = true;
                    }

                    // Transfer konten: Copy UsedRange dari sumber
                    Excel.Range srcUsed = wsSrc.UsedRange;
                    if (srcUsed == null) continue;

                    srcUsed.Copy();

                    // PasteSpecial Values — isi nilai tanpa formula
                    Excel.Range dstCell = (Excel.Range)wsDst.Cells[
                        srcUsed.Row, srcUsed.Column];
                    dstCell.PasteSpecial(
                        Paste:      Excel.XlPasteType.xlPasteValues,
                        Operation:  Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                        SkipBlanks: false,
                        Transpose:  false);

                    // PasteSpecial Formats — salin format (warna, border, lebar kolom)
                    dstCell.PasteSpecial(
                        Paste:      Excel.XlPasteType.xlPasteFormats,
                        Operation:  Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                        SkipBlanks: false,
                        Transpose:  false);

                    // PasteSpecial ColumnWidths — salin lebar kolom
                    dstCell.PasteSpecial(
                        Paste:      Excel.XlPasteType.xlPasteColumnWidths,
                        Operation:  Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                        SkipBlanks: false,
                        Transpose:  false);

                    // Hilangkan garis putus-putus mode copy
                    _app.CutCopyMode = (Excel.XlCutCopyMode)0;
                }

                // Aktifkan sheet pertama
                if (wbDst.Worksheets.Count > 0)
                    ((Excel.Worksheet)wbDst.Worksheets[1]).Activate();

                // SaveAs .xlsx — workbook ini murni tanpa VBA,
                // sehingga tidak ada dialog "macro-free workbook"
                wbDst.SaveAs(
                    savePath,
                    FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook,
                    ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges);

                wbDst.Close(false);
                wbDst = null;
                _app.DisplayAlerts = true;

                // Pesan selesai
                string pesanOK = "Export berhasil!\n\n" +
                                 "File tersimpan di:\n" + savePath + "\n\n" +
                                 "Sheet yang di-export (" + sheetAda.Count + "):\n  " +
                                 string.Join("\n  ", sheetAda);

                if (sheetTidakAda.Count > 0)
                    pesanOK += "\n\nSheet tidak ditemukan (dilewati):\n  " +
                               string.Join("\n  ", sheetTidakAda);

                System.Windows.Forms.MessageBox.Show(
                    pesanOK,
                    "Export CKPN — Selesai",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _app.CutCopyMode  = (Excel.XlCutCopyMode)0;
                _app.DisplayAlerts = true;

                if (wbDst != null)
                {
                    try { wbDst.Close(false); } catch { }
                }

                try
                {
                    if (System.IO.File.Exists(savePath))
                        System.IO.File.Delete(savePath);
                }
                catch { }

                throw new InvalidOperationException("Export gagal: " + ex.Message, ex);
            }
        }


        // ================================================================
        // BuatSuffixKC: bangun suffix nama file dari sheetKCList
        //
        // Semua KC aktif (6 KC)    → "all_segmen"
        // Sebagian KC              → "KC0600KC0700" (gabungan tanpa pemisah)
        // Hanya 1 KC               → "KC0600"
        // Kosong                   → "Export"
        //
        // Dipanggil dari VBA via Application.Run untuk mendapatkan
        // nama file default sebelum SaveFileDialog ditampilkan.
        // ================================================================
        public string BuatSuffixKC(string sheetKCList)
        {
            if (string.IsNullOrEmpty(sheetKCList)) return "Export";

            var parts = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return "Export";

            // Jika semua 6 KC aktif → gunakan label ringkas "all_segmen"
            if (parts.Length >= 6) return "all_segmen";

            // Sebagian KC → gabungkan nama KC tanpa pemisah
            var sb = new System.Text.StringBuilder();
            foreach (var p in parts)
                sb.Append(p.Trim().Replace(" ", ""));

            return sb.ToString();
        }
    }
}