using System;
using ExcelDna.Integration;
using CKPNLibrary.Modules;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary
{
    public class CKPNAddIn : IExcelAddIn
    {
        public void AutoOpen()  { }
        public void AutoClose() { }
    }

    public static class CKPNFunctions
    {
        [ExcelCommand(Name = "CKPN_HitungIndividu")]
        public static void HitungCKPNIndividu(
            string filePath, double topNDbl,
            string flagNPF, string flagKol2, string flagRestru,
            string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                int    topN   = (int)topNDbl;
                bool   npf    = flagNPF.Trim().Equals("Ya",    StringComparison.OrdinalIgnoreCase);
                bool   kol2   = flagKol2.Trim().Equals("Ya",   StringComparison.OrdinalIgnoreCase);
                bool   restru = flagRestru.Trim().Equals("Ya", StringComparison.OrdinalIgnoreCase);
                string[] sheets = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
                new CKPNIndividu(app).Hitung(filePath, topN, npf, kol2, restru, sheets);
            }
            catch (Exception ex) { TampilError("CKPN Individu", ex); }
        }

        [ExcelCommand(Name = "CKPN_IsiBucketPDNetflow")]
        public static void IsiBucketPDNetflow(
            string filePaths, string bulanLabels, string refDates,
            double topNDbl, string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                int      topN   = (int)topNDbl;
                string[] paths  = filePaths.Split(new[]{'|'}, StringSplitOptions.RemoveEmptyEntries);
                string[] labels = bulanLabels.Split(new[]{'|'}, StringSplitOptions.RemoveEmptyEntries);
                string[] dates  = refDates.Split(new[]{'|'}, StringSplitOptions.RemoveEmptyEntries);
                string[] sheets = sheetKCList.Split(new[]{','}, StringSplitOptions.RemoveEmptyEntries);
                new CKPNPDNetFlow(app).IsiBucket(paths, labels, dates, topN, sheets);
            }
            catch (Exception ex) { TampilError("PD Net Flow", ex); }
        }

        [ExcelCommand(Name = "CKPN_HitungPDMigration")]
        public static void HitungPDMigration(
            string triwulan, string pathAwal, string pathAkhir,
            string tglAwalStr, string tglAkhirStr,
            double topNDbl, string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new PDMigration(app).Hitung(
                    triwulan, pathAwal, pathAkhir,
                    tglAwalStr, tglAkhirStr,
                    (int)topNDbl, sheetKCList);
            }
            catch (Exception ex) { TampilError("PD Migration", ex); }
        }

        // ----------------------------------------------------------------
        // Run All Triwulan: hitung PD Migration Triwulan I..IV sekaligus.
        // Dipanggil VBA saat Master!C52 = "Run All Triwulan".
        //
        // Tiap parameter "…Gab" berisi 4 nilai dipisah '|' urut Triwulan I,II,III,IV:
        //   pathAwalGab | pathAkhirGab : path file awal & akhir tiap triwulan
        //   tglAwalGab  | tglAkhirGab  : label "yyyyMMdd" (opsional, hanya Audit Log)
        // topN & sheetKCList sama untuk seluruh triwulan.
        // ----------------------------------------------------------------
        [ExcelCommand(Name = "CKPN_HitungPDMigrationSemua")]
        public static void HitungPDMigrationSemua(
            string pathAwalGab, string pathAkhirGab,
            string tglAwalGab, string tglAkhirGab,
            double topNDbl, string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new PDMigration(app).HitungSemua(
                    pathAwalGab, pathAkhirGab, tglAwalGab, tglAkhirGab,
                    (int)topNDbl, sheetKCList);
            }
            catch (Exception ex) { TampilError("PD Migration - Run All Triwulan", ex); }
        }

        [ExcelCommand(Name = "CKPN_HitungLGD")]
        public static void HitungLGD(
            double currYearDbl, string filePathsStr, string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new LGDExpectedRecoveries(app).Hitung((int)currYearDbl, filePathsStr, sheetKCList);
            }
            catch (Exception ex) { TampilError("LGD Expected Recoveries", ex); }
        }

        // ----------------------------------------------------------------
        // LGD Collateral Shortfall (Macet)
        // Dipanggil VBA:
        //   Application.Run "CKPN_HitungLGDCS",
        //       fileConfigStr, haircutAgn, sheetKCList
        //
        // fileConfigStr: "label=tahun=bulan=path|label=tahun=bulan=path|..."
        //   mis. "Juni 2024=2024=6=F:\...\file.xlsx|Des 2024=2024=12=F:\...\file.xlsx|..."
        // haircutAgn: 0.2 (artinya 20%) — dari Master!C70
        // ----------------------------------------------------------------
        [ExcelCommand(Name = "CKPN_HitungLGDCS")]
        public static void HitungLGDCS(
            string fileConfigStr,
            double haircutAgn,
            string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new LGDCollateralShortfall(app).Hitung(fileConfigStr, haircutAgn, sheetKCList);
            }
            catch (Exception ex) { TampilError("LGD Collateral Shortfall", ex); }
        }

        // ----------------------------------------------------------------
        // Export sheet hasil perhitungan ke file Excel baru (values only)
        // Dipanggil VBA:
        //   Application.Run "CKPN_ExportHasil", savePath
        //
        // savePath: path lengkap file output, dipilih user via SaveFileDialog di VBA
        //           Format: "F:\Laporan\CKPN_Export_20250101_1430.xlsx"
        // ----------------------------------------------------------------
        // ----------------------------------------------------------------
        // CKPN_ExportHasil: export sheet hasil ke file Excel baru (values only)
        // Parameter:
        //   savePath    : path file output dari VBA SaveFileDialog
        //   sheetKCList : KC aktif pisah koma — untuk nama file
        // ----------------------------------------------------------------
        [ExcelCommand(Name = "CKPN_ExportHasil")]
        public static void ExportHasil(string savePath, string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new ExportToExcel(app).Export(savePath, sheetKCList);
            }
            catch (Exception ex) { TampilError("Export Hasil", ex); }
        }

        // ----------------------------------------------------------------
        // CKPN_BuatSuffixKC: bangun suffix KC untuk nama file default
        // Dipanggil VBA sebelum SaveFileDialog:
        //   Dim suffix As String
        //   suffix = Application.Run("CKPN_BuatSuffixKC", sheetKCList)
        // ----------------------------------------------------------------
        [ExcelFunction(Name = "CKPN_BuatSuffixKC", IsMacroType = true)]
        public static string BuatSuffixKC(string sheetKCList)
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                return new ExportToExcel(app).BuatSuffixKC(sheetKCList);
            }
            catch { return "Export"; }
        }

        [ExcelCommand(Name = "CKPN_RefreshDashboard")]
        public static void RefreshDashboard()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!ProtectionDashboard.CekProteksi(app)) return;
                new DashboardBuilder(app).Hitung();
            }
            catch (Exception ex) { TampilError("Dashboard CKPN", ex); }
        }

        [ExcelCommand(Name = "CKPN_DokumentasiCKPN")]
        public static void DokumentasiCKPN()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!ProtectionDashboard.CekProteksi(app)) return;
                new DokumentasiCKPNBuilder(app).Hitung();
            }
            catch (Exception ex) { TampilError("Dokumentasi CKPN", ex); }
        }

        private static void TampilError(string modul, Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                "Error " + modul + ":\n" + ex.Message,
                "CKPNLibrary",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        // ----------------------------------------------------------------
        // Refresh Summary: formula link + CKPN/PPKA/ABA dari file eksternal
        // Dipanggil VBA:
        //   Application.Run "CKPN_RefreshSummary", srcPath, sheetKCList
        //
        // srcPath     : Master!D14 (path file sumber)
        // sheetKCList : KC aktif pisah koma, mis. "KC0600,KC0700,KC1100"
        //               KC0500 (ABA) selalu ikut, tidak perlu dikirim
        // ----------------------------------------------------------------
        [ExcelCommand(Name = "CKPN_RefreshSummary")]
        public static void RefreshSummaryCmd(string filePath, string sheetKCList)
        {
            try
            {
                // Konfirmasi PALING AWAL — sebelum pengecekan proteksi. Verifikasi
                // proteksi (unprotect/protect beruntun) sesaat membuat layar tampak
                // abu-abu; dengan konfirmasi di depan, dialog muncul segera setelah
                // klik saat layar masih segar.
                var jawab = System.Windows.Forms.MessageBox.Show(
                    "Jalankan Refresh Summary sekarang?\n\n" +
                    "Proses akan membuka file sumber lalu memperbarui CKPN, PPKA, " +
                    "dan ABA di sheet Summary.",
                    "Konfirmasi Refresh Summary",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question);
                if (jawab != System.Windows.Forms.DialogResult.Yes) return;

                var app = (Excel.Application)ExcelDnaUtil.Application;
                if (!Protection.CekProteksi(app)) return;
                new RefreshSummary(app).Refresh(filePath, sheetKCList);
            }
            catch (Exception ex) { TampilError("Refresh Summary", ex); }
        }
    }
}