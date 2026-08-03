using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary
{
    /// <summary>
    /// Proteksi khusus file Dashboard CKPN (Dashboard_CKPN414.xlsm) —
    /// terpisah dari Protection.cs (untuk CKPN414_Himbarsi_Ed2026.xlsm)
    /// karena whitelist nama file, sheet wajib, dan sheet yang di-lock berbeda.
    /// </summary>
    internal static class ProtectionDashboard
    {
        private static readonly string[] NamaFileDiizinkan = new[]
        {
            "Dashboard_CKPN414.xlsm",
            "Dashboard_CKPN414.xlsb",
        };

        private static readonly string[] SheetWajib = new[]
        {
            "Sumber Data",
            "Dashboard"
        };

        // Password proteksi — sengaja sama dengan Protection.cs (PasswordMaster/PasswordWorkbook)
        private const string PasswordDashboard = "HaiiWhatt??";

        public static bool CekProteksi(Excel.Application app)
        {
            var wb = app.ActiveWorkbook;
            if (wb == null) { TampilkanError("Tidak ada workbook yang aktif."); return false; }
            if (!CekNamaFile(wb))            return false;
            if (!CekSheetWajib(wb))          return false;
            if (!CekProteksiSumberData(wb))  return false;
            if (!CekProteksiWorkbook(wb))    return false;
            return true;
        }

        private static bool CekNamaFile(Excel.Workbook wb)
        {
            string nama = System.IO.Path.GetFileName(wb.FullName);
            foreach (var s in NamaFileDiizinkan)
                if (nama.Equals(s, StringComparison.OrdinalIgnoreCase)) return true;

            TampilkanError(
                "Nama file tidak dikenali:\n  \"" + nama + "\"\n\n" +
                "Nama yang diizinkan:\n  \"" +
                string.Join("\"\n  \"", NamaFileDiizinkan) + "\"\n\n" +
                "Kembalikan nama file ke nama aslinya.");
            return false;
        }

        private static bool CekSheetWajib(Excel.Workbook wb)
        {
            var tidakAda = new List<string>();
            foreach (var nama in SheetWajib)
            {
                bool ada = false;
                foreach (Excel.Worksheet sh in wb.Worksheets)
                    if (sh.Name.Equals(nama, StringComparison.OrdinalIgnoreCase))
                    { ada = true; break; }
                if (!ada) tidakAda.Add(nama);
            }
            if (tidakAda.Count == 0) return true;

            TampilkanError(
                "Sheet berikut tidak ditemukan:\n  \"" +
                string.Join("\"\n  \"", tidakAda) + "\"\n\n" +
                "Kembalikan nama sheet ke nama aslinya.");
            return false;
        }

        // Sheet "Sumber Data" harus terproteksi dengan password yang benar.
        // Logika identik dengan CekProteksiMaster di Protection.cs.
        private static bool CekProteksiSumberData(Excel.Workbook wb)
        {
            Excel.Worksheet wsSumber = null;
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (sh.Name.Equals("Sumber Data", StringComparison.OrdinalIgnoreCase))
                { wsSumber = sh; break; }

            if (wsSumber == null)
                return true; // sudah ditangani CekSheetWajib

            if (!wsSumber.ProtectContents)
            {
                TampilkanError(
                    "Sheet 'Sumber Data' tidak dalam kondisi terproteksi.\n\n" +
                    "Sheet 'Sumber Data' harus diproteksi dengan password yang benar\n" +
                    "sebelum Dashboard dapat di-refresh.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            bool passwordCocok = false;
            try
            {
                wsSumber.Unprotect(PasswordDashboard);
                passwordCocok = true;
            }
            catch { passwordCocok = false; }
            finally
            {
                if (passwordCocok)
                {
                    try
                    {
                        wsSumber.Protect(
                            Password:          PasswordDashboard,
                            DrawingObjects:    true,
                            Contents:          true,
                            Scenarios:         true,
                            UserInterfaceOnly: false);
                    }
                    catch { /* abaikan error protect ulang */ }
                }
            }

            if (!passwordCocok)
            {
                TampilkanError(
                    "Proteksi sheet 'Sumber Data' tidak valid.\n\n" +
                    "Sheet 'Sumber Data' terdeteksi diproteksi dengan password\n" +
                    "yang berbeda dari yang dikenali oleh aplikasi ini.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            return true;
        }

        // Workbook harus dalam kondisi "Protect Workbook" (structure).
        // Logika identik dengan CekProteksiWorkbook di Protection.cs.
        private static bool CekProteksiWorkbook(Excel.Workbook wb)
        {
            if (!wb.ProtectStructure)
            {
                TampilkanError(
                    "Workbook tidak dalam kondisi \"Protect Workbook\" (structure).\n\n" +
                    "Workbook harus diproteksi (Review > Protect Workbook) dengan\n" +
                    "password yang benar sebelum Dashboard dapat di-refresh.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            bool passwordCocok = false;
            try
            {
                wb.Unprotect(PasswordDashboard);
                passwordCocok = true;
            }
            catch { passwordCocok = false; }
            finally
            {
                if (passwordCocok)
                {
                    try
                    {
                        wb.Protect(
                            Password:  PasswordDashboard,
                            Structure: true,
                            Windows:   false);
                    }
                    catch { /* abaikan error protect ulang */ }
                }
            }

            if (!passwordCocok)
            {
                TampilkanError("Proteksi Workbook tidak valid.\n\nHubungi administrator aplikasi.");
                return false;
            }

            return true;
        }

        private static void TampilkanError(string pesan)
        {
            System.Windows.Forms.MessageBox.Show(
                pesan, "Dashboard CKPN — Proteksi File",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Stop);
        }
    }
}