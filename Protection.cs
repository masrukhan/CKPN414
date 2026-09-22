using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary
{
    internal static class Protection
    {
        private static readonly string[] NamaFileDiizinkan = new[]
        {
            "Aplikasi_CKPN_414.xlsm",
            "Aplikasi_CKPN_414.xlsb",
        };

        private static readonly string[] SheetWajib = new[]
        {
            "Master",
            "Summary",
            "A. CKPN - INDV",
            "B. CKPN - KOL INDV",
            "B1.PD-Net Flow",
            "B2.PD-Migration",
            "B3.LGD-ER",
            "B4.LGD-CS MACET",
            "state PD Migration",
            "Audit Log"
        };

        // Password proteksi sheet Master yang valid.
        // Disimpan di DLL (tidak terlihat di VBA/Excel).
        private const string PasswordMaster = "HaiiWhatt??";

        // ----------------------------------------------------------------
        // Entry point utama — dipanggil sebelum setiap perhitungan
        // ----------------------------------------------------------------
        public static bool CekProteksi(Excel.Application app)
        {
            var wb = app.ActiveWorkbook;
            if (wb == null) { TampilkanError("Tidak ada workbook yang aktif."); return false; }

            // Ingat sheet yang aktif sebelum verifikasi proteksi.
            // Rangkaian Unprotect/Protect pada beberapa sheet (Master, Summary,
            // A. CKPN - INDV) dapat menggeser sheet aktif — mis. berhenti di
            // "Summary". Sheet aktif dipulihkan di blok finally agar setelah
            // pemeriksaan posisinya tetap seperti semula (umumnya "Master",
            // tempat tombol perhitungan dijalankan).
            Excel.Worksheet sheetAktifSemula = null;
            try { sheetAktifSemula = app.ActiveSheet as Excel.Worksheet; } catch { }

            try
            {
                if (!CekNamaFile(wb))          return false;
                if (!CekSheetWajib(wb))        return false;
                if (!CekProteksiSheet(wb, "Master"))                     return false;  // verifikasi password penuh
                if (!CekProteksiSheet(wb, "Summary", false))             return false;  // cukup cek terproteksi
                if (!CekProteksiSheet(wb, "A. CKPN - INDV", false))      return false;  // cukup cek terproteksi
                if (!CekProteksiSheet(wb, "B. CKPN - KOL INDV", false))  return false;  // cukup cek terproteksi
                if (!CekProteksiSheet(wb, "B1.PD-Net Flow", false))      return false;  // cukup cek terproteksi
                if (!CekProteksiSheet(wb, "B2.PD-Migration", false))     return false;  // cukup cek terproteksi
                if (!CekProteksiSheet(wb, "B3.LGD-ER", false))           return false;  // cukup cek terproteksi
                if (!CekProteksiWorkbook(wb))  return false;
                return true;
            }
            finally
            {
                // Pulihkan sheet aktif semula, apa pun hasil pemeriksaan.
                if (sheetAktifSemula != null)
                    try { sheetAktifSemula.Activate(); } catch { }
            }
        }

        // ----------------------------------------------------------------
        // Cek 1: nama file
        // ----------------------------------------------------------------
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

        // ----------------------------------------------------------------
        // Cek 2: sheet wajib ada
        // ----------------------------------------------------------------
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

        // ----------------------------------------------------------------
        // Cek 3: sheet hasil (Master, Summary, A. CKPN - INDV) HARUS terproteksi
        // dengan password yang benar. Bila salah satu tidak terproteksi atau
        // password-nya berbeda, perhitungan ditolak.
        //
        // Cara kerja:
        //   Excel tidak mengekspos password secara langsung — hanya hash-nya.
        //   Satu-satunya cara verifikasi adalah mencoba Unprotect dengan
        //   password yang diketahui. Jika berhasil (tidak throw exception),
        //   password cocok → langsung Protect ulang dengan password yang sama.
        //   Jika throw exception → password salah atau sheet tidak diproteksi
        //   dengan password ini.
        //
        //   Kasus yang ditolak:
        //     a) Sheet tidak diproteksi sama sekali (.ProtectContents = false)
        //     b) Sheet diproteksi tapi dengan password berbeda
        // ----------------------------------------------------------------
        private static bool CekProteksiSheet(Excel.Workbook wb, string namaSheet, bool verifikasiPassword = true)
        {
            // Cari sheet
            Excel.Worksheet ws = null;
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (sh.Name.Equals(namaSheet, StringComparison.OrdinalIgnoreCase))
                { ws = sh; break; }

            if (ws == null)
                return true;   // ketiadaan sheet sudah ditangani CekSheetWajib

            // Kasus a: sheet tidak diproteksi sama sekali
            if (!ws.ProtectContents)
            {
                TampilkanError(
                    "Sheet '" + namaSheet + "' tidak dalam kondisi terproteksi.\n\n" +
                    "Sheet ini harus diproteksi terlebih dahulu\n" +
                    "sebelum perhitungan dapat dijalankan.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            // Mode ringan: cukup pastikan sheet TERPROTEKSI, tanpa unprotect/protect.
            // Dipakai untuk Summary & A. CKPN - INDV agar gerbang tidak menjalankan
            // rangkaian unprotect/protect yang membuat layar sesaat tampak abu-abu.
            // Password tetap terjaga karena kedua sheet ini dibuka/dikunci oleh
            // modulnya sendiri (RefreshSummary / CKPNIndividu) dengan password
            // tertanam — mismatch password akan ketahuan di sana.
            if (!verifikasiPassword)
                return true;

            // Kasus b: sheet diproteksi — verifikasi password dengan try-unprotect
            //   Jika Unprotect berhasil  → password cocok → protect ulang → return true
            //   Jika Unprotect exception → password salah → return false
            bool passwordCocok = false;
            try
            {
                // Coba buka proteksi dengan password yang diketahui
                ws.Unprotect(PasswordMaster);

                // Sampai di sini = berhasil, password cocok
                passwordCocok = true;
            }
            catch
            {
                // Exception = password salah
                passwordCocok = false;
            }
            finally
            {
                // Selalu protect ulang jika tadi berhasil di-unprotect
                // agar sheet tidak tertinggal dalam kondisi tidak terproteksi
                if (passwordCocok)
                {
                    try
                    {
                        ws.Protect(
                            Password:             PasswordMaster,
                            DrawingObjects:       true,
                            Contents:             true,
                            Scenarios:            true,
                            UserInterfaceOnly:    false);
                    }
                    catch { /* abaikan error protect ulang */ }
                }
            }

            if (!passwordCocok)
            {
                TampilkanError(
                    "Proteksi sheet '" + namaSheet + "' tidak valid.\n\n" +
                    "Sheet terproteksi dengan password yang berbeda.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            return true;
        }

        // Password proteksi Workbook (Protect Workbook / structure).
        // Sengaja sama dengan PasswordMaster sesuai permintaan.
        private const string PasswordWorkbook = "HaiiWhatt??";
        // ----------------------------------------------------------------
        // Cek 4: Workbook harus dalam kondisi "Protect Workbook" (structure)
        // dengan password yang benar. Logika sama seperti CekProteksiSheet.
        // ----------------------------------------------------------------
        private static bool CekProteksiWorkbook(Excel.Workbook wb)
        {
            if (!wb.ProtectStructure)
            {
                TampilkanError(
                    "Workbook tidak dalam kondisi \"Protect Workbook\" (structure).\n\n" +
                    "Workbook harus diproteksi (Review > Protect Workbook)\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            bool passwordCocok = false;
            try
            {
                wb.Unprotect(PasswordWorkbook);
                passwordCocok = true;
            }
            catch
            {
                passwordCocok = false;
            }
            finally
            {
                if (passwordCocok)
                {
                    try
                    {
                        wb.Protect(
                            Password:  PasswordWorkbook,
                            Structure: true,
                            Windows:   false);
                    }
                    catch { /* abaikan error protect ulang */ }
                }
            }

            if (!passwordCocok)
            {
                TampilkanError(
                    "Proteksi Workbook tidak valid.\n\n" +
                    "Hubungi administrator aplikasi.");
                return false;
            }

            return true;
        }

        // ----------------------------------------------------------------
        // Helper: tampilkan pesan error
        // ----------------------------------------------------------------
        private static void TampilkanError(string pesan)
        {
            System.Windows.Forms.MessageBox.Show(
                pesan, "CKPN — Proteksi File",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Stop);
        }
    }
}