using System;
using System.IO;
using ExcelDna.Integration;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Pusat semua lokasi folder yang dipakai add-in.
    ///
    /// Dua kelompok lokasi dengan alasan berbeda:
    ///
    /// 1. Folder LIBRARY (di samping file .xll, mis. CKPN\library\):
    ///    data\ (ckpn.db), config\, backup\, logs\.
    ///    Dihitung dari ExcelDnaUtil.XllPath — bukan path absolut — sehingga
    ///    seluruh folder CKPN bisa dipindah ke drive/komputer lain tanpa setting.
    ///    Folder ini bisa berada di shared drive dan dipakai bersama satu tim.
    ///
    /// 2. Folder LOKAL per user (%LOCALAPPDATA%\CKPN414\):
    ///    DLL native WebView2Loader dan cache/profil WebView2.
    ///    Sengaja TIDAK di folder library karena:
    ///      - memuat DLL native dari network share sering diblokir antivirus/IT,
    ///      - profil WebView2 harus bisa ditulis dan tidak boleh dipakai
    ///        bersamaan oleh beberapa user (akan terkunci).
    /// </summary>
    internal static class AppPaths
    {
        public const string NamaFolderLokal = "CKPN414";

        // ---------------- Folder library (shared) ----------------

        /// <summary>Folder tempat file .xll berada (mis. F:\...\CKPN\library).</summary>
        public static string FolderLibrary
        {
            get
            {
                string xll = ExcelDnaUtil.XllPath;
                return string.IsNullOrEmpty(xll) ? "" : Path.GetDirectoryName(xll);
            }
        }

        /// <summary>Folder aplikasi CKPN (induk dari folder library).</summary>
        public static string FolderAplikasi
        {
            get
            {
                string lib = FolderLibrary;
                if (string.IsNullOrEmpty(lib)) return "";
                var parent = Directory.GetParent(lib);
                return parent == null ? lib : parent.FullName;
            }
        }

        public static string FolderData   { get { return Path.Combine(FolderLibrary, "data");   } }
        public static string FolderConfig { get { return Path.Combine(FolderLibrary, "config"); } }
        public static string FolderBackup { get { return Path.Combine(FolderLibrary, "backup"); } }
        public static string FolderLogs   { get { return Path.Combine(FolderLibrary, "logs");   } }

        /// <summary>Lokasi database SQLite (dipakai mulai Tahap 3).</summary>
        public static string FileDatabase { get { return Path.Combine(FolderData, "ckpn.db"); } }

        // ---------------- Folder lokal per user ----------------

        /// <summary>%LOCALAPPDATA%\CKPN414</summary>
        public static string FolderLokal
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    NamaFolderLokal);
            }
        }

        /// <summary>Profil/cache WebView2. Default WebView2 menaruhnya di samping
        /// EXCEL.EXE (Program Files) yang tidak bisa ditulis → panel blank.</summary>
        public static string FolderWebView2Data { get { return Path.Combine(FolderLokal, "WebView2"); } }

        /// <summary>Tempat ekstraksi DLL native (WebView2Loader, nanti SQLite).</summary>
        public static string FolderNative { get { return Path.Combine(FolderLokal, "native"); } }

        // ---------------- Inisialisasi ----------------

        /// <summary>
        /// Membuat folder yang dibutuhkan bila belum ada. Dipanggil di AutoOpen.
        /// Kegagalan membuat folder library (mis. shared drive read-only untuk
        /// user tertentu) TIDAK menggagalkan add-in: perhitungan tetap jalan,
        /// hanya fitur database yang nanti menolak menulis.
        /// </summary>
        public static void SiapkanFolder()
        {
            BuatAman(FolderLokal);
            BuatAman(FolderWebView2Data);
            BuatAman(FolderNative);

            if (string.IsNullOrEmpty(FolderLibrary)) return;
            BuatAman(FolderData);
            BuatAman(FolderConfig);
            BuatAman(FolderBackup);
            BuatAman(FolderLogs);
        }

        private static void BuatAman(string folder)
        {
            try { if (!Directory.Exists(folder)) Directory.CreateDirectory(folder); }
            catch { /* abaikan: hak akses / drive jaringan terputus */ }
        }
    }
}
