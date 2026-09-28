using System;
using System.IO;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Log teks sederhana untuk panel & pipeline.
    ///
    /// Lokasi utama : library\logs\panel_{user}.log  (ikut folder aplikasi,
    ///                mudah dikirim ke admin saat troubleshooting)
    /// Cadangan     : %LOCALAPPDATA%\CKPN414\panel.log
    ///                dipakai bila folder library tidak bisa ditulis
    ///                (mis. shared drive read-only untuk user tertentu).
    ///
    /// Log TIDAK PERNAH boleh menggagalkan proses: semua error ditelan.
    /// </summary>
    internal static class CatatanLog
    {
        private static readonly object _kunci = new object();

        // Batas ukuran sebelum file diputar (dipindah ke .1.log). Menjaga agar
        // file log di shared drive tidak membesar tanpa batas.
        private const long BatasUkuran = 2 * 1024 * 1024;

        public static void Tulis(string baris)
        {
            string teks = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + baris + Environment.NewLine;
            lock (_kunci)
            {
                if (TulisKe(AppPaths.FileLog, teks)) return;
                TulisKe(Path.Combine(AppPaths.FolderLokal, "panel.log"), teks);
            }
        }

        /// <summary>Lokasi file log yang sedang dipakai (untuk ditampilkan di panel).</summary>
        public static string LokasiAktif
        {
            get
            {
                try
                {
                    string folder = AppPaths.FolderLogs;
                    if (!string.IsNullOrEmpty(AppPaths.FolderLibrary) && Directory.Exists(folder))
                        return AppPaths.FileLog;
                }
                catch { }
                return Path.Combine(AppPaths.FolderLokal, "panel.log");
            }
        }

        private static bool TulisKe(string path, string teks)
        {
            try
            {
                if (string.IsNullOrEmpty(path)) return false;
                string folder = Path.GetDirectoryName(path);
                if (string.IsNullOrEmpty(folder)) return false;
                Directory.CreateDirectory(folder);

                var fi = new FileInfo(path);
                if (fi.Exists && fi.Length > BatasUkuran)
                {
                    string lama = Path.ChangeExtension(path, ".1.log");
                    try { if (File.Exists(lama)) File.Delete(lama); File.Move(path, lama); } catch { }
                }

                File.AppendAllText(path, teks);
                return true;
            }
            catch { return false; }
        }
    }
}
