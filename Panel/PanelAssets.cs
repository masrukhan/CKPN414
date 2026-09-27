using System;
using System.Collections.Generic;
using System.IO;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Menyajikan file HTML/CSS/JS panel yang di-embed di dalam DLL.
    ///
    /// Panel dibuka di alamat virtual https://ckpn.panel/ — setiap request ke
    /// alamat ini dicegat PanelHost (WebResourceRequested) dan dijawab dari
    /// resource di sini. Keuntungannya:
    ///   - tidak ada file web yang perlu diekstrak ke disk,
    ///   - isi panel ikut terkunci di dalam .xll (tidak bisa diubah user),
    ///   - update panel cukup dengan mengganti file .xll.
    /// </summary>
    internal static class PanelAssets
    {
        public const string Host    = "ckpn.panel";
        public const string BaseUrl = "https://" + Host + "/";

        private static Dictionary<string, string> _peta;   // "index.html" -> nama resource asli
        private static readonly object _kunci = new object();

        /// <summary>
        /// Buka isi file panel. Return null bila tidak ada.
        /// path: relatif terhadap folder web, mis. "index.html" atau "css/app.css".
        /// </summary>
        public static Stream Buka(string path, out string mime)
        {
            mime = TebakMime(path);
            string nama;
            if (!Peta().TryGetValue(Normalisasi(path), out nama)) return null;

            var asm = typeof(PanelAssets).Assembly;
            using (Stream s = asm.GetManifestResourceStream(nama))
            {
                if (s == null) return null;
                // Disalin ke MemoryStream: WebView2 membaca stream secara async
                // setelah handler selesai, jadi stream resource tidak boleh
                // ditutup lebih dulu.
                var ms = new MemoryStream();
                s.CopyTo(ms);
                ms.Position = 0;
                return ms;
            }
        }

        // Nama resource dari MSBuild memakai "\" untuk subfolder
        // (web/css\app.css). Dipetakan sekali ke bentuk "css/app.css".
        private static Dictionary<string, string> Peta()
        {
            lock (_kunci)
            {
                if (_peta != null) return _peta;
                var peta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (string nama in typeof(PanelAssets).Assembly.GetManifestResourceNames())
                {
                    string n = nama.Replace('\\', '/');
                    if (!n.StartsWith("web/", StringComparison.OrdinalIgnoreCase)) continue;
                    peta[n.Substring(4)] = nama;
                }
                _peta = peta;
                return _peta;
            }
        }

        private static string Normalisasi(string path)
        {
            string p = (path ?? "").Replace('\\', '/').TrimStart('/');
            int q = p.IndexOfAny(new[] { '?', '#' });
            if (q >= 0) p = p.Substring(0, q);
            return p.Length == 0 ? "index.html" : Uri.UnescapeDataString(p);
        }

        private static string TebakMime(string path)
        {
            string ext = Path.GetExtension(Normalisasi(path)).ToLowerInvariant();
            switch (ext)
            {
                case ".html": return "text/html; charset=utf-8";
                case ".css":  return "text/css; charset=utf-8";
                case ".js":   return "text/javascript; charset=utf-8";
                case ".json": return "application/json; charset=utf-8";
                case ".svg":  return "image/svg+xml";
                case ".png":  return "image/png";
                case ".woff2":return "font/woff2";
                default:      return "application/octet-stream";
            }
        }
    }
}
