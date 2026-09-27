using System;
using System.IO;
using System.Security.Cryptography;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Mengekstrak DLL native yang di-embed di dalam .xll ke folder lokal user.
    ///
    /// Kenapa perlu: Excel-DNA hanya bisa mem-pack DLL managed (.NET).
    /// WebView2Loader.dll adalah DLL native C++ yang harus ada sebagai file
    /// fisik agar bisa di-LoadLibrary oleh Windows. Maka DLL ini di-embed
    /// sebagai resource (lihat .csproj) lalu ditulis ke disk saat dibutuhkan.
    ///
    /// Lokasi: %LOCALAPPDATA%\CKPN414\native\{x86|x64}\{hash}\WebView2Loader.dll
    ///
    /// Subfolder {hash} (8 karakter SHA-256 isi DLL) sengaja dipakai:
    /// DLL yang sedang dimuat Excel terkunci dan tidak bisa ditimpa. Bila
    /// versi WebView2 di-upgrade, isi DLL berubah → hash berubah → ditulis ke
    /// folder baru tanpa bentrok dengan file lama yang mungkin masih terkunci.
    /// </summary>
    internal static class NativeLoader
    {
        private static readonly object _kunci = new object();
        private static string _folderWebView2Loader;

        /// <summary>
        /// Pastikan WebView2Loader.dll sesuai arsitektur Excel tersedia di disk.
        /// Return: folder berisi DLL tersebut (untuk SetLoaderDllFolderPath).
        /// </summary>
        public static string SiapkanWebView2Loader()
        {
            lock (_kunci)
            {
                if (_folderWebView2Loader != null) return _folderWebView2Loader;

                // Bitness proses Excel, bukan Windows: Excel 32-bit di Windows 64-bit
                // tetap butuh loader x86.
                string arch = Environment.Is64BitProcess ? "x64" : "x86";
                string resName = "native/" + arch + "/WebView2Loader.dll";

                byte[] data = BacaResource(resName);
                if (data == null)
                    throw new InvalidOperationException(
                        "Resource '" + resName + "' tidak ditemukan di dalam library. " +
                        "Periksa bagian EmbeddedResource WebView2Loader di CKPNLibrary.csproj.");

                string folder = Path.Combine(AppPaths.FolderNative, arch, HashPendek(data));
                string target = Path.Combine(folder, "WebView2Loader.dll");

                if (!File.Exists(target) || new FileInfo(target).Length != data.Length)
                {
                    Directory.CreateDirectory(folder);
                    // Tulis ke file sementara lalu rename: mencegah file setengah jadi
                    // bila dua jendela Excel membuka panel bersamaan.
                    string tmp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    File.WriteAllBytes(tmp, data);
                    try
                    {
                        // File lama dengan ukuran berbeda = rusak/terpotong → ganti.
                        if (File.Exists(target)) File.Delete(target);
                        File.Move(tmp, target);
                    }
                    catch (IOException)
                    {
                        // Jendela Excel lain sudah lebih dulu menulis file yang sama
                        // (atau sedang memakainya). Selama ukurannya benar, pakai saja.
                        if (!File.Exists(target) || new FileInfo(target).Length != data.Length)
                            throw;
                    }
                    finally
                    {
                        try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                    }
                }

                _folderWebView2Loader = folder;
                return folder;
            }
        }

        private static byte[] BacaResource(string nama)
        {
            var asm = typeof(NativeLoader).Assembly;
            using (Stream s = asm.GetManifestResourceStream(nama))
            {
                if (s == null) return null;
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    return ms.ToArray();
                }
            }
        }

        private static string HashPendek(byte[] data)
        {
            using (var sha = SHA256.Create())
            {
                byte[] h = sha.ComputeHash(data);
                return BitConverter.ToString(h, 0, 4).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
