using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Mengekstrak DLL native yang di-embed di dalam .xll ke folder lokal user.
    ///
    /// Kenapa perlu: Excel-DNA hanya bisa mem-pack DLL managed (.NET).
    /// DLL native C++ (WebView2Loader.dll, SQLite.Interop.dll) harus ada sebagai
    /// file fisik agar bisa dimuat Windows. Maka DLL ini di-embed sebagai
    /// resource (lihat .csproj) lalu ditulis ke disk saat pertama dibutuhkan.
    ///
    /// Lokasi: %LOCALAPPDATA%\CKPN414\native\{x86|x64}\{hash}\{nama}.dll
    ///
    /// Subfolder {hash} (8 karakter SHA-256 isi DLL) sengaja dipakai:
    /// DLL yang sedang dimuat Excel terkunci dan tidak bisa ditimpa. Bila
    /// versi paket di-upgrade, isi DLL berubah → hash berubah → ditulis ke
    /// folder baru tanpa bentrok dengan file lama yang mungkin masih terkunci.
    /// </summary>
    internal static class NativeLoader
    {
        private static readonly object _kunci = new object();
        private static readonly Dictionary<string, string> _sudahSiap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static bool _sqliteDimuat;

        /// <summary>WebView2Loader.dll → folder berisi DLL (untuk SetLoaderDllFolderPath).</summary>
        public static string SiapkanWebView2Loader()
        {
            return Path.GetDirectoryName(Siapkan("WebView2Loader.dll"));
        }

        /// <summary>
        /// Muat SQLite.Interop.dll SEBELUM koneksi System.Data.SQLite pertama.
        ///
        /// System.Data.SQLite memanggil fungsi native lewat DllImport("SQLite.Interop.dll").
        /// Bila modul dengan nama itu sudah dimuat di proses (LoadLibrary di bawah),
        /// Windows memakai modul tersebut — tidak perlu file di samping EXCEL.EXE.
        /// Mekanisme "preload" bawaan System.Data.SQLite dimatikan karena ia mencari
        /// di folder assembly, yang kosong saat DLL dimuat dari dalam .xll.
        /// </summary>
        public static void MuatSQLite()
        {
            lock (_kunci)
            {
                if (_sqliteDimuat) return;
                Environment.SetEnvironmentVariable("No_PreLoadSQLite", "1");
                Environment.SetEnvironmentVariable("No_SQLiteXmlConfigFile", "1");

                string path = Siapkan("SQLite.Interop.dll");
                IntPtr h = LoadLibrary(path);
                if (h == IntPtr.Zero)
                    throw new InvalidOperationException(
                        "Gagal memuat SQLite.Interop.dll (kode Windows " + Marshal.GetLastWin32Error() + ").\n" + path);
                _sqliteDimuat = true;
            }
        }

        /// <summary>
        /// Pastikan DLL native (sesuai bitness Excel) ada di disk.
        /// Return: path lengkap file DLL.
        /// </summary>
        public static string Siapkan(string namaFile)
        {
            lock (_kunci)
            {
                string sudah;
                if (_sudahSiap.TryGetValue(namaFile, out sudah)) return sudah;

                // Bitness proses Excel, bukan Windows: Excel 32-bit di Windows 64-bit
                // tetap butuh DLL x86.
                string arch = Environment.Is64BitProcess ? "x64" : "x86";
                string resName = "native/" + arch + "/" + namaFile;

                byte[] data = BacaResource(resName);
                if (data == null)
                    throw new InvalidOperationException(
                        "Resource '" + resName + "' tidak ditemukan di dalam library. " +
                        "Periksa bagian EmbeddedResource di CKPNLibrary.csproj (versi paket NuGet).");

                string folder = Path.Combine(AppPaths.FolderNative, arch, HashPendek(data));
                string target = Path.Combine(folder, namaFile);

                if (!File.Exists(target) || new FileInfo(target).Length != data.Length)
                {
                    Directory.CreateDirectory(folder);
                    // Tulis ke file sementara lalu rename: mencegah file setengah jadi
                    // bila dua jendela Excel menyiapkan DLL bersamaan.
                    string tmp = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
                    File.WriteAllBytes(tmp, data);
                    try
                    {
                        if (File.Exists(target)) File.Delete(target);
                        File.Move(tmp, target);
                    }
                    catch (IOException)
                    {
                        // Proses lain sudah lebih dulu menulis / sedang memakai file
                        // yang sama. Selama ukurannya benar, pakai saja.
                        if (!File.Exists(target) || new FileInfo(target).Length != data.Length)
                            throw;
                    }
                    finally
                    {
                        try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                    }
                }

                _sudahSiap[namaFile] = target;
                return target;
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

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
