using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Jembatan pesan antara panel (JavaScript) dan C#.
    ///
    /// Format pesan JS → C#  (chrome.webview.postMessage):
    ///   { "id": 7, "cmd": "ping", "args": { ... } }
    ///
    /// Format balasan C# → JS (PostWebMessageAsJson):
    ///   { "type": "reply", "id": 7, "ok": true,  "data": { ... } }
    ///   { "type": "reply", "id": 7, "ok": false, "error": "pesan" }
    ///
    /// Format event dorong C# → JS (mulai Tahap 2, untuk progres perhitungan):
    ///   { "type": "event", "name": "progress", "data": { ... } }
    ///
    /// Aturan threading:
    ///   Pesan tiba di thread UI Excel, tetapi BUKAN dalam konteks makro.
    ///   Setiap perintah yang menyentuh object model Excel dijalankan lewat
    ///   ExcelAsyncUtil.QueueAsMacro — Excel mengeksekusinya saat aman
    ///   (mis. tidak sedang mode edit sel), sehingga tidak ada error COM
    ///   "application is busy".
    /// </summary>
    internal static class PanelBridge
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // ================================================================
        // Terima & distribusi perintah
        // ================================================================
        public static void Terima(PanelHost host, string jsonPesan)
        {
            object id = null;
            try
            {
                var msg = _json.Deserialize<Dictionary<string, object>>(jsonPesan);
                id = Ambil(msg, "id");
                string cmd = Convert.ToString(Ambil(msg, "cmd") ?? "");
                var args = Ambil(msg, "args") as Dictionary<string, object> ?? new Dictionary<string, object>();

                switch (cmd)
                {
                    // ---- Perintah tanpa Excel: langsung dijawab ----
                    case "ping":
                        Balas(host, id, InfoLingkungan());
                        break;

                    case "bukaFolder":
                        BukaFolder(Convert.ToString(Ambil(args, "folder") ?? "library"));
                        Balas(host, id, null);
                        break;

                    // ---- Perintah yang membaca Excel: lewat QueueAsMacro ----
                    case "infoWorkbook":
                        JalankanDiExcel(host, id, InfoWorkbook);
                        break;

                    default:
                        Gagal(host, id, "Perintah tidak dikenal: " + cmd);
                        break;
                }
            }
            catch (Exception ex)
            {
                Gagal(host, id, ex.Message);
            }
        }

        // ================================================================
        // Kirim ke panel
        // ================================================================
        public static void Balas(PanelHost host, object id, object data)
        {
            host.KirimJson(_json.Serialize(new Dictionary<string, object>
            {
                { "type", "reply" }, { "id", id }, { "ok", true }, { "data", data }
            }));
        }

        public static void Gagal(PanelHost host, object id, string pesan)
        {
            host.KirimJson(_json.Serialize(new Dictionary<string, object>
            {
                { "type", "reply" }, { "id", id }, { "ok", false }, { "error", pesan }
            }));
        }

        /// <summary>Dorong event ke SEMUA panel yang terbuka (dipakai Tahap 2).</summary>
        public static void Siarkan(string nama, object data)
        {
            string json = _json.Serialize(new Dictionary<string, object>
            {
                { "type", "event" }, { "name", nama }, { "data", data }
            });
            foreach (var h in PanelManager.SemuaHost()) h.KirimJson(json);
        }

        // Jalankan fungsi yang butuh object model Excel dalam konteks makro,
        // lalu kirim hasil/errornya ke panel yang meminta.
        private static void JalankanDiExcel(PanelHost host, object id, Func<Excel.Application, object> kerja)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var app = (Excel.Application)ExcelDnaUtil.Application;
                    Balas(host, id, kerja(app));
                }
                catch (Exception ex)
                {
                    Gagal(host, id, ex.Message);
                }
            });
        }

        // ================================================================
        // Perintah: ping — informasi lingkungan (tanpa Excel COM)
        // Dipakai panel untuk memastikan jembatan JS <-> C# berfungsi dan
        // menampilkan lokasi library/database untuk troubleshooting.
        // ================================================================
        private static object InfoLingkungan()
        {
            string versiRuntime = "";
            try { versiRuntime = Microsoft.Web.WebView2.Core.CoreWebView2Environment.GetAvailableBrowserVersionString(); }
            catch { }

            return new Dictionary<string, object>
            {
                { "versiAddin",    typeof(PanelBridge).Assembly.GetName().Version.ToString() },
                { "arsitektur",    Environment.Is64BitProcess ? "64-bit" : "32-bit" },
                { "versiExcel",    ExcelDnaUtil.ExcelVersion.ToString("0.0") },
                { "webview2",      versiRuntime },
                { "xllPath",       ExcelDnaUtil.XllPath },
                { "folderLibrary", AppPaths.FolderLibrary },
                { "folderData",    AppPaths.FolderData },
                { "databaseAda",   File.Exists(AppPaths.FileDatabase) },
                { "user",          Environment.UserName },
                { "komputer",      Environment.MachineName }
            };
        }

        // ================================================================
        // Perintah: infoWorkbook — membaca workbook aplikasi CKPN
        // ================================================================
        private static object InfoWorkbook(Excel.Application app)
        {
            // Cari workbook aplikasi berdasarkan NAMA FILE, bukan ActiveWorkbook.
            // Panel bisa diklik saat workbook lain (mis. file sumber) sedang aktif.
            Excel.Workbook wb = CariWorkbookAplikasi(app);
            if (wb == null)
            {
                return new Dictionary<string, object>
                {
                    { "ditemukan", false },
                    { "pesan", "Workbook aplikasi CKPN (" + string.Join(" / ", Protection.NamaFileDiizinkan) + ") tidak sedang terbuka." }
                };
            }

            Excel.Worksheet master = null;
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (sh.Name.Equals("Master", StringComparison.OrdinalIgnoreCase)) { master = sh; break; }

            var info = new Dictionary<string, object>
            {
                { "ditemukan", true },
                { "namaFile",  wb.Name },
                { "folder",    wb.Path }
            };
            if (master == null) { info["pesan"] = "Sheet Master tidak ditemukan."; return info; }

            info["bulanLaporan"]   = FormatTanggal(((Excel.Range)master.Range["C4"]).Value2);
            info["fileIndividu"]   = Convert.ToString(((Excel.Range)master.Range["D14"]).Value2 ?? "");
            info["topN"]           = Convert.ToString(((Excel.Range)master.Range["C10"]).Value2 ?? "");
            info["kcDicentang"]    = BacaKCDicentang(master);
            return info;
        }

        /// <summary>
        /// Workbook aplikasi CKPN yang sedang terbuka (dicocokkan dengan daftar
        /// nama file resmi di Protection). Null bila tidak ada.
        /// </summary>
        internal static Excel.Workbook CariWorkbookAplikasi(Excel.Application app)
        {
            foreach (Excel.Workbook wb in app.Workbooks)
                foreach (var nama in Protection.NamaFileDiizinkan)
                    if (wb.Name.Equals(nama, StringComparison.OrdinalIgnoreCase)) return wb;
            return null;
        }

        // Checkbox KC0600..KC1100 di Master adalah Form Control lama
        // (Worksheet.CheckBoxes). Tidak ada tipe interop yang kuat untuknya,
        // jadi properti Value dibaca lewat late binding (reflection COM).
        // Nilai xlOn = 1.
        private static readonly string[] DaftarKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        private static List<string> BacaKCDicentang(Excel.Worksheet master)
        {
            var hasil = new List<string>();
            foreach (string kc in DaftarKC)
            {
                try
                {
                    object chk = master.CheckBoxes(kc);
                    object val = chk.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, chk, null);
                    if (Convert.ToInt32(val) == 1) hasil.Add(kc);
                }
                catch { /* checkbox tidak ada → dianggap tidak dicentang */ }
            }
            return hasil;
        }

        private static string FormatTanggal(object v)
        {
            if (v is double) return DateTime.FromOADate((double)v).ToString("dd MMMM yyyy", new System.Globalization.CultureInfo("id-ID"));
            return Convert.ToString(v ?? "");
        }

        // ================================================================
        // Perintah: bukaFolder — buka Windows Explorer
        // ================================================================
        private static void BukaFolder(string jenis)
        {
            string path;
            switch (jenis)
            {
                case "data": path = AppPaths.FolderData;   break;
                case "logs": path = AppPaths.FolderLogs;   break;
                case "lokal": path = AppPaths.FolderLokal; break;
                default:     path = AppPaths.FolderLibrary; break;
            }
            // Hanya folder yang dikenal add-in yang boleh dibuka — panel tidak
            // bisa meminta path sembarang.
            if (Directory.Exists(path)) Process.Start("explorer.exe", "\"" + path + "\"");
        }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }
    }
}
