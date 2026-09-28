using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Data;
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
    /// Format event dorong C# → JS (progres perhitungan):
    ///   { "type": "event", "name": "langkahMulai", "data": { ... } }
    ///   Nama event: runMulai, langkahMulai, progresDetail, pesanModul,
    ///               langkahSelesai, runSelesai
    ///
    /// Aturan threading:
    ///   Pesan tiba di thread UI Excel, tetapi BUKAN dalam konteks makro.
    ///   Setiap perintah yang menyentuh object model Excel dijalankan lewat
    ///   ExcelAsyncUtil.QueueAsMacro — Excel mengeksekusinya saat aman
    ///   (mis. tidak sedang mode edit sel).
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
                    // ---- Tanpa Excel: langsung dijawab ----
                    case "ping":
                        Balas(host, id, InfoLingkungan());
                        break;

                    case "bukaFolder":
                        BukaFolder(Convert.ToString(Ambil(args, "folder") ?? "library"));
                        Balas(host, id, null);
                        break;

                    case "statusRun":
                        Balas(host, id, CKPNPipeline.Status());
                        break;

                    case "batal":
                        CKPNPipeline.MintaBatal();
                        Balas(host, id, null);
                        break;

                    // ---- Membaca / menulis Excel: lewat QueueAsMacro ----
                    case "infoWorkbook":
                        JalankanDiExcel(host, id, InfoWorkbook);
                        break;

                    case "siapkanRun":
                        JalankanDiExcel(host, id, SiapkanRun);
                        break;

                    case "jalankan":
                        var ids = DaftarString(Ambil(args, "langkah"));
                        bool terapkan = !(Ambil(args, "terapkanPenyesuaian") is bool) || (bool)Ambil(args, "terapkanPenyesuaian");
                        JalankanDiExcel(host, id, app =>
                        {
                            string tolak = CKPNPipeline.Mulai(app, ids, terapkan);
                            if (tolak != null) throw new InvalidOperationException(tolak);
                            return new Dictionary<string, object> { { "diterima", true } };
                        });
                        break;

                    // ---- Tahap 3a: review, simpan grup, staging, penyesuaian ----
                    case "reviewInfo":
                        Balas(host, id, Penyesuaian.InfoReview());
                        break;

                    case "lompatKe":
                        string sheet = Convert.ToString(Ambil(args, "sheet") ?? "");
                        string sel   = Convert.ToString(Ambil(args, "sel") ?? "A1");
                        JalankanDiExcel(host, id, app => { LompatKe(app, sheet, sel); return null; });
                        break;

                    case "periksaSimpan":
                        JalankanDiExcel(host, id, app => StagingGrup.Periksa(app));
                        break;

                    case "simpanGrup":
                        var alasan = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        var aMap = Ambil(args, "alasan") as Dictionary<string, object>;
                        if (aMap != null) foreach (var kv in aMap) alasan[kv.Key] = Convert.ToString(kv.Value);
                        string catatan = Convert.ToString(Ambil(args, "catatan") ?? "");
                        bool refresh = !(Ambil(args, "refreshSummary") is bool) || (bool)Ambil(args, "refreshSummary");
                        JalankanDiExcel(host, id, app => StagingGrup.Simpan(app, alasan, catatan, refresh));
                        break;

                    case "daftarStaging":
                        Balas(host, id, DataPanel.DaftarStaging(Convert.ToString(Ambil(args, "tanggal") ?? "")));
                        break;

                    case "bukaSnapshot":
                        DataPanel.BukaSnapshot(Convert.ToInt64(Ambil(args, "runId")));
                        Balas(host, id, null);
                        break;

                    case "daftarPenyesuaian":
                        Balas(host, id, DataPanel.DaftarPenyesuaian());
                        break;

                    case "hapusPenyesuaian":
                        DataPanel.HapusPenyesuaian(Convert.ToString(Ambil(args, "modul") ?? ""),
                                                   Convert.ToString(Ambil(args, "kunci") ?? ""),
                                                   Convert.ToString(Ambil(args, "alasan") ?? ""));
                        Balas(host, id, null);
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

        /// <summary>Dorong event ke SEMUA panel yang terbuka.</summary>
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
        // ping — informasi lingkungan (tanpa Excel COM)
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
                { "fileLog",       CatatanLog.LokasiAktif },
                { "databaseAda",   File.Exists(AppPaths.FileDatabase) },
                { "fileDatabase",  AppPaths.FileDatabase },
                { "user",          Environment.UserName },
                { "komputer",      Environment.MachineName }
            };
        }

        // ================================================================
        // infoWorkbook — ringkasan workbook aplikasi CKPN
        // ================================================================
        private static object InfoWorkbook(Excel.Application app)
        {
            Excel.Workbook wb = CariWorkbookAplikasi(app);
            if (wb == null)
            {
                return new Dictionary<string, object>
                {
                    { "ditemukan", false },
                    { "pesan", "Workbook aplikasi CKPN (" + string.Join(" / ", Protection.NamaFileDiizinkan) + ") tidak sedang terbuka." }
                };
            }

            var info = new Dictionary<string, object>
            {
                { "ditemukan", true }, { "namaFile", wb.Name }, { "folder", wb.Path }
            };
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            if (master == null) { info["pesan"] = "Sheet Master tidak ditemukan."; return info; }

            info["kcDicentang"] = ParameterMaster.BacaKCDicentang(master);
            return info;
        }

        // ================================================================
        // siapkanRun — baca Master & validasi setiap langkah
        // ================================================================
        private static object SiapkanRun(Excel.Application app)
        {
            Excel.Workbook wb = CariWorkbookAplikasi(app);
            if (wb == null)
                throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");

            ParameterCKPN p = ParameterMaster.Baca(wb);
            var langkah = new List<object>();
            foreach (var l in CKPNPipeline.SemuaLangkah)
            {
                List<string> err;
                p.Error.TryGetValue(l.Id, out err);
                string ringkas;
                p.Ringkas.TryGetValue(l.Id, out ringkas);
                langkah.Add(new Dictionary<string, object>
                {
                    { "id", l.Id }, { "nama", l.Nama }, { "sheet", l.SheetDitulis },
                    { "siap", err == null }, { "error", err ?? new List<string>() },
                    { "ringkas", ringkas ?? "" }
                });
            }

            return new Dictionary<string, object>
            {
                { "namaFile", wb.Name },
                { "periode",  p.BulanLaporan },
                { "kc",       p.KC },
                { "langkah",  langkah },
                { "berjalan", CKPNPipeline.SedangBerjalan }
            };
        }

        // ================================================================
        // lompatKe — aktifkan sheet & pilih sel (dari daftar review panel)
        // ================================================================
        private static readonly string[] SheetBolehDituju = { "A. CKPN - INDV", "B4.LGD-CS MACET", "Summary", "Master" };

        private static void LompatKe(Excel.Application app, string namaSheet, string sel)
        {
            if (Array.IndexOf(SheetBolehDituju, namaSheet) < 0)
                throw new InvalidOperationException("Sheet tidak diizinkan: " + namaSheet);
            Excel.Workbook wb = CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, namaSheet);
            if (ws == null) throw new InvalidOperationException("Sheet '" + namaSheet + "' tidak ditemukan.");
            wb.Activate();
            ws.Activate();
            ((Excel.Range)ws.Range[sel]).Select();
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

        // ================================================================
        // bukaFolder — buka Windows Explorer (hanya folder yang dikenal add-in)
        // ================================================================
        private static void BukaFolder(string jenis)
        {
            string path;
            switch (jenis)
            {
                case "data":  path = AppPaths.FolderData;  break;
                case "logs":  path = Path.GetDirectoryName(CatatanLog.LokasiAktif); break;
                case "lokal": path = AppPaths.FolderLokal; break;
                default:      path = AppPaths.FolderLibrary; break;
            }
            if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                Process.Start("explorer.exe", "\"" + path + "\"");
        }

        // ================================================================
        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        // Array JSON dideserialisasi JavaScriptSerializer menjadi ArrayList / object[]
        private static List<string> DaftarString(object v)
        {
            var hasil = new List<string>();
            var e = v as IEnumerable;
            if (e == null || v is string) return hasil;
            foreach (var x in e) if (x != null) hasil.Add(Convert.ToString(x));
            return hasil;
        }
    }
}
