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
                        BatchGrup.MintaBatal();
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
                            if (BatchGrup.Aktif) throw new InvalidOperationException("Hitung semua grup sedang berjalan.");
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
                        JalankanDiExcel(host, id, app =>
                        {
                            if (BatchGrup.Aktif) throw new InvalidOperationException("Hitung semua grup sedang berjalan.");
                            return StagingGrup.Simpan(app, alasan, catatan, refresh);
                        });
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

                    // ---- Tahap 4: riwayat antarperiode & analisis PD ----
                    case "riwayat":
                        Balas(host, id, Riwayat.Data());
                        break;

                    case "analisisSumber":
                        Balas(host, id, AnalisisPD.DaftarSumber());
                        break;

                    case "analisisData":
                        string jenisSumber = Convert.ToString(Ambil(args, "sumber") ?? "workbook");
                        if (jenisSumber == "run")
                            Balas(host, id, AnalisisPD.BacaRun(Convert.ToInt64(Ambil(args, "runId"))));
                        else if (jenisSumber == "gabungan")
                            Balas(host, id, AnalisisPD.BacaGabungan(Convert.ToString(Ambil(args, "tanggal") ?? "")));
                        else
                            JalankanDiExcel(host, id, app =>
                            {
                                if (CKPNPipeline.SedangBerjalan)
                                    throw new InvalidOperationException("Perhitungan sedang berjalan. Muat analisis setelah selesai.");
                                Excel.Workbook wbA = CariWorkbookAplikasi(app);
                                if (wbA == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
                                var d = AnalisisPD.BacaWorkbook(wbA);
                                DateTime tA;
                                d["tanggal"] = ParameterMaster.BacaTanggalLaporan(wbA, out tA) ? tA.ToString("yyyy-MM-dd") : "";
                                Excel.Worksheet mA = ParameterMaster.CariSheet(wbA, "Master");
                                d["kodeKC"] = mA == null ? "" : string.Join(",", ParameterMaster.BacaKCDicentang(mA).ToArray());
                                d["workbook"] = true;
                                return d;
                            });
                        break;

                    // ---- Tahap 3d: pengelolaan penyesuaian & database ----
                    case "simpanPenyesuaian":
                        Balas(host, id, DataPanel.SimpanPenyesuaian(Convert.ToString(Ambil(args, "modul") ?? ""),
                                                                    Ambil(args, "data") as Dictionary<string, object>,
                                                                    Convert.ToString(Ambil(args, "alasan") ?? "")));
                        break;

                    case "riwayatPenyesuaian":
                        Balas(host, id, DataPanel.RiwayatPenyesuaian(Convert.ToString(Ambil(args, "modul") ?? ""),
                                                                     Convert.ToString(Ambil(args, "kunci") ?? "")));
                        break;

                    case "infoDatabase":
                        Balas(host, id, KelolaData.InfoDatabase());
                        break;

                    case "hapusDataPeriode":
                        Balas(host, id, KelolaData.HapusDataPeriode(
                            Convert.ToString(Ambil(args, "tanggal") ?? ""),
                            Convert.ToString(Ambil(args, "konfirmasi") ?? ""),
                            Convert.ToString(Ambil(args, "alasan") ?? ""),
                            Ambil(args, "hapusPenyesuaian") is bool && (bool)Ambil(args, "hapusPenyesuaian"),
                            Ambil(args, "hapusSnapshot") is bool && (bool)Ambil(args, "hapusSnapshot")));
                        break;

                    case "kosongkanDatabase":
                        Balas(host, id, KelolaData.KosongkanDatabase(
                            Convert.ToString(Ambil(args, "konfirmasi") ?? ""),
                            Convert.ToString(Ambil(args, "alasan") ?? ""),
                            Ambil(args, "penyesuaian") is bool && (bool)Ambil(args, "penyesuaian"),
                            Ambil(args, "susunan") is bool && (bool)Ambil(args, "susunan"),
                            Ambil(args, "hapusSnapshot") is bool && (bool)Ambil(args, "hapusSnapshot")));
                        break;

                    // ---- Tahap 4e: cek cepat isi Master (periode, KC dicentang) tanpa membaca parameter lengkap ----
                    case "infoMaster":
                        JalankanDiExcel(host, id, app =>
                        {
                            Excel.Workbook wbM = CariWorkbookAplikasi(app);
                            if (wbM == null) return new Dictionary<string, object> { { "ditemukan", false } };
                            DateTime tM;
                            Excel.Worksheet mM = ParameterMaster.CariSheet(wbM, "Master");
                            return new Dictionary<string, object>
                            {
                                { "ditemukan", true },
                                { "tanggal", ParameterMaster.BacaTanggalLaporan(wbM, out tM) ? tM.ToString("yyyy-MM-dd") : "" },
                                { "kodeKC", mM == null ? "" : string.Join(",", ParameterMaster.BacaKCDicentang(mM).ToArray()) },
                                { "topN", ParameterMaster.BacaTopN(wbM) }
                            };
                        });
                        break;

                    // ---- Tahap 3b: alur periode ----
                    case "statusPeriode":
                        string tglDiminta = Convert.ToString(Ambil(args, "tanggal") ?? "");
                        JalankanDiExcel(host, id, app => StatusPeriode(app, tglDiminta));
                        break;

                    case "hitungGrup":
                        string kodeGrup = Convert.ToString(Ambil(args, "kodeKC") ?? "");
                        bool terapkanG = !(Ambil(args, "terapkanPenyesuaian") is bool) || (bool)Ambil(args, "terapkanPenyesuaian");
                        JalankanDiExcel(host, id, app =>
                        {
                            if (BatchGrup.Aktif) throw new InvalidOperationException("Hitung semua grup sedang berjalan.");
                            string tolak = CKPNPipeline.MulaiGrup(app, kodeGrup, terapkanG);
                            if (tolak != null) throw new InvalidOperationException(tolak);
                            return new Dictionary<string, object> { { "diterima", true } };
                        });
                        break;

                    // ---- Tahap 4d: hitung semua grup sekaligus ----
                    case "hitungSemuaGrup":
                        bool lewati = !(Ambil(args, "lewatiTersimpan") is bool) || (bool)Ambil(args, "lewatiTersimpan");
                        bool berhenti = !(Ambil(args, "berhentiBilaTemuan") is bool) || (bool)Ambil(args, "berhentiBilaTemuan");
                        JalankanDiExcel(host, id, app =>
                        {
                            string tolak = BatchGrup.Mulai(app, lewati, berhenti);
                            if (tolak != null) throw new InvalidOperationException(tolak);
                            return new Dictionary<string, object> { { "diterima", true } };
                        });
                        break;

                    case "refreshSummary":
                        JalankanDiExcel(host, id, app => StagingGrup.RefreshSaja(app));
                        break;

                    case "hapusVersi":
                        Periode.HapusVersi(Convert.ToInt64(Ambil(args, "runId")), Convert.ToString(Ambil(args, "alasan") ?? ""));
                        Balas(host, id, null);
                        break;

                    case "konsolidasi":
                        Balas(host, id, Periode.Konsolidasi(Convert.ToString(Ambil(args, "tanggal") ?? "")));
                        break;

                    case "tetapkanPeriode":
                        Periode.Tetapkan(Convert.ToString(Ambil(args, "tanggal") ?? ""),
                                         Convert.ToString(Ambil(args, "catatan") ?? ""));
                        Balas(host, id, null);
                        break;

                    case "bukaKunci":
                        Periode.BukaKunci(Convert.ToString(Ambil(args, "tanggal") ?? ""), Convert.ToString(Ambil(args, "alasan") ?? ""));
                        Balas(host, id, null);
                        break;

                    case "simpanSusunan":
                        var peta = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        var sMap = Ambil(args, "peta") as Dictionary<string, object>;
                        if (sMap != null) foreach (var kv in sMap) peta[kv.Key] = Convert.ToString(kv.Value);
                        Periode.SimpanSusunan(Convert.ToInt32(Ambil(args, "tahun")), peta,
                                              Convert.ToString(Ambil(args, "dasar") ?? ""),
                                              Convert.ToString(Ambil(args, "metode") ?? ""),
                                              Convert.ToString(Ambil(args, "kebijakanSaldo") ?? ""),
                                              PetaBulat(Ambil(args, "topN")));
                        Balas(host, id, null);
                        break;

                    // ---- Tahap 3c: Overview Data (OS, EAD, NPF per KC dari template Master!D14) ----
                    case "overviewData":
                        JalankanDiExcel(host, id, app =>
                        {
                            Excel.Workbook wbApp = CariWorkbookAplikasi(app);
                            if (wbApp == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
                            if (CKPNPipeline.SedangBerjalan)
                                throw new InvalidOperationException("Perhitungan sedang berjalan. Muat Overview Data setelah selesai.");
                            CKPNPipeline.MatikanGhostingSekali();
                            var kursor = app.Cursor;
                            app.Cursor = Excel.XlMousePointer.xlWait;
                            try { return new CKPNLibrary.Modules.DataOverviewBuilder(app).UntukPanel(wbApp); }
                            finally { app.Cursor = kursor; }
                        });
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

            // Top-N Master vs ketetapan tahunan (hanya peringatan di tab Manual; simpan grup menolak bila berbeda)
            DateTime tglLap;
            string namaGrupMaster = null;
            int? topNTahun = ParameterMaster.BacaTanggalLaporan(wb, out tglLap)
                ? Periode.TopNGrup(tglLap.Year, p.KCList, out namaGrupMaster) : null;

            return new Dictionary<string, object>
            {
                { "namaFile", wb.Name },
                { "topNMaster", ParameterMaster.BacaTopN(wb) }, { "topNTahun", topNTahun }, { "namaGrup", namaGrupMaster },
                { "tahun", tglLap == DateTime.MinValue ? 0 : tglLap.Year },
                { "periode",  p.BulanLaporan },
                { "kc",       p.KC },
                { "langkah",  langkah },
                { "berjalan", CKPNPipeline.SedangBerjalan }
            };
        }

        // ================================================================
        // statusPeriode — periode Master (atau periode lain yang dipilih) + status grup
        // ================================================================
        private static object StatusPeriode(Excel.Application app, string tanggalDiminta)
        {
            string tanggalMaster = "", kodeKCMaster = "";
            Excel.Workbook wb = CariWorkbookAplikasi(app);
            if (wb != null)
            {
                DateTime t;
                if (ParameterMaster.BacaTanggalLaporan(wb, out t)) tanggalMaster = t.ToString("yyyy-MM-dd");
                Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
                if (m != null) kodeKCMaster = string.Join(",", ParameterMaster.BacaKCDicentang(m).ToArray());
            }
            string tanggal = string.IsNullOrEmpty(tanggalDiminta) ? tanggalMaster : tanggalDiminta;
            if (string.IsNullOrEmpty(tanggal))
                throw new InvalidOperationException("Tanggal laporan (Master!C4) belum diisi.");
            return Periode.Status(tanggal, kodeKCMaster, tanggalMaster);
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

        // Objek JSON { namaGrup: angka } → Dictionary<string,int>; nilai tidak valid dilewati (divalidasi di Periode)
        private static Dictionary<string, int> PetaBulat(object v)
        {
            var hasil = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var d = v as Dictionary<string, object>;
            if (d == null) return hasil;
            foreach (var kv in d)
            {
                try { if (kv.Value != null) hasil[kv.Key.Trim()] = Convert.ToInt32(kv.Value, System.Globalization.CultureInfo.InvariantCulture); }
                catch { }
            }
            return hasil;
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
