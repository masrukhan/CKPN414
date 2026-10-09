using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using CKPNLibrary.Helpers;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Isi task pane "Panel CKPN": sebuah UserControl yang memuat WebView2.
    ///
    /// Wajib public + ComVisible(true): Excel-DNA mendaftarkan kontrol ini
    /// sebagai ActiveX sementara agar bisa ditampung di Custom Task Pane Office.
    ///
    /// Alur inisialisasi (sekali per panel):
    ///   1. Ekstrak WebView2Loader.dll native → SetLoaderDllFolderPath.
    ///   2. Buat environment dengan user-data folder di %LOCALAPPDATA%.
    ///   3. Cegat semua request ke https://ckpn.panel/ → layani dari resource.
    ///   4. Buka https://ckpn.panel/index.html.
    /// Bila WebView2 Runtime tidak terpasang, tampilkan pesan pengganti
    /// (panel tidak crash dan tombol VBA lama tetap berfungsi).
    /// </summary>
    [ComVisible(true)]
    public class PanelHost : UserControl
    {
        // SetLoaderDllFolderPath hanya boleh dipanggil SEKALI per proses dan
        // sebelum environment pertama dibuat — makanya statis.
        private static bool _loaderSudahDiset;
        private static readonly object _kunciLoader = new object();

        // Satu environment dipakai bersama semua panel (mis. beberapa jendela
        // Excel) karena user-data folder yang sama tidak boleh dibuka dua
        // environment berbeda dalam satu proses.
        private static Task<CoreWebView2Environment> _envTask;

        private WebView2 _web;
        private Label    _pesan;
        private bool     _mulai;

        public PanelHost()
        {
            BackColor = Color.FromArgb(0xF4, 0xF3, 0xEE);   // sama dengan latar panel HTML → tidak "kedip" putih

            _pesan = new Label
            {
                Dock      = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                Font      = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(0x5B, 0x62, 0x5F),
                Text      = "Memuat Panel CKPN…"
            };
            Controls.Add(_pesan);
        }

        /// <summary>True bila WebView sudah siap menerima pesan.</summary>
        public bool Siap
        {
            get { return _web != null && _web.CoreWebView2 != null; }
        }

        // Inisialisasi dimulai saat handle jendela sudah ada (kontrol sudah
        // benar-benar ditampung task pane), bukan di constructor.
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (_mulai) return;
            _mulai = true;
            var _ = InisialisasiAsync();   // fire-and-forget; error ditangani di dalam
        }

        private async Task InisialisasiAsync()
        {
            try
            {
                // 0. Cek runtime lebih dulu agar pesan ke user jelas.
                string versiRuntime;
                try
                {
                    SetLoaderSekali();
                    versiRuntime = CoreWebView2Environment.GetAvailableBrowserVersionString();
                }
                catch (WebView2RuntimeNotFoundException)
                {
                    TampilkanPesan(
                        "Microsoft Edge WebView2 Runtime belum terpasang di komputer ini.\n\n" +
                        "Panel CKPN membutuhkan WebView2 Runtime (sudah bawaan Windows 11 " +
                        "dan Windows 10 yang Edge-nya ter-update).\n\n" +
                        "Hubungi IT untuk memasang \"WebView2 Runtime (Evergreen)\".\n\n" +
                        "Tombol perhitungan di sheet Master tetap dapat dipakai.");
                    return;
                }
                if (string.IsNullOrEmpty(versiRuntime))
                {
                    TampilkanPesan("WebView2 Runtime tidak terdeteksi.");
                    return;
                }

                // 1. Environment bersama
                CoreWebView2Environment env = await AmbilEnvironmentAsync();

                // 2. Kontrol WebView2
                _web = new WebView2
                {
                    Dock                  = DockStyle.Fill,
                    DefaultBackgroundColor = Color.FromArgb(0xF4, 0xF3, 0xEE)
                };
                Controls.Add(_web);
                _web.BringToFront();

                await _web.EnsureCoreWebView2Async(env);
                CoreWebView2 core = _web.CoreWebView2;

                // 3. Pengaturan: panel adalah aplikasi, bukan browser umum.
                core.Settings.AreDefaultContextMenusEnabled   = false;
                core.Settings.IsStatusBarEnabled              = false;
                core.Settings.IsZoomControlEnabled            = false;
                core.Settings.AreBrowserAcceleratorKeysEnabled = false;   // Ctrl+R/F5/Ctrl+P dll.
                core.Settings.IsWebMessageEnabled             = true;
#if DEBUG
                core.Settings.AreDevToolsEnabled = true;    // F12 hanya di build Debug
#else
                core.Settings.AreDevToolsEnabled = false;
#endif

                // 4. Layani aset panel dari resource DLL
                core.AddWebResourceRequestedFilter(PanelAssets.BaseUrl + "*", CoreWebView2WebResourceContext.All);
                core.WebResourceRequested += (s, e) => LayaniAset(env, e);

                // 5. Keamanan navigasi: hanya alamat panel yang boleh dibuka.
                core.NavigationStarting += (s, e) =>
                {
                    if (!e.Uri.StartsWith(PanelAssets.BaseUrl, StringComparison.OrdinalIgnoreCase))
                        e.Cancel = true;
                };
                core.NewWindowRequested += (s, e) => { e.Handled = true; };

                // 6. Pesan dari JS → C#
                core.WebMessageReceived += (s, e) =>
                {
                    // Tolak pesan dari sumber selain panel (defensif).
                    if (!e.Source.StartsWith(PanelAssets.BaseUrl, StringComparison.OrdinalIgnoreCase)) return;
                    PanelBridge.Terima(this, e.WebMessageAsJson);
                };

                core.Navigate(PanelAssets.BaseUrl + "index.html");
                _pesan.Visible = false;
            }
            catch (Exception ex)
            {
                TampilkanPesan("Panel CKPN gagal dimuat:\n\n" + ex.Message);
                CatatLog("Inisialisasi panel gagal: " + ex);
            }
        }

        private static void SetLoaderSekali()
        {
            lock (_kunciLoader)
            {
                if (_loaderSudahDiset) return;
                string folder = NativeLoader.SiapkanWebView2Loader();
                CoreWebView2Environment.SetLoaderDllFolderPath(folder);
                _loaderSudahDiset = true;
            }
        }

        private static Task<CoreWebView2Environment> AmbilEnvironmentAsync()
        {
            lock (_kunciLoader)
            {
                if (_envTask == null || _envTask.IsFaulted || _envTask.IsCanceled)
                {
                    Directory.CreateDirectory(AppPaths.FolderWebView2Data);
                    _envTask = CoreWebView2Environment.CreateAsync(null, AppPaths.FolderWebView2Data, null);
                }
                return _envTask;
            }
        }

        private static void LayaniAset(CoreWebView2Environment env, CoreWebView2WebResourceRequestedEventArgs e)
        {
            try
            {
                var uri = new Uri(e.Request.Uri);
                string mime;
                Stream isi = PanelAssets.Buka(uri.AbsolutePath, out mime);
                if (isi == null)
                {
                    e.Response = env.CreateWebResourceResponse(null, 404, "Not Found", "");
                    return;
                }
                // no-store: setelah .xll diganti versi baru, panel langsung
                // memakai file baru tanpa sisa cache lama.
                e.Response = env.CreateWebResourceResponse(
                    isi, 200, "OK",
                    "Content-Type: " + mime + "\r\nCache-Control: no-store");
            }
            catch
            {
                e.Response = env.CreateWebResourceResponse(null, 500, "Error", "");
            }
        }

        /// <summary>Kirim JSON ke panel. Harus dipanggil di thread UI Excel.</summary>
        public void KirimJson(string json)
        {
            if (!Siap) return;
            try { _web.CoreWebView2.PostWebMessageAsJson(json); }
            catch (Exception ex) { CatatLog("PostWebMessageAsJson gagal: " + ex.Message); }
        }

        /// <summary>
        /// Tahap 6: cetak halaman panel ke PDF A4. Yang tercetak hanya #cetak-root (CSS @media print),
        /// yaitu dokumen yang sedang dipratinjau. Dipanggil di thread UI.
        /// </summary>
        public Task<bool> CetakPdfAsync(string path)
        {
            if (!Siap) throw new InvalidOperationException("Panel belum siap.");
            CoreWebView2 core = _web.CoreWebView2;
            CoreWebView2PrintSettings s = core.Environment.CreatePrintSettings();
            s.Orientation = CoreWebView2PrintOrientation.Portrait;
            s.PageWidth = 8.27;      // A4 dalam inci
            s.PageHeight = 11.69;
            s.MarginTop = 0.47; s.MarginBottom = 0.55; s.MarginLeft = 0.51; s.MarginRight = 0.51;
            s.ShouldPrintBackgrounds = true;
            s.ShouldPrintHeaderAndFooter = false;
            return core.PrintToPdfAsync(path, s);
        }

        private void TampilkanPesan(string teks)
        {
            _pesan.Text    = teks;
            _pesan.Visible = true;
            _pesan.BringToFront();
        }

        /// <summary>Diteruskan ke CatatanLog (library\logs\panel_{user}.log).</summary>
        internal static void CatatLog(string baris)
        {
            CatatanLog.Tulis(baris);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _web != null)
            {
                try { _web.Dispose(); } catch { }
                _web = null;
            }
            base.Dispose(disposing);
        }
    }
}
