using System;
using System.Collections.Generic;
using CKPNLibrary.Helpers;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Mengelola Custom Task Pane "Panel CKPN".
    ///
    /// Sejak Excel 2013 setiap workbook punya jendela sendiri (SDI), dan task
    /// pane menempel ke SATU jendela. Maka pane disimpan per jendela (kunci:
    /// Hwnd jendela Excel). Tombol "Panel CKPN" membuka/menutup pane di
    /// jendela yang sedang aktif saja.
    ///
    /// Tahap 9a: panel bisa dipisah ke jendela sendiri (<see cref="PanelPopout"/>).
    /// Selama terpisah hanya ada SATU panel: semua task pane dihapus, dan tombol
    /// "Panel CKPN" memunculkan jendela itu. Menutup jendela → panel kembali ke
    /// samping jendela Excel yang aktif.
    /// </summary>
    internal static class PanelManager
    {
        public const string Judul = "Panel CKPN";
        public const int    Lebar = 480;          // lebar awal; tabel Overview per KC butuh ±480 px
        private const int   LebarMin = 360, LebarMaks = 1000;

        /// <summary>Lebar terakhir yang digeser user, disimpan per user di PC ini.</summary>
        private static string FileLebar
        {
            get { return System.IO.Path.Combine(Helpers.AppPaths.FolderLokal, "panel_lebar.txt"); }
        }

        private static int LebarTersimpan()
        {
            try
            {
                int v;
                if (System.IO.File.Exists(FileLebar) &&
                    int.TryParse(System.IO.File.ReadAllText(FileLebar).Trim(), out v) && v >= LebarMin && v <= LebarMaks)
                    return v;
            }
            catch { }
            return Lebar;
        }

        private static void SimpanLebar(CustomTaskPane ctp)
        {
            try
            {
                if (ctp.DockPosition != MsoCTPDockPosition.msoCTPDockPositionRight &&
                    ctp.DockPosition != MsoCTPDockPosition.msoCTPDockPositionLeft) return;
                int w = ctp.Width;
                if (w < LebarMin || w > LebarMaks) return;
                System.IO.Directory.CreateDirectory(Helpers.AppPaths.FolderLokal);
                System.IO.File.WriteAllText(FileLebar, w.ToString());
            }
            catch { /* tidak penting */ }
        }

        private static readonly Dictionary<int, CustomTaskPane> _pane = new Dictionary<int, CustomTaskPane>();

        // Tahap 9a: jendela panel terpisah (null = panel menempel di task pane)
        private static PanelPopout _popout;

        /// <summary>True bila panel sedang dipisah ke jendela sendiri.</summary>
        public static bool Terpisah
        {
            get { return _popout != null && !_popout.IsDisposed; }
        }

        /// <summary>Buka/tutup panel di jendela aktif. Bila panel terpisah: munculkan jendelanya.</summary>
        public static void Toggle()
        {
            if (Terpisah) { _popout.Munculkan(); return; }
            CustomTaskPane ctp = AmbilAtauBuat();
            if (ctp == null) return;
            if (ctp.Visible) SimpanLebar(ctp);
            ctp.Visible = !ctp.Visible;
        }

        /// <summary>Tampilkan panel di jendela aktif (atau munculkan jendela terpisah).</summary>
        public static void Tampilkan()
        {
            if (Terpisah) { _popout.Munculkan(); return; }
            CustomTaskPane ctp = AmbilAtauBuat();
            if (ctp != null) ctp.Visible = true;
        }

        /// <summary>
        /// Tahap 9a: pindahkan panel ke jendela sendiri. Dipanggil lewat QueueAsMacro (bukan langsung
        /// dari handler pesan WebView) karena task pane pengirim pesan ikut dihapus di sini.
        /// </summary>
        public static void Pisahkan()
        {
            if (Terpisah) { _popout.Munculkan(); return; }

            var jendela = new PanelPopout();
            jendela.FormClosed += (s, e) =>
            {
                bool tempel = !jendela.TutupTanpaTempel;
                if (ReferenceEquals(_popout, jendela)) _popout = null;
                // Tempel kembali ke jendela Excel yang aktif. Ditunda agar jendela selesai ditutup
                // dan Excel dalam kondisi aman untuk membuat task pane.
                if (tempel)
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { Tampilkan(); }
                        catch (Exception ex) { CatatanLog.Tulis("Tempel panel kembali gagal: " + ex.Message); }
                    });
            };
            _popout = jendela;
            jendela.Show();
            jendela.Activate();

            // Satu panel saja: task pane di semua jendela workbook dihapus (lebar diingat dulu).
            HapusSemuaPane();
            CatatanLog.Tulis("Panel dipisah ke jendela sendiri.");
        }

        /// <summary>Tahap 9a: tutup jendela terpisah dan kembalikan panel ke samping jendela Excel aktif.</summary>
        public static void Tempel()
        {
            if (!Terpisah) { Tampilkan(); return; }
            _popout.Close();   // FormClosed → Tampilkan()
            CatatanLog.Tulis("Panel ditempel kembali ke Excel.");
        }

        /// <summary>Semua panel yang sudah dibuat (untuk broadcast event progres).</summary>
        public static IEnumerable<PanelHost> SemuaHost()
        {
            if (Terpisah && _popout.Host != null && !_popout.Host.IsDisposed) yield return _popout.Host;
            foreach (var ctp in _pane.Values)
            {
                PanelHost host = null;
                try { host = ctp.ContentControl as PanelHost; } catch { }
                if (host != null) yield return host;
            }
        }

        /// <summary>Hapus semua pane dan jendela terpisah. Dipanggil saat add-in di-unload (AutoClose).</summary>
        public static void TutupSemua()
        {
            if (Terpisah)
            {
                try { _popout.TutupTanpaTempel = true; _popout.Close(); } catch { }
                _popout = null;
            }
            HapusSemuaPane();
        }

        private static void HapusSemuaPane()
        {
            foreach (var ctp in _pane.Values)
            {
                try { if (ctp.Visible) SimpanLebar(ctp); } catch { }
                try { ctp.Delete(); } catch { }
            }
            _pane.Clear();
        }

        private static CustomTaskPane AmbilAtauBuat()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            Excel.Window win = app.ActiveWindow;
            if (win == null) return null;   // tidak ada workbook terbuka

            int hwnd = win.Hwnd;
            CustomTaskPane ctp;
            if (_pane.TryGetValue(hwnd, out ctp))
            {
                // Pane bisa menjadi tidak valid bila jendelanya sudah ditutup.
                try { var _ = ctp.Visible; return ctp; }
                catch { _pane.Remove(hwnd); }
            }

            BersihkanYangMati();

            ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(PanelHost), Judul);
            ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            ctp.Width        = LebarTersimpan();
            // Lebar diingat saat panel disembunyikan (tombol Panel CKPN / tanda X) dan saat add-in ditutup
            ctp.VisibleStateChange += p => { try { if (!p.Visible) SimpanLebar(p); } catch { } };
            _pane[hwnd] = ctp;
            return ctp;
        }

        // Buang pane milik jendela yang sudah ditutup user.
        private static void BersihkanYangMati()
        {
            var mati = new List<int>();
            foreach (var kv in _pane)
            {
                try { var _ = kv.Value.Visible; }
                catch { mati.Add(kv.Key); }
            }
            foreach (int k in mati) _pane.Remove(k);
        }
    }
}
