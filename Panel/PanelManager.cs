using System;
using System.Collections.Generic;
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

        /// <summary>Buka/tutup panel di jendela aktif.</summary>
        public static void Toggle()
        {
            CustomTaskPane ctp = AmbilAtauBuat();
            if (ctp == null) return;
            if (ctp.Visible) SimpanLebar(ctp);
            ctp.Visible = !ctp.Visible;
        }

        /// <summary>Tampilkan panel di jendela aktif.</summary>
        public static void Tampilkan()
        {
            CustomTaskPane ctp = AmbilAtauBuat();
            if (ctp != null) ctp.Visible = true;
        }

        /// <summary>Semua panel yang sudah dibuat (untuk broadcast event progres).</summary>
        public static IEnumerable<PanelHost> SemuaHost()
        {
            foreach (var ctp in _pane.Values)
            {
                PanelHost host = null;
                try { host = ctp.ContentControl as PanelHost; } catch { }
                if (host != null) yield return host;
            }
        }

        /// <summary>Hapus semua pane. Dipanggil saat add-in di-unload (AutoClose).</summary>
        public static void TutupSemua()
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
