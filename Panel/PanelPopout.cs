using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Tahap 9a — Panel CKPN di jendela sendiri ("pisahkan panel").
    ///
    /// Kenapa bukan task pane yang di-float: task pane Office tetap menempel ke satu jendela workbook,
    /// sulit dipindah ke monitor lain, dan WebView2 kadang kosong setelah bingkainya dipindah Office.
    /// Jendela biasa bebas digeser ke monitor mana pun, bisa dimaksimalkan, dan muncul di taskbar.
    ///
    /// Jendela sengaja TIDAK di-own oleh jendela Excel: di Excel 2013+ (SDI) setiap workbook punya
    /// jendela sendiri. Kalau jendela pemilik ditutup, Windows ikut menghancurkan jendela panel.
    /// Sebagai gantinya ada pilihan "Selalu di atas" di menu sistem (klik kanan judul jendela).
    ///
    /// Isinya PanelHost yang sama dengan task pane (ModePopout = true). Semua perintah bridge tetap
    /// lewat QueueAsMacro di thread Excel, karena jendela ini dibuat di thread UI Excel.
    /// Posisi, ukuran, status maksimal, dan "selalu di atas" diingat per user di panel_jendela.txt.
    /// </summary>
    internal sealed class PanelPopout : Form
    {
        // ID menu sistem buatan sendiri: kelipatan 16 dan di bawah 0xF000 (4 bit bawah dipakai Windows)
        private const int IdSelaluDiAtas = 0x1F00, IdTempel = 0x1F10;
        private const int WM_SYSCOMMAND = 0x0112;
        private const int MF_STRING = 0x0, MF_SEPARATOR = 0x800, MF_CHECKED = 0x8, MF_UNCHECKED = 0x0;

        [DllImport("user32.dll")] private static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern bool AppendMenu(IntPtr hMenu, int uFlags, int uIDNewItem, string lpNewItem);
        [DllImport("user32.dll")] private static extern int CheckMenuItem(IntPtr hMenu, int uIDCheckItem, int uCheck);

        private const int LebarAwal = 560, TinggiAwal = 820;

        public PanelHost Host { get; private set; }

        /// <summary>True bila jendela ditutup karena add-in di-unload: jangan tempel kembali ke task pane.</summary>
        public bool TutupTanpaTempel { get; set; }

        private static string FileJendela
        {
            get { return Path.Combine(AppPaths.FolderLokal, "panel_jendela.txt"); }
        }

        public PanelPopout()
        {
            Text            = PanelManager.Judul;
            StartPosition   = FormStartPosition.Manual;
            ShowInTaskbar   = true;                       // mudah ditemukan bila tertutup Excel
            MinimumSize     = new Size(380, 420);
            BackColor       = Color.FromArgb(0xF4, 0xF3, 0xEE);
            KeyPreview      = false;                      // keyboard sepenuhnya milik WebView2

            // Ikon Excel agar jendela mudah dikenali di taskbar / Alt+Tab
            try { Icon = Icon.ExtractAssociatedIcon(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName); }
            catch { /* ikon bawaan */ }

            Host = new PanelHost { Dock = DockStyle.Fill, ModePopout = true };
            Controls.Add(Host);

            MuatPosisi();
        }

        // ------------------------------------------------------------
        // Menu sistem: "Selalu di atas" dan "Tempel kembali ke Excel"
        // ------------------------------------------------------------
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            try
            {
                IntPtr m = GetSystemMenu(Handle, false);
                AppendMenu(m, MF_SEPARATOR, 0, null);
                AppendMenu(m, MF_STRING | (TopMost ? MF_CHECKED : MF_UNCHECKED), IdSelaluDiAtas, "Selalu di atas");
                AppendMenu(m, MF_STRING, IdTempel, "Tempel kembali ke Excel");
            }
            catch { /* menu tambahan tidak penting */ }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_SYSCOMMAND)
            {
                int cmd = (int)(m.WParam.ToInt64() & 0xFFF0);
                if (cmd == IdSelaluDiAtas)
                {
                    TopMost = !TopMost;
                    try { CheckMenuItem(GetSystemMenu(Handle, false), IdSelaluDiAtas, TopMost ? MF_CHECKED : MF_UNCHECKED); } catch { }
                    SimpanPosisi();
                    return;
                }
                if (cmd == IdTempel)
                {
                    Close();      // FormClosed di PanelManager menempelkan kembali ke task pane
                    return;
                }
            }
            base.WndProc(ref m);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            SimpanPosisi();
            base.OnFormClosing(e);
        }

        /// <summary>Munculkan jendela ke depan (dipakai tombol "Panel CKPN" saat panel sudah terpisah).</summary>
        public void Munculkan()
        {
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            if (!Visible) Show();
            Activate();
            BringToFront();
        }

        // ------------------------------------------------------------
        // Posisi & ukuran: "x;y;lebar;tinggi;maks;atas" (piksel layar, invariant)
        // ------------------------------------------------------------
        private void MuatPosisi()
        {
            Rectangle r = Rectangle.Empty;
            bool maks = false, atas = false;
            try
            {
                if (File.Exists(FileJendela))
                {
                    string[] p = File.ReadAllText(FileJendela).Trim().Split(';');
                    if (p.Length >= 4)
                    {
                        r = new Rectangle(int.Parse(p[0], CultureInfo.InvariantCulture), int.Parse(p[1], CultureInfo.InvariantCulture),
                                          int.Parse(p[2], CultureInfo.InvariantCulture), int.Parse(p[3], CultureInfo.InvariantCulture));
                        maks = p.Length > 4 && p[4] == "1";
                        atas = p.Length > 5 && p[5] == "1";
                    }
                }
            }
            catch { r = Rectangle.Empty; }

            // Posisi tersimpan hanya dipakai bila masih terlihat di salah satu layar
            // (mis. monitor kedua sudah dicabut → kembali ke posisi awal).
            if (r.Width < MinimumSize.Width || r.Height < MinimumSize.Height || !TerlihatDiLayar(r))
                r = PosisiAwal();

            Bounds = r;
            TopMost = atas;
            if (maks) WindowState = FormWindowState.Maximized;
        }

        private static bool TerlihatDiLayar(Rectangle r)
        {
            foreach (Screen s in Screen.AllScreens)
            {
                Rectangle x = Rectangle.Intersect(s.WorkingArea, r);
                if (x.Width >= 120 && x.Height >= 80) return true;   // judul jendela cukup terlihat untuk digeser
            }
            return false;
        }

        // Awal: sisi kanan layar tempat kursor berada (layar tempat user menekan tombol).
        private static Rectangle PosisiAwal()
        {
            Rectangle wa = Screen.FromPoint(Cursor.Position).WorkingArea;
            int w = Math.Min(LebarAwal, wa.Width), h = Math.Min(TinggiAwal, wa.Height);
            return new Rectangle(wa.Right - w - 8, wa.Top + Math.Max(0, (wa.Height - h) / 2), w, h);
        }

        private void SimpanPosisi()
        {
            try
            {
                Rectangle r = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
                bool maks = WindowState == FormWindowState.Maximized;
                Directory.CreateDirectory(AppPaths.FolderLokal);
                File.WriteAllText(FileJendela, string.Join(";", new[]
                {
                    r.X.ToString(CultureInfo.InvariantCulture), r.Y.ToString(CultureInfo.InvariantCulture),
                    r.Width.ToString(CultureInfo.InvariantCulture), r.Height.ToString(CultureInfo.InvariantCulture),
                    maks ? "1" : "0", TopMost ? "1" : "0"
                }));
            }
            catch { /* tidak penting */ }
        }
    }
}
