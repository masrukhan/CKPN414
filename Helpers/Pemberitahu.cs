using System.Collections.Generic;
using System.Windows.Forms;
using CKPNLibrary.Panel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Pengganti MessageBox.Show di dalam modul perhitungan.
    ///
    /// Modul yang sama dipanggil dari dua pintu:
    ///   1. Tombol VBA di sheet Master  → pesan tetap tampil sebagai MessageBox
    ///      (perilaku lama, tidak berubah).
    ///   2. Panel CKPN (pipeline)       → pesan TIDAK memblok proses; dikirim ke
    ///      panel sebagai kartu hasil dan dicatat ke log.
    ///
    /// Tanpa ini, setiap modul akan memunculkan dialog di tengah pipeline dan
    /// user harus menekan OK berkali-kali sebelum langkah berikutnya jalan.
    ///
    /// Semua method dipanggil di thread utama Excel (modul berjalan di sana),
    /// jadi tidak perlu sinkronisasi.
    /// </summary>
    internal static class Pemberitahu
    {
        /// <summary>True selama pipeline panel berjalan.</summary>
        public static bool ModePanel { get; private set; }

        private static readonly List<Dictionary<string, object>> _terkumpul =
            new List<Dictionary<string, object>>();

        public static void MulaiModePanel()
        {
            _terkumpul.Clear();
            ModePanel = true;
        }

        /// <summary>Akhiri mode panel dan kembalikan semua pesan modul yang terkumpul.</summary>
        public static List<Dictionary<string, object>> SelesaiModePanel()
        {
            ModePanel = false;
            var salinan = new List<Dictionary<string, object>>(_terkumpul);
            _terkumpul.Clear();
            return salinan;
        }

        /// <summary>
        /// Pesan ringkasan dari modul (dulu MessageBox di akhir modul).
        /// peringatan = true → ikon Warning di mode VBA / kartu kuning di panel.
        /// </summary>
        public static void Info(string judul, string pesan, bool peringatan = false)
        {
            if (!ModePanel)
            {
                MessageBox.Show(pesan, judul, MessageBoxButtons.OK,
                    peringatan ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                return;
            }

            var item = new Dictionary<string, object>
            {
                { "judul", judul }, { "pesan", pesan }, { "peringatan", peringatan }
            };
            _terkumpul.Add(item);
            CatatanLog.Tulis("[" + judul + "] " + (pesan ?? "").Replace("\r", "").Replace("\n", " | "));
            try { PanelBridge.Siarkan("pesanModul", item); } catch { }
        }

        /// <summary>
        /// Progres di dalam satu modul (mis. bulan ke-9 dari 12 di PD Net Flow).
        /// Tidak melakukan apa-apa saat dipanggil dari tombol VBA.
        /// </summary>
        public static void Progres(string modul, int ke, int dari, string label)
        {
            if (!ModePanel) return;
            try
            {
                PanelBridge.Siarkan("progresDetail", new Dictionary<string, object>
                {
                    { "modul", modul }, { "ke", ke }, { "dari", dari }, { "label", label ?? "" }
                });
            }
            catch { }
        }
    }
}
