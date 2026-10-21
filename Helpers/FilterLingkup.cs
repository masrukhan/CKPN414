using System;
using System.Collections.Generic;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Tahap 8 — saringan baris untuk modul perhitungan pada mode selain "per akad".
    ///
    /// Panel menulis lingkup grup ke Master!H8, mis. "kantor:001,002". Modul memanggil
    /// <see cref="Dari"/> sekali di awal perhitungan; hasil null = tanpa saringan (mode akad,
    /// edisi Lite, atau H8 kosong) sehingga perilaku lama tidak berubah.
    ///
    ///   var f = FilterLingkup.Dari(_app);
    ///   ...
    ///   if (f != null &amp;&amp; !f.Cocok(nilaiKolomB)) continue;   // baris di luar grup
    ///
    /// Kolom yang dibaca mengikuti mode (<see cref="Kolom"/>: "B" untuk sandi kantor). Untuk sheet
    /// KC0600–KC1100 dan KC2900, sandi kantor ada di kolom B pada SEMUA baris, termasuk baris agunan.
    /// </summary>
    internal sealed class FilterLingkup
    {
        public string ModeId { get; private set; }
        /// <summary>Huruf kolom sumber saringan di sheet template (mis. "B").</summary>
        public string Kolom { get; private set; }
        /// <summary>Nomor kolom (1-based) dari <see cref="Kolom"/>.</summary>
        public int NomorKolom { get; private set; }
        /// <summary>Kode grup kanonik, mis. "001,002".</summary>
        public string Kode { get; private set; }
        private HashSet<string> _kode;
        private ModeHitung.Mode _mode;

        /// <summary>Teks untuk log/label, mis. "sandi kantor 001, 002".</summary>
        public string Label { get { return ModeHitung.LabelKode(_mode, Kode); } }

        /// <summary>
        /// Saringan aktif dari Master!H8 workbook aplikasi; null bila tidak ada saringan.
        /// Aman dipanggil dari modul (tombol Master maupun panel).
        /// </summary>
        public static FilterLingkup Dari(Excel.Application app)
        {
            try
            {
                ModeHitung.Mode m = ModeHitung.Aktif;
                if (m.Utama || string.IsNullOrEmpty(m.KolomFilter)) return null;
                Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app) ?? app.ActiveWorkbook;
                Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
                string kode = ModeHitung.BacaFilter(master, m);
                if (kode.Length == 0) return null;
                return Buat(m, kode);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca saringan lingkup Master!" + ModeHitung.SelFilter + " gagal: " + ex.Message);
                return null;
            }
        }

        internal static FilterLingkup Buat(ModeHitung.Mode m, string kode)
        {
            var f = new FilterLingkup
            {
                _mode = m, ModeId = m.Id, Kolom = m.KolomFilter, NomorKolom = NomorDariHuruf(m.KolomFilter),
                Kode = ModeHitung.NormKode(m, kode)
            };
            f._kode = new HashSet<string>(f.Kode.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries), StringComparer.OrdinalIgnoreCase);
            return f;
        }

        /// <summary>True bila nilai sel kolom saringan (mis. "001", 1.0) termasuk grup ini.</summary>
        public bool Cocok(object nilaiSel)
        {
            return _kode.Contains(ModeHitung.NormUnit(_mode, nilaiSel));
        }

        /// <summary>True bila baris r (1-based, sesuai array Value2) pada data template termasuk grup ini.</summary>
        public bool CocokBaris(object[,] data, int r)
        {
            if (data == null || NomorKolom < data.GetLowerBound(1) || NomorKolom > data.GetUpperBound(1)) return false;
            return Cocok(data[r, NomorKolom]);
        }

        public static int NomorDariHuruf(string huruf)
        {
            int n = 0;
            foreach (char c in (huruf ?? "").ToUpperInvariant())
                if (c >= 'A' && c <= 'Z') n = n * 26 + (c - 'A' + 1);
            return n;
        }
    }
}
