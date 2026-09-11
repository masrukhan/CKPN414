namespace CKPNLibrary.Models
{
    /// <summary>
    /// Mewakili satu baris kontrak/rekening milik sebuah CIF.
    /// Setara dengan Array(NoKontrak, OS, Jaminan, Kualitas, SheetKC) di VBA.
    /// </summary>
    public class KontrakData
    {
        /// <summary>Nomor kontrak / nomor rekening.</summary>
        public string NoKontrak { get; set; }

        /// <summary>Outstanding (saldo pokok) — diisi dari baris pertama kontrak ini.</summary>
        public double OS { get; set; }

        /// <summary>Nilai jaminan — diakumulasi dari semua baris kontrak yang sama.</summary>
        public double Jaminan { get; set; }

        /// <summary>Kualitas piutang (1–5). Diisi dari baris pertama kontrak ini.</summary>
        public int Kualitas { get; set; }

        /// <summary>Nama sheet KC sumber data (mis. "KC0600").</summary>
        public string SheetKC { get; set; }

        /// <summary>
        /// Hari tunggakan kontrak — dipakai kriteria penurunan nilai 7–30 hari.
        /// Diisi dari baris pertama; diperbarui ke nilai MAKSIMUM bila kontrak
        /// yang sama muncul di beberapa baris (tunggakan terlama yang dipakai).
        /// </summary>
        public double HariTunggakan { get; set; }

        public KontrakData(string noKontrak, double os, double jaminan, int kualitas, string sheetKC, double hariTunggakan = 0)
        {
            NoKontrak     = noKontrak;
            OS            = os;
            Jaminan       = jaminan;
            Kualitas      = kualitas;
            SheetKC       = sheetKC;
            HariTunggakan = hariTunggakan;
        }
    }
}