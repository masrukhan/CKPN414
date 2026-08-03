using System.Collections.Generic;

namespace CKPNLibrary.Models
{
    /// <summary>
    /// Mewakili satu CIF beserta semua kontrak yang dimilikinya.
    /// Setara dengan Array(CIF, Nama, SheetKC, TotalOS, kontrakDict) di VBA.
    /// </summary>
    public class CIFData
    {
        /// <summary>Nomor CIF (dapat mengandung leading-zero).</summary>
        public string CIF { get; set; }

        /// <summary>Nama debitur.</summary>
        public string Nama { get; set; }

        /// <summary>Sheet KC pertama yang mencatat CIF ini.</summary>
        public string SheetKC { get; set; }

        /// <summary>Total OS seluruh kontrak CIF ini (diakumulasi saat baca data).</summary>
        public double TotalOS { get; set; }

        /// <summary>
        /// Dictionary kontrak: key = NoKontrak, value = KontrakData.
        /// Digunakan untuk deduplikasi kontrak dan akumulasi jaminan.
        /// </summary>
        public Dictionary<string, KontrakData> Kontrak { get; set; }

        public CIFData(string cif, string nama, string sheetKC)
        {
            CIF     = cif;
            Nama    = nama;
            SheetKC = sheetKC;
            TotalOS = 0;
            Kontrak = new Dictionary<string, KontrakData>(System.StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Tambahkan atau perbarui kontrak.
        /// — Baris pertama kontrak: catat OS, jaminan awal, kualitas; tambahkan ke TotalOS.
        /// — Baris berikutnya kontrak yang sama: akumulasi jaminan saja (OS tidak ditambah lagi).
        /// </summary>
        public void TambahKontrak(string noKontrak, double os, double jaminan, int kualitas, string sheetKC)
        {
            if (!Kontrak.ContainsKey(noKontrak))
            {
                // Baris pertama kontrak ini
                TotalOS += os;
                Kontrak[noKontrak] = new KontrakData(noKontrak, os, jaminan, kualitas, sheetKC);
            }
            else
            {
                // Baris berikutnya: akumulasi jaminan saja
                Kontrak[noKontrak].Jaminan += jaminan;
            }
        }
    }
}