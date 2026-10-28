using System;
using System.Collections.Generic;
using System.Globalization;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Kelompok sektor ekonomi untuk mode perhitungan "per kelompok sektor ekonomi".
    ///
    /// Sandi sektor ekonomi APOLLO (kolom T sheet KC0600–KC1100, 6 digit) mengikuti struktur KBLI 2020:
    /// dua digit pertama = golongan pokok (01–99), "00xxxx" = bukan lapangan usaha (rumah tangga).
    /// Ratusan sandi diringkas menjadi 10 kelompok tetap supaya tiap kelompok cukup besar untuk PD/LGD;
    /// kelompok dapat digabung lagi di susunan grup (mis. 08–10 menjadi satu grup konsumtif).
    ///
    /// Pemetaan memakai awalan, bukan daftar sandi lengkap, sehingga sandi baru tetap masuk kelompok
    /// yang benar. Sandi kosong / tidak terbaca → kelompok cadangan 07 (dihitung dan dilaporkan).
    /// </summary>
    internal static class SektorEkonomi
    {
        public const string Cadangan = "07";

        /// <summary>Kode kelompok dan namanya (urut tampil).</summary>
        public static readonly string[,] Kelompok =
        {
            { "01", "Pertanian, Kehutanan & Perikanan" },
            { "02", "Industri & Pertambangan" },
            { "03", "Konstruksi & Real Estat (Developer)" },
            { "04", "Perdagangan" },
            { "05", "Akomodasi & Makan Minum" },
            { "06", "Transportasi, Pergudangan & Komunikasi" },
            { "07", "Jasa & Lainnya" },
            { "08", "Konsumtif Penghasilan Tetap – Pemilikan Rumah" },
            { "09", "Konsumtif Penghasilan Tetap – Kendaraan" },
            { "10", "Konsumtif Penghasilan Tetap – Multiguna" }
        };

        /// <summary>Rentang golongan pokok KBLI 2020 (2 digit) per kelompok, untuk keterangan di panel.</summary>
        public static readonly string[,] Rentang =
        {
            { "01", "01–03" }, { "02", "05–33" }, { "03", "41–43, 68" }, { "04", "45–47" }, { "05", "55–56" },
            { "06", "49–53, 58–63" }, { "07", "35–39, 64–66, 69–99 dan sandi lain" },
            { "08", "001xxx" }, { "09", "002xxx" }, { "10", "003xxx, 004xxx, 009xxx" }
        };

        public static string Nama(string kode)
        {
            for (int i = 0; i < Kelompok.GetLength(0); i++) if (Kelompok[i, 0] == kode) return Kelompok[i, 1];
            return kode;
        }

        public static bool AdaKode(string kode)
        {
            for (int i = 0; i < Kelompok.GetLength(0); i++) if (Kelompok[i, 0] == kode) return true;
            return false;
        }

        /// <summary>
        /// Sandi sektor 6 digit kanonik dari isian sel: "009000    " / 9000.0 / "'004190" / "410119 - Konstruksi" → "009000" dst.
        /// "" bila kosong / tidak berupa angka.
        /// </summary>
        public static string NormSandi(object v)
        {
            string d = Digit(v);
            if (d.Length < 3 || d.Length > 6) return "";
            return d.PadLeft(6, '0');
        }

        /// <summary>
        /// Kode kelompok ("01"–"10") dari isian sel sandi sektor ATAU dari kode kelompok yang sudah kanonik.
        /// 1–2 digit = kode kelompok; 3–6 digit = sandi sektor. "" bila kosong / tidak dikenal.
        /// </summary>
        public static string NormKelompok(object v)
        {
            string d = Digit(v);
            if (d.Length == 0)
            {
                // Teks nama kelompok (mis. isian manual "Perdagangan")
                string s = v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
                for (int i = 0; i < Kelompok.GetLength(0) && s.Length > 0; i++)
                    if (string.Equals(Kelompok[i, 1], s, StringComparison.OrdinalIgnoreCase)) return Kelompok[i, 0];
                return "";
            }
            if (d.Length <= 2)
            {
                int k;
                if (!int.TryParse(d, out k) || k < 1 || k > Kelompok.GetLength(0)) return "";
                return k.ToString("00", CultureInfo.InvariantCulture);
            }
            return DariSandi(NormSandi(d));
        }

        /// <summary>Kelompok dari sandi sektor 6 digit (lihat tabel <see cref="Rentang"/>).</summary>
        public static string DariSandi(string sandi6)
        {
            if (string.IsNullOrEmpty(sandi6) || sandi6.Length != 6) return "";
            int gp;
            if (!int.TryParse(sandi6.Substring(0, 2), out gp)) return "";
            if (gp == 0)
            {
                // Bukan lapangan usaha: 001 pemilikan rumah, 002 kendaraan, lainnya (peralatan, multiguna, 009) multiguna
                char c = sandi6[2];
                if (c == '1') return "08";
                if (c == '2') return "09";
                return "10";
            }
            if (gp >= 1 && gp <= 3) return "01";
            if (gp >= 5 && gp <= 33) return "02";
            if ((gp >= 41 && gp <= 43) || gp == 68) return "03";
            if (gp >= 45 && gp <= 47) return "04";
            if (gp >= 55 && gp <= 56) return "05";
            if ((gp >= 49 && gp <= 53) || (gp >= 58 && gp <= 63)) return "06";
            return "07";
        }

        private static string Digit(object v)
        {
            if (v == null || v is DBNull) return "";
            string s = v is double ? ((long)Math.Round((double)v)).ToString(CultureInfo.InvariantCulture)
                                   : Convert.ToString(v, CultureInfo.InvariantCulture);
            s = s.Replace(' ', ' ').Trim().Trim('\'', '"').Trim();
            var sb = new System.Text.StringBuilder();
            foreach (char c in s)
            {
                if (c >= '0' && c <= '9') sb.Append(c);
                else if (sb.Length > 0) break;
                else if (!char.IsWhiteSpace(c)) return "";   // diawali huruf → bukan sandi
            }
            return sb.ToString();
        }

        /// <summary>Daftar kelompok untuk panel (kode, nama, rentang).</summary>
        public static List<Dictionary<string, object>> Daftar()
        {
            var l = new List<Dictionary<string, object>>();
            for (int i = 0; i < Kelompok.GetLength(0); i++)
                l.Add(new Dictionary<string, object> { { "kode", Kelompok[i, 0] }, { "nama", Kelompok[i, 1] }, { "rentang", Rentang[i, 1] } });
            return l;
        }
    }
}
