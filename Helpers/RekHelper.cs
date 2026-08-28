using System;
using System.Globalization;
using System.Text;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Normalisasi nomor rekening &amp; tanggal yang dibaca dari sel Excel.
    ///
    /// Tujuan: pencocokan nomor rekening antar file/sheet tidak boleh gagal
    /// hanya karena perbedaan tipe sel (Text vs Number). Kasus yang ditangani:
    ///   - Sel Number panjang  -> Value2 = double -> "6.0012345E+15"
    ///   - Sel Text            -> "0060012345" (leading zero utuh)
    ///   - Apostrof, spasi, non-breaking space (hasil copy-paste dari web/PDF)
    ///   - Angka desimal palsu -> "12345.00"
    ///
    /// Dipakai oleh PDMigration (KC2900 / staging) dan dapat dipakai ulang
    /// oleh modul lain yang mencocokkan nomor rekening antar file.
    /// </summary>
    internal static class RekHelper
    {
        // ================================================================
        // NormNoRek: object (hasil Value2) -> string nomor rekening bersih
        //
        // Nilai dari Range.Value2 bisa berupa string (sel Text) atau double
        // (sel Number/General). Keduanya harus menghasilkan string yang sama
        // bentuknya supaya bisa dibandingkan.
        // ================================================================
        public static string NormNoRek(object val)
        {
            if (val == null) return "";

            if (val is double)
                return Bersihkan(((double)val).ToString("F0", CultureInfo.InvariantCulture));
            if (val is float)
                return Bersihkan(((float)val).ToString("F0", CultureInfo.InvariantCulture));
            if (val is decimal)
                return Bersihkan(((decimal)val).ToString("F0", CultureInfo.InvariantCulture));
            if (val is int)
                return Bersihkan(((int)val).ToString(CultureInfo.InvariantCulture));
            if (val is long)
                return Bersihkan(((long)val).ToString(CultureInfo.InvariantCulture));

            return Bersihkan(val.ToString());
        }

        // ================================================================
        // Bersihkan: buang apostrof/spasi, lalu pulihkan notasi ilmiah
        //
        // Contoh:
        //   "  0060012345 "   -> "0060012345"
        //   "'0060012345"     -> "0060012345"
        //   "6.0012345E+15"   -> "6001234500000000"
        //   "12345.00"        -> "12345"
        // ================================================================
        public static string Bersihkan(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";

            var sb = new StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '\'' || c == '"') continue;                  // apostrof teks Excel
                if (c == '\u00A0' || char.IsWhiteSpace(c)) continue;  // spasi & non-breaking space
                sb.Append(c);
            }
            s = sb.ToString();
            if (s.Length == 0) return "";

            // "6.0012345E+15" atau "12345.00" -> kembalikan ke bentuk bulat
            if (s.IndexOf('E') >= 0 || s.IndexOf('e') >= 0 || s.IndexOf('.') >= 0)
            {
                decimal d;
                if (decimal.TryParse(s, NumberStyles.Float,
                                     CultureInfo.InvariantCulture, out d)
                    && d == decimal.Truncate(d) && d >= 0)
                {
                    return d.ToString("F0", CultureInfo.InvariantCulture);
                }
            }
            return s;
        }

        // ================================================================
        // Longgar: kunci cadangan tanpa leading zero.
        //
        // Dipakai HANYA sebagai fallback jika kunci ketat tidak ketemu,
        // mis. satu file menyimpan "0060012345" (Text) dan file lain
        // menyimpan 60012345 (Number, leading zero sudah hilang di sumber).
        //
        // Jika nomor mengandung huruf, dikembalikan apa adanya (upper-case)
        // supaya tidak salah potong.
        // ================================================================
        public static string Longgar(string rekNorm)
        {
            if (string.IsNullOrEmpty(rekNorm)) return "";

            for (int i = 0; i < rekNorm.Length; i++)
                if (!char.IsDigit(rekNorm[i]))
                    return rekNorm.ToUpperInvariant();

            string t = rekNorm.TrimStart('0');
            return t.Length == 0 ? "0" : t;
        }

        // ================================================================
        // NormYyyymmdd: object -> long yyyyMMdd (0 jika tidak valid)
        //
        // Kolom tanggal WO di KC2900 bisa tersimpan dalam beberapa bentuk:
        //   45688        (sel Date -> Value2 = OADate)  -> 20250131
        //   20250131     (Number / Text)                -> 20250131
        //   "31/01/2025" (Text)                         -> 20250131
        //
        // Tanpa normalisasi ini, sel bertipe Date akan selalu gugur di
        // filter periode (45688 < 20250101) sehingga dictWO kosong.
        // ================================================================
        public static long NormYyyymmdd(object val)
        {
            if (val == null) return 0;

            if (val is double || val is int || val is long || val is decimal || val is float)
            {
                double d = Convert.ToDouble(val, CultureInfo.InvariantCulture);

                // Sudah dalam bentuk yyyyMMdd
                if (d >= 19000101 && d <= 29991231) return (long)d;

                // OADate Excel (1900-01-01 = 1 ... 2100-ish < 75000)
                if (d > 0 && d < 200000)
                {
                    try
                    {
                        DateTime dt = DateTime.FromOADate(d);
                        return dt.Year * 10000L + dt.Month * 100L + dt.Day;
                    }
                    catch { return 0; }
                }
                return 0;
            }

            string s = val.ToString().Trim();
            if (s.Length == 0) return 0;

            long n;
            if (s.Length == 8 && long.TryParse(s, out n) && n >= 19000101 && n <= 29991231)
                return n;

            string[] fmt = { "yyyyMMdd", "dd/MM/yyyy", "d/M/yyyy",
                             "yyyy-MM-dd", "dd-MM-yyyy", "dd.MM.yyyy" };
            DateTime dt2;
            if (DateTime.TryParseExact(s, fmt, CultureInfo.InvariantCulture,
                                       DateTimeStyles.None, out dt2))
                return dt2.Year * 10000L + dt2.Month * 100L + dt2.Day;

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture,
                                  DateTimeStyles.None, out dt2))
                return dt2.Year * 10000L + dt2.Month * 100L + dt2.Day;

            return 0;
        }
    }
}