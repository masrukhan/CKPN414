using System;
using System.Collections.Generic;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Pembacaan data hapus buku (sheet KC2900, Form 29 APOLLO) untuk PD Net Flow dan LGD ER.
    ///
    /// Kolom KC2900 (data mulai baris 3):
    ///   H = Nomor Rekening, I = Tanggal Hapus Buku, J = Jumlah (nilai saat dihapus buku),
    ///   K = Jumlah Dipulihkan, L = Baki Debet (sisa).
    ///
    /// 1. PD Net Flow — sesuai PAPSI BPRS, baris WO diisi pembiayaan yang DIHAPUS BUKU PADA BULAN
    ///    posisi tersebut (bukan seluruh saldo KC2900). Rekening dipilih dari kolom I (tanggal hapus
    ///    buku di bulan posisi), nominal = kolom J (nilai hapus buku); bila J kosong dipakai L.
    ///
    /// 2. LGD ER — file tahunan yang kosong / sama dengan file tahun lain dianggap belum tersedia,
    ///    sehingga kohort lama memakai nilai hapus buku (kolom J) dari file pertama yang ada.
    ///
    /// Format tanggal APOLLO kolom I = teks yyyymmdd (mis. "20211130"); tanggal Excel dan teks
    /// dd-mm-yyyy juga diterima sebagai cadangan.
    /// </summary>
    internal static class HapusBuku
    {
        public const string Sheet = "KC2900";
        public const int BarisAwal = 3;

        private static readonly CultureInfo Id = new CultureInfo("id-ID");

        // =============================================================
        // Tanggal
        // =============================================================

        /// <summary>Nilai sel → yyyyMMdd (0 bila tidak terbaca). Menerima tanggal Excel, angka yyyymmdd,
        /// dan teks dd-MM-yyyy / d-M-yyyy / dd/MM/yyyy / yyyy-MM-dd / dd.MM.yyyy.</summary>
        public static long KeYyyymmdd(object v)
        {
            if (v == null || v is DBNull) return 0;
            if (v is double || v is int || v is long || v is decimal || v is float)
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (d >= 19000101 && d <= 29991231) return (long)d;
                if (d > 0 && d < 200000)
                {
                    try { DateTime t = DateTime.FromOADate(d); return t.Year * 10000L + t.Month * 100L + t.Day; }
                    catch { return 0; }
                }
                return 0;
            }
            string s = Convert.ToString(v, CultureInfo.InvariantCulture).Trim().Replace(' ', ' ').Trim();
            if (s.Length == 0) return 0;
            long n;
            if (s.Length == 8 && long.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out n) && n >= 19000101 && n <= 29991231)
                return n;
            string[] fmt =
            {
                "dd-MM-yyyy", "d-M-yyyy", "dd/MM/yyyy", "d/M/yyyy", "dd.MM.yyyy", "d.M.yyyy",
                "yyyy-MM-dd", "yyyy/MM/dd", "yyyyMMdd",
                "dd-MM-yyyy HH:mm:ss", "dd/MM/yyyy HH:mm:ss", "yyyy-MM-dd HH:mm:ss"
            };
            DateTime dt;
            if (DateTime.TryParseExact(s, fmt, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return dt.Year * 10000L + dt.Month * 100L + dt.Day;
            if (DateTime.TryParse(s, Id, DateTimeStyles.None, out dt))   // mis. "27 Oktober 2016"
                return dt.Year * 10000L + dt.Month * 100L + dt.Day;
            return 0;
        }

        /// <summary>Kolom tanggal sebagai yyyyMMdd (double, agar bisa menggantikan BacaKolomDouble).</summary>
        public static double[] BacaTanggal(Excel.Worksheet ws, string kolom, int awal, int akhir)
        {
            object[] mentah = BacaMentah(ws, kolom, awal, akhir);
            var hasil = new double[mentah.Length];
            for (int i = 0; i < mentah.Length; i++) hasil[i] = KeYyyymmdd(mentah[i]);
            return hasil;
        }

        private static object[] BacaMentah(Excel.Worksheet ws, string kolom, int awal, int akhir)
        {
            int n = akhir - awal + 1;
            if (n <= 0) return new object[0];
            var hasil = new object[n];
            object v = ((Excel.Range)ws.Range[kolom + awal, kolom + akhir]).Value2;
            var arr = v as object[,];
            if (arr == null) { hasil[0] = v; return hasil; }
            int lb = arr.GetLowerBound(0), lc = arr.GetLowerBound(1);
            for (int i = 0; i < n && i < arr.GetLength(0); i++) hasil[i] = arr[lb + i, lc];
            return hasil;
        }

        private static double Angka(object o)
        {
            if (o == null || o is DBNull || o is int) return 0;   // int = kode error formula di interop
            if (o is double) return (double)o;
            double d;
            string s = Convert.ToString(o, CultureInfo.InvariantCulture).Trim().Replace(" ", "");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d)) return d;
            if (double.TryParse(s, NumberStyles.Any, Id, out d)) return d;
            return 0;
        }

        // =============================================================
        // PD Net Flow: hapus buku per bulan posisi
        // =============================================================

        private static readonly string[][] NamaBulan =
        {
            new[] { "januari", "jan" }, new[] { "februari", "feb", "pebruari" }, new[] { "maret", "mar" },
            new[] { "april", "apr" }, new[] { "mei", "may" }, new[] { "juni", "jun" }, new[] { "juli", "jul" },
            new[] { "agustus", "agu", "agt", "aug", "august" }, new[] { "september", "sep", "sept" },
            new[] { "oktober", "okt", "oct", "october" }, new[] { "november", "nov", "nop" },
            new[] { "desember", "des", "dec", "december" }
        };

        /// <summary>
        /// Bulan posisi kolom Net Flow dari label Master!B20:B32 (mis. "Desember 2024") atau tanggal ref
        /// Master!F20:F32. Label didahulukan karena tidak bergantung format tanggal komputer.
        /// </summary>
        public static bool BulanPosisi(string label, string refDate, out int tahun, out int bulan)
        {
            tahun = 0; bulan = 0;
            if (DariLabel(label, out tahun, out bulan)) return true;

            string s = (refDate ?? "").Trim();
            if (s.Length == 0) return false;
            double oa;
            if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out oa) && oa > 20000 && oa < 80000)
            {
                DateTime t = DateTime.FromOADate(oa); tahun = t.Year; bulan = t.Month; return true;
            }
            long ymd = KeYyyymmdd(s);
            if (ymd > 0) { tahun = (int)(ymd / 10000); bulan = (int)(ymd / 100 % 100); return true; }
            DateTime dt;
            // Teks hasil CStr(tanggal) dari Master → format tanggal komputer ini
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt)) { tahun = dt.Year; bulan = dt.Month; return true; }
            return false;
        }

        private static bool DariLabel(string label, out int tahun, out int bulan)
        {
            tahun = 0; bulan = 0;
            string s = (label ?? "").Trim().ToLowerInvariant();
            if (s.Length == 0) return false;
            var kata = s.Split(new[] { ' ', '-', '/', '.', ',', '\'' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var k in kata)
            {
                int angka;
                if (int.TryParse(k, out angka))
                {
                    if (angka >= 1900 && angka <= 2999) tahun = angka;
                    else if (angka >= 0 && angka <= 99 && tahun == 0 && bulan > 0) tahun = 2000 + angka;   // "Des-24"
                    continue;
                }
                for (int b = 0; b < 12 && bulan == 0; b++)
                    foreach (var nm in NamaBulan[b])
                        if (k == nm) { bulan = b + 1; break; }
            }
            return tahun > 0 && bulan > 0;
        }

        /// <summary>Hasil hapus buku satu bulan posisi.</summary>
        internal class WoBulanan
        {
            public double Nominal;
            public int Rekening;
            public int TanggalTakTerbaca;      // baris bernomor rekening tanpa tanggal hapus buku yang terbaca
            public int PakaiSaldo;             // baris dengan kolom J kosong → dipakai L
        }

        /// <summary>
        /// Jumlah pembiayaan yang dihapus buku pada bulan (tahun, bulan) dari KC2900 file posisi itu.
        /// Nominal = kolom J (nilai hapus buku); bila J kosong/0 dipakai L (baki debet).
        /// </summary>
        public static WoBulanan Bulanan(Excel.Workbook wb, int tahun, int bulan)
        {
            var h = new WoBulanan();
            if (!ExcelHelper.SheetAda(wb, Sheet)) return h;
            var ws = (Excel.Worksheet)wb.Worksheets[Sheet];
            int akhir = ExcelHelper.CariLastRow(ws, "H", BarisAwal - 1);
            if (akhir < BarisAwal) return h;

            object[] rek = BacaMentah(ws, "H", BarisAwal, akhir);
            object[] tgl = BacaMentah(ws, "I", BarisAwal, akhir);
            object[] jml = BacaMentah(ws, "J", BarisAwal, akhir);
            object[] saldo = BacaMentah(ws, "L", BarisAwal, akhir);
            long dari = tahun * 10000L + bulan * 100L + 1, sampai = tahun * 10000L + bulan * 100L + 31;

            for (int i = 0; i < rek.Length; i++)
            {
                if (rek[i] == null || Convert.ToString(rek[i], CultureInfo.InvariantCulture).Trim().Length == 0) continue;
                long t = KeYyyymmdd(tgl[i]);
                if (t == 0) { h.TanggalTakTerbaca++; continue; }
                if (t < dari || t > sampai) continue;
                double nilai = Angka(jml[i]);
                if (nilai <= 0) { nilai = Angka(saldo[i]); if (nilai > 0) h.PakaiSaldo++; }
                h.Nominal += nilai;
                h.Rekening++;
            }
            return h;
        }

        // =============================================================
        // LGD ER: file tahunan yang tersedia
        // =============================================================

        /// <summary>
        /// Saring daftar file tahunan LGD ER (Master!D60:D65):
        ///   • path kosong → tahun itu belum ada data;
        ///   • path yang SAMA dengan tahun lain → hanya dipakai untuk tahun terbarunya
        ///     (mis. Des 2023 diisikan ke slot 2020–2022 → slot 2020–2022 dianggap belum ada data).
        /// Tahun berjalan wajib tersedia, dan tahun yang tersedia harus berurutan sampai tahun berjalan.
        /// Mengembalikan tahun pertama yang punya data; daftar tahun yang dilewati di <paramref name="dilewati"/>.
        /// </summary>
        public static int SaringFileLgdEr(Dictionary<int, string> filePaths, int currYear, out List<int> dilewati)
        {
            dilewati = new List<int>();
            var tahun = new List<int>(filePaths.Keys);
            tahun.Sort();
            tahun.Reverse();                        // terbaru dulu: path ganda dipertahankan di tahun terbaru
            var dipakai = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (int y in tahun)
            {
                string p = (filePaths[y] ?? "").Trim();
                string kunci = p;
                try { if (p.Length > 0) kunci = System.IO.Path.GetFullPath(p); } catch { }
                if (p.Length == 0 || p.IndexOf("pilih", StringComparison.OrdinalIgnoreCase) >= 0 || !dipakai.Add(kunci))
                {
                    filePaths.Remove(y);
                    dilewati.Add(y);
                }
            }
            dilewati.Sort();
            if (!filePaths.ContainsKey(currYear))
                throw new InvalidOperationException("File LGD ER tahun berjalan (" + currYear + ") wajib diisi di Master!D65.");
            int awal = currYear;
            while (filePaths.ContainsKey(awal - 1)) awal--;
            foreach (int y in new List<int>(filePaths.Keys))
                if (y < awal)
                    throw new InvalidOperationException("File LGD ER harus berurutan sampai tahun berjalan. Tahun " + (awal - 1) +
                        " kosong, tetapi tahun " + y + " diisi. Kosongkan tahun sebelum " + awal + " atau lengkapi tahun " + (awal - 1) + ".");
            return awal;
        }

        /// <summary>Catatan untuk sheet B3.LGD-ER bila data tidak mencakup seluruh periode.</summary>
        public static string CatatanLgdEr(int tahunAwalData, int minYear)
        {
            if (tahunAwalData <= minYear) return null;
            return "Data KC2900 tersedia mulai Des " + tahunAwalData + ". Kohort " + minYear +
                   (tahunAwalData - 1 > minYear ? "–" + (tahunAwalData - 1) : "") +
                   ": saldo awal = nilai hapus buku (KC2900 kolom J); pemulihan s.d. Des " + tahunAwalData +
                   " = nilai hapus buku − saldo Des " + tahunAwalData + ".";
        }
    }
}
