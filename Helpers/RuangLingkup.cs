using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using CKPNLibrary.Data;
using CKPNLibrary.Models;
using CKPNLibrary.Panel;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Helpers
{
    /// <summary>
    /// Ruang lingkup tunggakan KC1000 (Musyarakah/Mudharabah).
    ///
    /// Ketentuan:
    ///   • Tunggakan pokok (AL) hanya dihitung bila kualitas pembiayaan (AE) = 5 (Macet).
    ///     Kualitas 1–4 → tunggakan pokok dianggap 0.
    ///   • Tunggakan bagi hasil (AN) secara bawaan TIDAK masuk ruang lingkup CKPN, karena pelaporan
    ///     APOLLO hanya memuat piutang pokok Musyarakah (khusus kualitas Macet). Bila di ⚙ Pengaturan
    ///     dicentang "masukkan tunggakan bagi hasil", AN ikut dihitung — juga hanya untuk kualitas Macet (5).
    ///
    /// Pilihan bagi hasil ditetapkan per tahun di database utama (tabel ruang_lingkup_tahun) dan
    /// ditulis panel ke Master!H11 (kolom tersembunyi) sebelum perhitungan:
    ///   kosong  = bawaan (bagi hasil tidak masuk)
    ///   "basil" = bagi hasil masuk ruang lingkup
    ///
    /// Modul cukup memanggil <see cref="TerapkanKC1000"/> sesudah membaca larik tunggakan KC1000,
    /// sehingga CKPN Individu, PD Net Flow, PD Migration, dan bucket kolektif tahunan memakai
    /// ruang lingkup yang sama:
    ///
    ///   RuangLingkup.TerapkanKC1000(spec, ws, dataStart, lastRow, nomPokok, nomBasil);
    /// </summary>
    internal static class RuangLingkup
    {
        public const string Sel = "H11", SelLabel = "G11";
        public const string IsiBasil = "basil";
        public const string LabelSel = "Ruang lingkup KC1000 (diisi panel)";

        // =============================================================
        // Dipanggil modul perhitungan
        // =============================================================

        /// <summary>
        /// Terapkan ruang lingkup KC1000 pada larik nominal tunggakan (diubah di tempat).
        /// Tidak berbuat apa pun untuk sheet selain mode Tunggakan2 (KC1000).
        /// </summary>
        /// <param name="nomPokok">tunggakan pokok (ColNomPokok, AL) — dinolkan bila kualitas ≠ 5</param>
        /// <param name="nomBasil">tunggakan bagi hasil (ColNomBasil, AN) — dinolkan bila tidak masuk lingkup atau kualitas ≠ 5</param>
        public static void TerapkanKC1000(SheetSpec spec, Excel.Worksheet ws, int dataStart, int lastRow,
                                          double[] nomPokok, double[] nomBasil)
        {
            if (spec == null || spec.Mode != SheetMode.Tunggakan2) return;
            TerapkanKC1000(spec, ws, dataStart, lastRow, nomPokok, nomBasil, BasilKC1000());
        }

        /// <summary>Sama dengan di atas, dengan pilihan bagi hasil ditentukan pemanggil.</summary>
        public static void TerapkanKC1000(SheetSpec spec, Excel.Worksheet ws, int dataStart, int lastRow,
                                          double[] nomPokok, double[] nomBasil, bool basil)
        {
            if (spec == null || spec.Mode != SheetMode.Tunggakan2 || ws == null || lastRow < dataStart) return;
            // Pokok dan (bila dicentang) bagi hasil sama-sama hanya untuk kualitas Macet (5)
            int[] kual = BacaKualitas(ws, spec.ColKualitas, dataStart, lastRow);
            if (nomPokok != null)
                for (int i = 0; i < nomPokok.Length; i++)
                    if (i >= kual.Length || kual[i] != 5) nomPokok[i] = 0;
            if (nomBasil != null)
                for (int i = 0; i < nomBasil.Length; i++)
                    if (!basil || i >= kual.Length || kual[i] != 5) nomBasil[i] = 0;
        }

        /// <summary>True bila tunggakan bagi hasil KC1000 masuk ruang lingkup (Master!H11 workbook aplikasi).</summary>
        public static bool BasilKC1000()
        {
            try
            {
                var app = (Excel.Application)ExcelDnaUtil.Application;
                Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app) ?? app.ActiveWorkbook;
                return BasilKC1000(wb);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca ruang lingkup KC1000 Master!" + Sel + " gagal: " + ex.Message);
                return false;
            }
        }

        public static bool BasilKC1000(Excel.Workbook wb)
        {
            Excel.Worksheet m = wb == null ? null : ParameterMaster.CariSheet(wb, "Master");
            return Urai(BacaMentah(m));
        }

        public static bool Urai(string isi)
        {
            return !string.IsNullOrEmpty(isi) && isi.Trim().StartsWith(IsiBasil, StringComparison.OrdinalIgnoreCase);
        }

        public static string BacaMentah(Excel.Worksheet master)
        {
            if (master == null) return "";
            try
            {
                object v = ((Excel.Range)master.Range[Sel]).Value2;
                return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            }
            catch { return ""; }
        }

        /// <summary>Kualitas 1–5 per baris (angka atau teks berawalan digit, mis. "5 - Macet"); lainnya 0.</summary>
        private static int[] BacaKualitas(Excel.Worksheet ws, string kolom, int awal, int akhir)
        {
            int n = akhir - awal + 1;
            var hasil = new int[n];
            if (string.IsNullOrEmpty(kolom) || n <= 0) return hasil;
            object v = ((Excel.Range)ws.Range[kolom + awal, kolom + akhir]).Value2;
            var arr = v as object[,];
            if (arr == null) { hasil[0] = NormKualitas(v); return hasil; }
            int lb = arr.GetLowerBound(0), lc = arr.GetLowerBound(1);
            for (int i = 0; i < n && i < arr.GetLength(0); i++) hasil[i] = NormKualitas(arr[lb + i, lc]);
            return hasil;
        }

        private static int NormKualitas(object o)
        {
            if (o == null) return 0;
            if (o is double) { double d = (double)o; return d >= 1 && d <= 5 && Math.Abs(d - Math.Round(d)) < 1e-9 ? (int)Math.Round(d) : 0; }
            string s = Convert.ToString(o, CultureInfo.InvariantCulture).Trim();
            int j = 0;
            while (j < s.Length && char.IsDigit(s[j])) j++;
            int k;
            return j > 0 && int.TryParse(s.Substring(0, j), out k) && k >= 1 && k <= 5 ? k : 0;
        }

        // =============================================================
        // Label
        // =============================================================
        public static string Label(bool basil)
        {
            return basil ? "Pokok + bagi hasil, kualitas Macet" : "Pokok kualitas Macet saja";
        }

        public static string LabelPanjang(bool basil)
        {
            return basil
                ? "Tunggakan pokok dan bagi hasil hanya kualitas Macet (5)"
                : "Tunggakan pokok hanya kualitas Macet (5); tunggakan bagi hasil tidak dihitung";
        }

        /// <summary>Ketentuan yang berlaku untuk tahun Master!C4 (untuk log).</summary>
        public static string LabelTahunMaster(Excel.Workbook wb)
        {
            try
            {
                DateTime tgl;
                if (wb != null && ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return LabelPanjang(BasilTahun(tgl.Year));
            }
            catch { }
            return LabelPanjang(false);
        }

        // =============================================================
        // Database: pilihan per tahun (tabel bersama di ckpn.db)
        // =============================================================
        internal class Nilai
        {
            public int Tahun;
            public bool Basil;
            public string Dasar, Pengguna, Waktu;
        }

        /// <summary>Pilihan bagi hasil yang berlaku untuk tahun ini (belum diatur = tidak masuk).</summary>
        public static bool BasilTahun(int tahun)
        {
            var n = Baca(tahun);
            return n != null && n.Basil;
        }

        public static Nilai Baca(int tahun)
        {
            if (!Database.AdaBersama) return null;
            try
            {
                using (var con = Database.BukaBersama(false)) return Baca(con, tahun);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca ruang lingkup KC1000 " + tahun + " gagal: " + ex.Message);
                return null;
            }
        }

        private static Nilai Baca(SQLiteConnection con, int tahun)
        {
            using (var cmd = Database.Cmd(con,
                "SELECT kc1000_basil, dasar, pengguna, waktu FROM ruang_lingkup_tahun WHERE tahun=@p0", tahun))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) return null;
                return new Nilai
                {
                    Tahun = tahun,
                    Basil = Convert.ToInt32(rd[0]) != 0,
                    Dasar = rd[1] is DBNull ? "" : rd.GetString(1),
                    Pengguna = rd[2] is DBNull ? "" : rd.GetString(2),
                    Waktu = rd[3] is DBNull ? "" : rd.GetString(3)
                };
            }
        }

        /// <summary>Ada periode Final di tahun ini pada mode perhitungan mana pun.</summary>
        private static bool AdaFinal(int tahun)
        {
            foreach (var m in ModeHitung.Semua)
            {
                if (!Database.AdaMode(m)) continue;
                try
                {
                    using (var con = Database.BukaMode(m, false))
                        if (Convert.ToInt32(Database.Scalar(con,
                            "SELECT COUNT(*) FROM periode WHERE status='Final' AND substr(tanggal,1,4)=@p0", tahun.ToString())) > 0) return true;
                }
                catch (Exception ex) { CatatanLog.Tulis("Cek periode Final (" + m.Id + "): " + ex.Message); }
            }
            return false;
        }

        public static Dictionary<string, object> Simpan(int tahun, bool basil, string dasar)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (tahun < 2000 || tahun > 2100) throw new InvalidOperationException("Tahun tidak valid.");
            if (string.IsNullOrWhiteSpace(dasar)) throw new InvalidOperationException("Dasar penetapan (kebijakan / memo) wajib diisi.");

            bool adaFinal = AdaFinal(tahun);
            Database.Cadangkan();
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                var lama = Baca(con, tahun);
                bool basilLama = lama != null && lama.Basil;
                if (basilLama != basil && adaFinal)
                    throw new InvalidOperationException("Ruang lingkup KC1000 tahun " + tahun + " tidak dapat diubah karena sudah ada periode Final di tahun itu (" +
                        Label(basilLama) + "). Buka kunci periode tersebut terlebih dahulu bila perubahan memang disetujui.");

                Database.Exec(con,
                    "INSERT OR REPLACE INTO ruang_lingkup_tahun(tahun,kc1000_basil,dasar,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4)",
                    tahun, basil ? 1 : 0, dasar.Trim(), Environment.UserName, Database.Sekarang());
                Database.CatatAktivitas(con, tahun.ToString(), "ruang-lingkup-kc1000",
                    (basilLama == basil ? "Ditetapkan: " : "Diubah: " + Label(basilLama) + " → ") + Label(basil) + " · dasar: " + dasar.Trim());
                tx.Commit();
            }
            LogProses.CatatPanel(tahun + "-12-31", "Ruang lingkup KC1000", "Tahun " + tahun, LogProses.OK, LogProses.R()
                .Tambah("Ketentuan", LabelPanjang(basil)).Tambah("Dasar", dasar.Trim()));
            return Info(tahun, null);
        }

        /// <param name="app">bila tidak null, isi Master!H11 dan tahun Master ikut dibaca</param>
        public static Dictionary<string, object> Info(int tahun, Excel.Application app)
        {
            var h = new Dictionary<string, object> { { "tahun", tahun } };
            Nilai n = null;
            bool adaFinal = false;
            if (Database.AdaBersama)
            {
                try
                {
                    using (var con = Database.BukaBersama(false)) n = Baca(con, tahun);
                    adaFinal = AdaFinal(tahun);
                }
                catch (Exception ex) { h["error"] = ex.Message; }
            }
            bool basil = n != null && n.Basil;
            h["basil"] = basil;
            h["label"] = Label(basil);
            h["labelPanjang"] = LabelPanjang(basil);
            h["tersimpan"] = n == null ? null : new Dictionary<string, object>
            {
                { "basil", n.Basil }, { "dasar", n.Dasar }, { "pengguna", n.Pengguna }, { "waktu", n.Waktu }
            };
            h["terkunci"] = adaFinal;
            string info;
            h["bolehMenulis"] = Database.BolehMenulis(out info);
            if (app != null)
            {
                try
                {
                    Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
                    if (wb != null)
                    {
                        h["basilMaster"] = BasilKC1000(wb);
                        DateTime tgl;
                        if (ParameterMaster.BacaTanggalLaporan(wb, out tgl)) h["tahunMaster"] = tgl.Year;
                    }
                }
                catch (Exception ex) { h["errorSheet"] = ex.Message; }
            }
            return h;
        }

        // =============================================================
        // Excel: tulis ke Master!H11
        // =============================================================

        /// <summary>
        /// Tulis pilihan tahun Master!C4 ke Master!H11. Proteksi dibuka hanya bila isi berubah.
        /// Mengembalikan keterangan perubahan (null = tidak ada perubahan).
        /// </summary>
        public static string Terapkan(Excel.Workbook wb)
        {
            if (wb == null) return null;
            DateTime tgl;
            if (!ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return null;
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            if (master == null) throw new InvalidOperationException("Sheet 'Master' tidak ditemukan.");
            bool basil = BasilTahun(tgl.Year);
            string isi = basil ? IsiBasil : "";
            if (BacaMentah(master) == isi) return null;

            Protection.UbahMaster(wb, ms =>
            {
                Excel.Range r = (Excel.Range)ms.Range[Sel];
                r.NumberFormat = "@";
                r.Value2 = isi.Length == 0 ? null : isi;
                try
                {
                    Excel.Range lbl = (Excel.Range)ms.Range[SelLabel];
                    if (lbl.Value2 == null) lbl.Value2 = LabelSel;
                }
                catch { }
            });
            string ket = "Ruang lingkup KC1000 tahun " + tgl.Year + ": " + LabelPanjang(basil);
            CatatanLog.Tulis("  " + ket);
            LogProses.Catat("Ruang lingkup KC1000", "Master!" + Sel, LogProses.Info, LogProses.R()
                .Tambah("Tahun", tgl.Year).Tambah("Ketentuan", LabelPanjang(basil)));
            return ket;
        }

        /// <summary>
        /// Penolakan Simpan grup bila hasil di workbook dihitung dengan ruang lingkup yang berbeda dari
        /// ketetapan tahun ini (mis. Master!H11 diubah manual). Null = sesuai.
        /// </summary>
        public static string CekBeda(Excel.Workbook wb, int tahun)
        {
            try
            {
                bool master = BasilKC1000(wb), tetap = BasilTahun(tahun);
                if (master == tetap) return null;
                return "Ruang lingkup KC1000 di workbook (" + Label(master) + ") berbeda dengan ketetapan tahun " + tahun +
                       " (" + Label(tetap) + "). Hasil tidak dapat disimpan sebagai versi grup. Hitung ulang grup dari panel " +
                       "(ruang lingkup otomatis disesuaikan).";
            }
            catch { return null; }
        }
    }
}
