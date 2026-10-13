using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Lisensi &amp; edisi aplikasi (Tahap 7).
    ///
    ///   TRIAL : tanpa file lisensi, sampai <see cref="TrialSampai"/> — semua fitur.
    ///   PRO   : file lisensi sah (ditandatangani penerbit) — semua fitur, sampai tanggal di lisensi
    ///           (kosong = tanpa batas waktu).
    ///   LITE  : trial / lisensi berakhir — hitung manual untuk SELURUH segmen (KC0600–KC1100),
    ///           Ringkasan &amp; overview, dokumen terbatas (lihat &amp; cetak). Tanpa susunan grup / hitung
    ///           per grup, simpan ke database, penetapan periode, Analisis, dan penyusunan memo.
    ///
    /// File lisensi (*.lic, JSON) diletakkan di folder library (atau library\config):
    ///   { "format": "CKPN414-LISENSI-1", "data": base64(JSON isi), "ttd": base64(tanda tangan) }
    /// Tanda tangan = RSA-2048 PKCS#1 v1.5 + SHA-256 atas byte ASCII string "data".
    /// Hanya kunci PUBLIK yang ada di add-in; kunci privat disimpan penerbit (alat-lisensi\*.ps1),
    /// sehingga isi lisensi (nama pemegang, tanggal) tidak dapat diubah tanpa membatalkan tanda tangan.
    ///
    /// Catatan keamanan: ini mencegah penyalinan / pengubahan lisensi oleh pengguna biasa, bukan
    /// pembongkaran .xll (kode .NET dapat didekompilasi). Nama pemegang lisensi dicetak di setiap dokumen.
    /// </summary>
    internal static class Lisensi
    {
        public const string NamaProduk = "CKPN PSAK 414";
        public const string Format = "CKPN414-LISENSI-1";
        public const string Pengembang = "Rukhan · PT BPR Syariah Barokah Dana Sejahtera";

        /// <summary>Akhir masa trial tanpa file lisensi (inklusif).</summary>
        public static readonly DateTime TrialSampai = new DateTime(2026, 12, 31);

        /// <summary>
        /// Kunci publik penerbit (output alat-lisensi\BuatKunci.ps1 → kunci_publik.xml).
        /// GANTI dengan kunci Anda sendiri, lalu build. Selama masih placeholder, tidak ada file lisensi yang sah.
        /// </summary>
        private const string KunciPublikXml =
            "<RSAKeyValue><Modulus>0if2BANYjeiAp3BtY2LXqIPKnIrsmgJDCis40BCfFm8blVdUsBLGIA/AE+B3ADuE8cgyvfICRfq0o5v4pBganNpbk5RZdfTctq0juofDkRXngT1ozpaa3Sl6WlzUf+7dH4o3+LHy0qbyDr8H1AFtLz+MSgNpbJM7hHpqph7mC5LhtSs0XsQ0M2jXZd5aq9/KAqI4UfvfzurOIS7BsU+ZFU2WIipTtU9BRGkDBXkmWh1NiCNA1rJzQjG/Tlsngmfi1TkXLeazf0WZTMxFZcEvnhcEl1RHLF5mNTnauXU+L1xk+fXdPmPTT61rYYX+8sRIeCsC/2mGtfZ9bmbvA/iH2Q==</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>";

        public const string Trial = "trial", Pro = "pro", Lite = "lite";

        // Fitur yang TIDAK tersedia di Lite (dipakai PanelBridge &amp; panel)
        public const string FSegmentasi = "segmentasi", FSimpan = "simpan", FTetapkan = "tetapkan",
                            FAnalisis = "analisis", FMemo = "memo";
        private static readonly Dictionary<string, string> NamaFitur = new Dictionary<string, string>
        {
            { FSegmentasi, "Susunan grup segmen & hitung per grup" },
            { FSimpan, "Simpan hasil ke database (versi, koreksi)" },
            { FTetapkan, "Konsolidasi & penetapan periode" },
            { FAnalisis, "Tab Analisis (tren, rincian, komposisi, versi, log)" },
            { FMemo, "Penyusunan memo Individu & LGD CS" },
        };

        /// <summary>Perintah panel → fitur yang dibutuhkan.</summary>
        private static readonly Dictionary<string, string> FiturPerintah = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "hitungGrup", FSegmentasi }, { "hitungSemuaGrup", FSegmentasi }, { "simpanSusunan", FSegmentasi },
            { "simpanGrup", FSimpan }, { "hapusVersi", FSimpan }, { "koreksiSimpan", FSimpan },
            { "konsolidasi", FTetapkan }, { "tetapkanPeriode", FTetapkan }, { "bukaKunci", FTetapkan },
            { "analisisSumber", FAnalisis }, { "analisisData", FAnalisis }, { "riwayat", FAnalisis },
            { "rincianCkpn", FAnalisis }, { "penjelasanCKPN", FAnalisis }, { "koreksiData", FAnalisis },
            { "simpanMemo", FMemo }, { "bukaMemo", FMemo }, { "hapusMemo", FMemo }, { "ambilProfilTemplate", FMemo },
        };

        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer();
        private static readonly string[] SemuaKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        internal class Info
        {
            public string Edisi = Trial;
            public string Pemegang, Alamat, Id, Catatan, File, Masalah, DiterbitkanOleh;
            public DateTime? Diterbitkan, BerlakuSampai;
            public bool LisensiBerakhir;
            public DateTime Hari;
            public DateTime WaktuCek;
            public string KunciCache;
        }

        private static Info _cache;

        // =================================================================
        // Status
        // =================================================================
        public static Info Status()
        {
            DateTime hari = HariEfektif();
            var files = DaftarFile();
            var kunci = new StringBuilder(hari.ToString("yyyyMMdd"));
            foreach (var f in files) kunci.Append('|').Append(f).Append('@').Append(Cap(f));
            if (_cache != null && _cache.KunciCache == kunci.ToString() && (DateTime.Now - _cache.WaktuCek).TotalMinutes < 10) return _cache;

            var info = new Info { Hari = hari, WaktuCek = DateTime.Now, KunciCache = kunci.ToString() };
            Dictionary<string, object> terbaik = null;
            string fileTerbaik = null, masalah = null;
            foreach (var f in files)
            {
                string m;
                var isi = BacaFile(f, out m);
                if (isi == null) { masalah = Path.GetFileName(f) + ": " + m; continue; }
                if (terbaik == null || LebihBaik(isi, terbaik)) { terbaik = isi; fileTerbaik = f; }
            }

            if (terbaik != null)
            {
                info.Pemegang = S(terbaik, "pemegang");
                info.Alamat = S(terbaik, "alamat");
                info.Id = S(terbaik, "id");
                info.Catatan = S(terbaik, "catatan");
                info.DiterbitkanOleh = S(terbaik, "penerbit");
                info.Diterbitkan = Tgl(S(terbaik, "diterbitkan"));
                info.BerlakuSampai = Tgl(S(terbaik, "berlakuSampai"));
                info.File = fileTerbaik;
                info.LisensiBerakhir = info.BerlakuSampai.HasValue && hari > info.BerlakuSampai.Value;
                if (!info.LisensiBerakhir) info.Edisi = Pro;
            }
            else info.Masalah = masalah;

            if (info.Edisi != Pro) info.Edisi = hari <= TrialSampai ? Trial : Lite;
            _cache = info;
            return info;
        }

        /// <summary>Pesan penolakan bila fitur tidak tersedia di edisi saat ini; null = boleh.</summary>
        public static string Tolak(string fitur)
        {
            if (fitur == null || Status().Edisi != Lite) return null;
            string nama;
            NamaFitur.TryGetValue(fitur, out nama);
            return (nama ?? "Fitur ini") + " tidak tersedia di edisi Lite. " + AlasanLite() +
                   " Pasang file lisensi Pro di ⚙ Pengaturan › Tentang & lisensi.";
        }

        public static string TolakPerintah(string cmd)
        {
            string f;
            return FiturPerintah.TryGetValue(cmd ?? "", out f) ? Tolak(f) : null;
        }

        /// <summary>Lite: perhitungan wajib mencakup seluruh KC0600–KC1100 (tanpa segmentasi).</summary>
        public static string TolakHitung(IEnumerable<string> kcDicentang)
        {
            if (Status().Edisi != Lite) return null;
            var ada = new HashSet<string>(kcDicentang ?? new string[0], StringComparer.OrdinalIgnoreCase);
            var kurang = new List<string>();
            foreach (var k in SemuaKC) if (!ada.Contains(k)) kurang.Add(k);
            if (kurang.Count == 0) return null;
            return "Edisi Lite hanya menghitung CKPN untuk seluruh segmen sekaligus. Centang semua KC0600–KC1100 di sheet Master " +
                   "(belum dicentang: " + string.Join(", ", kurang.ToArray()) + "). " + AlasanLite();
        }

        /// <summary>
        /// Lite: daftar KC yang benar-benar dihitung = seluruh KC0600–KC1100, apa pun centang di Master.
        /// Dipakai CKPNFunctions (tombol VBA) untuk menimpa parameter sheetKCList. Trial/Pro: tidak diubah.
        /// </summary>
        public static string KCEfektif(string sheetKCList)
        {
            try { return Status().Edisi == Lite ? string.Join(",", SemuaKC) : sheetKCList; }
            catch { return sheetKCList; }
        }

        /// <summary>
        /// Lite: centang seluruh KC0600–KC1100 di sheet Master bila belum, supaya isi Master sesuai dengan
        /// yang dihitung. Return pesan penjelasan bila centang diubah; null bila tidak ada perubahan / bukan Lite.
        /// </summary>
        public static string CentangSeluruhKC(Excel.Workbook wb)
        {
            if (wb == null || Status().Edisi != Lite) return null;
            Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
            if (m == null) return null;
            var lama = ParameterMaster.BacaKCDicentang(m);
            if (lama.Count == SemuaKC.Length) return null;
            ParameterMaster.SetKCDicentang(wb, SemuaKC);
            string pesan = "Edisi Lite menghitung CKPN untuk seluruh segmen sekaligus. Centang KC di Master diubah dari " +
                           (lama.Count == 0 ? "(kosong)" : string.Join(", ", lama.ToArray())) + " menjadi KC0600–KC1100. " + AlasanLite();
            LogProses.Catat("Lisensi", "Edisi Lite", LogProses.Info, LogProses.R()
                .Tambah("Centang KC sebelum", lama.Count == 0 ? "(kosong)" : string.Join(",", lama.ToArray()))
                .Tambah("Dihitung", string.Join(",", SemuaKC)));
            return pesan;
        }

        /// <summary>
        /// Dipanggil Protection.CekProteksi (tombol VBA di Master). Tidak lagi menolak: pada edisi Lite centang
        /// KC di Master disamakan menjadi seluruh KC0600–KC1100 dan perhitungan tetap jalan menyeluruh
        /// (parameter sheetKCList ditimpa lewat <see cref="KCEfektif"/>). Selalu null.
        /// </summary>
        public static string TolakTombolMaster(Excel.Application app)
        {
            try
            {
                if (Status().Edisi != Lite) return null;
                string pesan = CentangSeluruhKC(PanelBridge.CariWorkbookAplikasi(app) ?? app.ActiveWorkbook);
                if (pesan != null) Pemberitahu.Info("Edisi Lite", pesan);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Centang KC edisi Lite: " + ex.Message);
            }
            return null;
        }

        private static string AlasanLite()
        {
            var s = Status();
            return s.LisensiBerakhir && s.BerlakuSampai.HasValue
                ? "Lisensi berlaku sampai " + s.BerlakuSampai.Value.ToString("d MMMM yyyy", LogProses.Id) + "."
                : "Masa trial berakhir " + TrialSampai.ToString("d MMMM yyyy", LogProses.Id) + ".";
        }

        // =================================================================
        // Panel
        // =================================================================
        public static Dictionary<string, object> InfoPanel()
        {
            var s = Status();
            var fiturLite = new List<object>();
            foreach (var kv in NamaFitur) fiturLite.Add(new Dictionary<string, object> { { "id", kv.Key }, { "nama", kv.Value } });
            var izin = new Dictionary<string, object>();
            foreach (var k in NamaFitur.Keys) izin[k] = s.Edisi != Lite;
            return new Dictionary<string, object>
            {
                { "produk", NamaProduk }, { "pengembang", Pengembang }, { "edisi", s.Edisi },
                { "pemegang", s.Pemegang }, { "alamat", s.Alamat }, { "id", s.Id }, { "catatan", s.Catatan },
                { "penerbit", s.DiterbitkanOleh },
                { "diterbitkan", s.Diterbitkan.HasValue ? s.Diterbitkan.Value.ToString("yyyy-MM-dd") : null },
                { "berlakuSampai", s.BerlakuSampai.HasValue ? s.BerlakuSampai.Value.ToString("yyyy-MM-dd") : null },
                { "lisensiBerakhir", s.LisensiBerakhir }, { "file", s.File }, { "masalah", s.Masalah },
                { "trialSampai", TrialSampai.ToString("yyyy-MM-dd") },
                { "sisaHariTrial", Math.Max(0, (int)(TrialSampai - s.Hari).TotalDays + 1) },
                { "sisaHariLisensi", s.BerlakuSampai.HasValue ? (object)Math.Max(0, (int)(s.BerlakuSampai.Value - s.Hari).TotalDays + 1) : null },
                { "hari", s.Hari.ToString("yyyy-MM-dd") },
                { "kunciTerpasang", !KunciPublikXml.Contains("GANTI") },
                { "folder", AppPaths.FolderLibrary },
                { "izin", izin }, { "fiturLite", fiturLite }
            };
        }

        /// <summary>Validasi lalu salin file lisensi ke library\lisensi.lic.</summary>
        public static Dictionary<string, object> Pasang(string sumber)
        {
            if (string.IsNullOrEmpty(sumber) || !System.IO.File.Exists(sumber)) throw new FileNotFoundException("File lisensi tidak ditemukan: " + sumber);
            string masalah;
            var isi = BacaFile(sumber, out masalah);
            if (isi == null) throw new InvalidOperationException("File lisensi tidak sah: " + masalah);
            string tujuan = Path.Combine(AppPaths.FolderLibrary, "lisensi.lic");
            if (!string.Equals(Path.GetFullPath(sumber), Path.GetFullPath(tujuan), StringComparison.OrdinalIgnoreCase))
            {
                if (System.IO.File.Exists(tujuan))
                    System.IO.File.Copy(tujuan, Path.Combine(AppPaths.FolderLibrary, "lisensi_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".lic.lama"), true);
                System.IO.File.Copy(sumber, tujuan, true);
            }
            _cache = null;
            LogProses.CatatPanel(DateTime.Today.ToString("yyyy-MM-dd"), "Lisensi", "Pasang lisensi", LogProses.OK, LogProses.R()
                .Tambah("Pemegang", S(isi, "pemegang")).Tambah("ID", S(isi, "id"))
                .Tambah("Berlaku sampai", string.IsNullOrEmpty(S(isi, "berlakuSampai")) ? "tanpa batas" : S(isi, "berlakuSampai")));
            return InfoPanel();
        }

        // =================================================================
        // Verifikasi file
        // =================================================================
        private static List<string> DaftarFile()
        {
            var hasil = new List<string>();
            foreach (var folder in new[] { AppPaths.FolderLibrary, AppPaths.FolderConfig })
            {
                try
                {
                    if (!string.IsNullOrEmpty(folder) && Directory.Exists(folder))
                        foreach (var f in Directory.GetFiles(folder, "*.lic")) hasil.Add(f);
                }
                catch { }
            }
            hasil.Sort(StringComparer.OrdinalIgnoreCase);
            return hasil;
        }

        /// <summary>Isi lisensi bila tanda tangan &amp; format sah; null + masalah bila tidak.</summary>
        internal static Dictionary<string, object> BacaFile(string path, out string masalah)
        {
            masalah = null;
            try
            {
                var luar = _json.DeserializeObject(System.IO.File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                if (luar == null || S(luar, "format") != Format) { masalah = "format tidak dikenal"; return null; }
                string data = S(luar, "data"), ttd = S(luar, "ttd");
                if (string.IsNullOrEmpty(data) || string.IsNullOrEmpty(ttd)) { masalah = "data / tanda tangan kosong"; return null; }
                if (KunciPublikXml.Contains("GANTI")) { masalah = "kunci publik penerbit belum dipasang di add-in (Lisensi.cs)"; return null; }
                if (!Verifikasi(Encoding.ASCII.GetBytes(data), Convert.FromBase64String(ttd))) { masalah = "tanda tangan tidak cocok (file diubah atau bukan dari penerbit)"; return null; }
                var isi = _json.DeserializeObject(Encoding.UTF8.GetString(Convert.FromBase64String(data))) as Dictionary<string, object>;
                if (isi == null || S(isi, "produk") != NamaProduk) { masalah = "lisensi untuk produk lain"; return null; }
                if (string.IsNullOrEmpty(S(isi, "pemegang"))) { masalah = "nama pemegang kosong"; return null; }
                if (!string.IsNullOrEmpty(S(isi, "berlakuSampai")) && !Tgl(S(isi, "berlakuSampai")).HasValue) { masalah = "tanggal berlaku tidak valid"; return null; }
                return isi;
            }
            catch (Exception ex)
            {
                masalah = "tidak dapat dibaca (" + ex.Message + ")";
                return null;
            }
        }

        private static bool Verifikasi(byte[] data, byte[] ttd)
        {
            var p = new RSAParameters
            {
                Modulus = Convert.FromBase64String(Tag(KunciPublikXml, "Modulus")),
                Exponent = Convert.FromBase64String(Tag(KunciPublikXml, "Exponent"))
            };
            using (var rsa = new RSACng())
            {
                rsa.ImportParameters(p);
                return rsa.VerifyData(data, ttd, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
            }
        }

        private static string Tag(string xml, string tag)
        {
            var m = Regex.Match(xml, "<" + tag + ">([^<]*)</" + tag + ">");
            return m.Success ? m.Groups[1].Value.Trim() : "";
        }

        /// <summary>Tanpa batas waktu lebih baik; lalu tanggal berlaku paling akhir.</summary>
        private static bool LebihBaik(Dictionary<string, object> a, Dictionary<string, object> b)
        {
            DateTime? ta = Tgl(S(a, "berlakuSampai")), tb = Tgl(S(b, "berlakuSampai"));
            if (!ta.HasValue) return tb.HasValue;
            if (!tb.HasValue) return false;
            return ta.Value > tb.Value;
        }

        // =================================================================
        // Tanggal efektif — penangkal memundurkan jam komputer
        // =================================================================
        /// <summary>
        /// max(hari ini, tanggal tertinggi yang pernah tercatat di registry user ini, tanggal ubah database).
        /// Jam yang dimundurkan tidak memperpanjang trial / lisensi.
        /// </summary>
        private static DateTime HariEfektif()
        {
            DateTime hari = DateTime.Today;
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(@"Software\CKPN414"))
                {
                    DateTime t;
                    string v = Convert.ToString(k.GetValue("Jejak") ?? "");
                    if (DateTime.TryParseExact(v, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out t) && t > hari && t < hari.AddYears(5)) hari = t;
                    else if (hari > t) k.SetValue("Jejak", hari.ToString("yyyyMMdd"));
                }
            }
            catch { }
            try
            {
                var fi = new FileInfo(AppPaths.FileDatabase);
                if (fi.Exists && fi.LastWriteTime.Date > hari && fi.LastWriteTime.Date < DateTime.Today.AddYears(5)) hari = fi.LastWriteTime.Date;
            }
            catch { }
            return hari;
        }

        // =================================================================
        // Util
        // =================================================================
        private static string S(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) && v != null ? Convert.ToString(v).Trim() : "";
        }

        private static DateTime? Tgl(string s)
        {
            DateTime t;
            if (string.IsNullOrEmpty(s)) return null;
            return DateTime.TryParseExact(s, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture,
                                          System.Globalization.DateTimeStyles.None, out t) ? (DateTime?)t : null;
        }

        private static string Cap(string f)
        {
            try { var fi = new FileInfo(f); return fi.Length + ":" + fi.LastWriteTimeUtc.Ticks; } catch { return ""; }
        }
    }
}
