using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Text;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Tahap 8 — mode perhitungan CKPN.
    ///
    ///   akad   : grup = kombinasi kode segmen KC0600–KC1100 (cara lama). Database ckpn.db.
    ///   kantor : grup = kumpulan sandi lokasi kantor (kolom B setiap sheet KC). Seluruh KC0600–KC1100
    ///            dihitung, lalu baris disaring per sandi kantor. Database ckpn_kantor.db.
    ///
    /// Aturan:
    ///   - Mode AKTIF dipilih per pengguna di panel (file lokal %LOCALAPPDATA%\CKPN414\mode.txt).
    ///     Setiap mode punya database sendiri: versi, periode, susunan grup, konsolidasi, jurnal, dan
    ///     riwayat tidak saling terlihat. Penyesuaian agunan, memo, parameter ABA, profil template, dan
    ///     referensi KC disimpan di database utama dan dipakai bersama.
    ///   - Mode RESMI ditetapkan per tahun (tabel mode_tahun, database utama). Penetapan Final dan
    ///     jurnal hanya dari mode resmi; mode lain = simulasi. Mode resmi tidak dapat diganti setelah
    ///     ada periode Final di tahun itu (praktis: setahun sekali).
    ///   - Edisi Lite selalu mode akad.
    ///
    /// Lingkup grup untuk modul ditulis panel ke Master!H8 (kolom H tersembunyi), format
    /// "kantor:001,002". Kosong = tanpa saringan. Modul membacanya lewat <see cref="FilterLingkup"/>.
    ///
    /// Mode berikutnya (mis. jenis penggunaan, kolom S) cukup ditambahkan ke <see cref="Semua"/> dengan
    /// KolomFilter-nya sendiri; database, susunan grup, saringan modul, dan panel mengikuti.
    /// </summary>
    internal static class ModeHitung
    {
        internal class Mode
        {
            public string Id, Nama, NamaPendek, FileDb, KolomFilter, NamaUnit, Keterangan;
            public bool Utama;
        }

        public const string Akad = "akad", Kantor = "kantor", Penggunaan = "penggunaan";

        public static readonly Mode[] Semua =
        {
            new Mode
            {
                Id = Akad, Nama = "Per jenis akad (segmen KC)", NamaPendek = "Per akad", FileDb = "ckpn.db", Utama = true,
                KolomFilter = null, NamaUnit = "kode segmen",
                Keterangan = "Grup disusun dari jenis akad / segmen KC0600–KC1100 (murabahah, istishna, ijarah, dst.)."
            },
            new Mode
            {
                Id = Kantor, Nama = "Per sandi lokasi kantor", NamaPendek = "Per kantor", FileDb = "ckpn_kantor.db",
                KolomFilter = "B", NamaUnit = "sandi kantor",
                Keterangan = "Grup disusun dari kantor (sandi kantor di template APOLLO). Semua jenis akad di kantor tersebut dihitung bersama."
            },
            new Mode
            {
                Id = Penggunaan, Nama = "Per jenis penggunaan", NamaPendek = "Per penggunaan", FileDb = "ckpn_penggunaan.db",
                KolomFilter = "S", NamaUnit = "jenis penggunaan",
                Keterangan = "Grup disusun dari jenis penggunaan pembiayaan: modal kerja, investasi, dan konsumsi. Semua jenis akad dihitung bersama."
            },
        };

        /// <summary>Tahap 8c: kode jenis penggunaan (kolom S KC0600–KC1100) dan namanya.</summary>
        public static readonly string[,] JenisPenggunaan = { { "1", "Modal Kerja" }, { "2", "Investasi" }, { "3", "Konsumsi" } };
        /// <summary>Isian kolom S kosong / tidak dikenal pada baris pembiayaan → Konsumsi (keputusan kebijakan).</summary>
        public const string PenggunaanCadangan = "3";

        public static string NamaPenggunaan(string kode)
        {
            for (int i = 0; i < JenisPenggunaan.GetLength(0); i++) if (JenisPenggunaan[i, 0] == kode) return JenisPenggunaan[i, 1];
            return kode;
        }

        /// <summary>
        /// Kode jenis penggunaan kanonik dari isian template: "1", 1.0, "'1'", "01", "1 (Modal Kerja)", "Modal Kerja" → "1".
        /// Kosong / tidak dikenal → "" (pemanggil memutuskan cadangannya).
        /// </summary>
        public static string NormPenggunaan(object v)
        {
            if (v == null || v is DBNull) return "";
            string s = v is double ? ((long)Math.Round((double)v)).ToString(CultureInfo.InvariantCulture)
                                   : Convert.ToString(v, CultureInfo.InvariantCulture).Trim().Trim('\'', '"').Trim();
            if (s.Length == 0) return "";
            string digit = "";
            foreach (char c in s)
            {
                if (c >= '0' && c <= '9') digit += c;
                else if (digit.Length > 0) break;
                else if (!char.IsWhiteSpace(c) && c != '(' && c != '-') break;
            }
            if (digit.Length > 0)
            {
                string d = digit.TrimStart('0');
                return d == "1" || d == "2" || d == "3" ? d : "";
            }
            string u = s.ToUpperInvariant();
            if (u.Contains("MODAL")) return "1";
            if (u.Contains("INVESTASI")) return "2";
            if (u.Contains("KONSUMSI") || u.Contains("KONSUMTIF")) return "3";
            return "";
        }

        /// <summary>Kode unit kanonik sesuai mode (sandi kantor / jenis penggunaan).</summary>
        public static string NormUnit(Mode m, object v)
        {
            if (m != null && m.Id == Penggunaan) return NormPenggunaan(v);
            return NormSandi(v);
        }

        public static readonly string[] SemuaKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        /// <summary>Sel saringan lingkup di Master (kolom H tersembunyi). Diisi panel, dibaca modul.</summary>
        public const string SelFilter = "H8";

        public static Mode Cari(string id)
        {
            foreach (var m in Semua) if (string.Equals(m.Id, (id ?? "").Trim(), StringComparison.OrdinalIgnoreCase)) return m;
            return null;
        }

        public static Mode Utama { get { return Semua[0]; } }

        // =================================================================
        // Mode aktif (pilihan lokal per pengguna)
        // =================================================================
        private static string FileSetelan { get { return Path.Combine(AppPaths.FolderLokal, "mode.txt"); } }

        private static string _idCache;
        private static DateTime _waktuCache;

        /// <summary>Mode yang tersimpan di setelan lokal (tanpa memeriksa lisensi).</summary>
        public static string IdTersimpan
        {
            get
            {
                if (_idCache != null && (DateTime.Now - _waktuCache).TotalSeconds < 5) return _idCache;
                string id = Akad;
                try
                {
                    if (File.Exists(FileSetelan))
                    {
                        string s = File.ReadAllText(FileSetelan).Trim().ToLowerInvariant();
                        if (Cari(s) != null) id = s;
                    }
                }
                catch { }
                _idCache = id; _waktuCache = DateTime.Now;
                return id;
            }
        }

        public static bool Lite
        {
            get { try { return Lisensi.Status().Edisi == Lisensi.Lite; } catch { return false; } }
        }

        /// <summary>Mode yang berlaku sekarang. Edisi Lite selalu mode akad.</summary>
        public static Mode Aktif
        {
            get
            {
                string id = IdTersimpan;
                if (id == Akad) return Utama;
                if (Lite) return Utama;
                return Cari(id) ?? Utama;
            }
        }

        public static bool AktifKantor { get { return Aktif.Id == Kantor; } }

        /// <summary>Ganti mode aktif di komputer ini. Panel memuat ulang seluruh tampilan sesudahnya.</summary>
        public static void Pilih(string id)
        {
            Mode m = Cari(id);
            if (m == null) throw new InvalidOperationException("Mode perhitungan tidak dikenal: " + id);
            if (!m.Utama && Lite)
                throw new InvalidOperationException("Edisi Lite hanya menghitung CKPN per jenis akad. Mode lain tersedia di edisi Pro.");
            if (CKPNPipeline.SedangBerjalan || BatchGrup.Aktif)
                throw new InvalidOperationException("Perhitungan sedang berjalan. Ganti mode setelah selesai.");
            Directory.CreateDirectory(AppPaths.FolderLokal);
            File.WriteAllText(FileSetelan, m.Id);
            _idCache = m.Id; _waktuCache = DateTime.Now;
            CatatanLog.Tulis("Mode perhitungan diganti ke " + m.Nama + " oleh " + Environment.UserName);
            LogProses.CatatPanel("", "Panel", "Mode perhitungan", LogProses.Info, LogProses.R()
                .Tambah("Mode aktif", m.Nama).Tambah("Database", m.FileDb));
        }

        // =================================================================
        // Mode resmi per tahun (database utama, tabel mode_tahun)
        // =================================================================
        internal class Resmi
        {
            public int Tahun;          // tahun yang diminta
            public int SumberTahun;    // tahun penetapan yang berlaku (0 = belum pernah ditetapkan)
            public string Mode = Akad, Dasar = "", Pengguna = "", Waktu = "";
        }

        /// <summary>Mode resmi tahun tertentu dari koneksi yang memuat tabel mode_tahun (utama atau terlampir).</summary>
        public static Resmi ModeResmi(SQLiteConnection con, int tahun)
        {
            var r = new Resmi { Tahun = tahun };
            try
            {
                object t = Database.Scalar(con, "SELECT MAX(tahun) FROM mode_tahun WHERE tahun<=@p0", tahun);
                if (t == null || t is DBNull) return r;
                using (var cmd = Database.Cmd(con, "SELECT tahun, mode, dasar, pengguna, waktu FROM mode_tahun WHERE tahun=@p0", Convert.ToInt32(t)))
                using (var rd = cmd.ExecuteReader())
                    if (rd.Read())
                    {
                        r.SumberTahun = rd.GetInt32(0);
                        r.Mode = Cari(Str(rd[1])) != null ? Str(rd[1]) : Akad;
                        r.Dasar = Str(rd[2]); r.Pengguna = Str(rd[3]); r.Waktu = Str(rd[4]);
                    }
            }
            catch { /* database lama tanpa tabel mode_tahun → bawaan akad */ }
            return r;
        }

        public static Resmi ModeResmi(int tahun)
        {
            try
            {
                if (!Database.AdaBersama) return new Resmi { Tahun = tahun };
                using (var con = Database.BukaBersama(false)) return ModeResmi(con, tahun);
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Baca mode resmi: " + ex.Message);
                return new Resmi { Tahun = tahun };
            }
        }

        /// <summary>Pesan penolakan bila mode aktif bukan mode resmi tahun periode ini; null = boleh.</summary>
        public static string TolakResmi(string tanggal, string aksi)
        {
            int tahun;
            if (string.IsNullOrEmpty(tanggal) || tanggal.Length < 4 || !int.TryParse(tanggal.Substring(0, 4), out tahun)) return null;
            var r = ModeResmi(tahun);
            Mode aktif = Aktif;
            if (r.Mode == aktif.Id) return null;
            Mode resmi = Cari(r.Mode) ?? Utama;
            return aksi + " hanya dapat dilakukan dari mode resmi tahun " + tahun + " (" + resmi.Nama + "). " +
                   "Mode aktif saat ini " + aktif.Nama + " — hasilnya simulasi. Ganti mode di ⚙ Pengaturan › Mode perhitungan.";
        }

        /// <summary>
        /// Tetapkan mode resmi tahun tertentu. Ditolak bila sudah ada periode Final di tahun itu (atau tahun
        /// berikutnya yang mengikuti penetapan ini) pada mode lain.
        /// </summary>
        public static Dictionary<string, object> TetapkanResmi(int tahun, string modeId, string dasar)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (Lite) throw new InvalidOperationException("Edisi Lite hanya memakai mode per jenis akad.");
            Mode m = Cari(modeId);
            if (m == null) throw new InvalidOperationException("Mode perhitungan tidak dikenal: " + modeId);
            if (string.IsNullOrWhiteSpace(dasar)) throw new InvalidOperationException("Dasar penetapan (kebijakan / memo direksi) wajib diisi.");
            if (tahun < 2000 || tahun > 2100) throw new InvalidOperationException("Tahun tidak valid.");

            // Rentang tahun yang ikut berubah: tahun ini sampai sebelum penetapan eksplisit berikutnya
            int sampai = 9999;
            if (Database.AdaBersama)
                using (var con = Database.BukaBersama(false))
                {
                    try
                    {
                        object n = Database.Scalar(con, "SELECT MIN(tahun) FROM mode_tahun WHERE tahun>@p0", tahun);
                        if (n != null && !(n is DBNull)) sampai = Convert.ToInt32(n) - 1;
                    }
                    catch { }
                }

            var bentrok = new List<string>();
            foreach (var lain in Semua)
            {
                if (lain.Id == m.Id || !Database.AdaMode(lain)) continue;
                try
                {
                    using (var con = Database.BukaMode(lain, false))
                    using (var cmd = Database.Cmd(con,
                        "SELECT tanggal FROM periode WHERE status='Final' AND CAST(substr(tanggal,1,4) AS INTEGER) BETWEEN @p0 AND @p1 ORDER BY tanggal",
                        tahun, sampai))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) bentrok.Add(rd.GetString(0) + " (" + lain.NamaPendek + ")");
                }
                catch (Exception ex) { CatatanLog.Tulis("Cek periode Final mode " + lain.Id + ": " + ex.Message); }
            }
            if (bentrok.Count > 0)
                throw new InvalidOperationException("Mode resmi tahun " + tahun + " tidak dapat diganti ke " + m.Nama +
                    " karena sudah ada periode Final di mode lain: " + string.Join(", ", bentrok.ToArray()) +
                    ". Buka kunci periode tersebut terlebih dahulu bila perubahan memang disetujui.");

            Resmi lama = ModeResmi(tahun);
            Database.Cadangkan();
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                Database.Exec(con, "INSERT OR REPLACE INTO mode_tahun(tahun,mode,dasar,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4)",
                    tahun, m.Id, dasar.Trim(), Environment.UserName, Database.Sekarang());
                Database.CatatAktivitas(con, tahun.ToString(), "mode-resmi",
                    (Cari(lama.Mode) ?? Utama).NamaPendek + " → " + m.NamaPendek + " | " + dasar.Trim());
                tx.Commit();
            }
            return Info(tahun);
        }

        // =================================================================
        // Kode grup & lingkup
        // =================================================================

        /// <summary>
        /// Normalisasi daftar kode unit menjadi kunci grup kanonik.
        ///   akad   : "KC0700, kc0600" → "KC0600,KC0700"
        ///   kantor : "1, 003 ,002"    → "001,002,003" (angka dipadatkan 3 digit, urut numerik, tanpa duplikat)
        /// </summary>
        public static string NormKode(Mode m, string kode)
        {
            var hasil = new List<string>();
            foreach (var bagian in (kode ?? "").Split(new[] { ',', ';', ' ', '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string k = m != null && m.Utama ? bagian.Trim().ToUpperInvariant() : NormUnit(m, bagian);
                if (k.Length > 0 && !hasil.Contains(k)) hasil.Add(k);
            }
            if (m != null && m.Utama)
            {
                var urut = new List<string>();
                foreach (var kc in SemuaKC) if (hasil.Contains(kc)) urut.Add(kc);
                foreach (var k in hasil) if (!urut.Contains(k)) urut.Add(k);
                hasil = urut;
            }
            else hasil.Sort(BandingSandi);
            return string.Join(",", hasil.ToArray());
        }

        /// <summary>Sandi kantor kanonik: angka → 3 digit ("1" / 1.0 / "0001" → "001"); selain angka → huruf besar.</summary>
        public static string NormSandi(object v)
        {
            if (v == null || v is DBNull) return "";
            string s;
            if (v is double) s = ((long)Math.Round((double)v)).ToString(CultureInfo.InvariantCulture);
            else s = Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            if (s.Length == 0) return "";
            bool angka = true;
            foreach (char c in s) if (c < '0' || c > '9') { angka = false; break; }
            if (!angka)
            {
                double d;
                if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d) && d >= 0 && d == Math.Floor(d) && d < 1e9)
                { s = ((long)d).ToString(CultureInfo.InvariantCulture); angka = true; }
            }
            if (!angka) return s.ToUpperInvariant();
            s = s.TrimStart('0');
            if (s.Length == 0) s = "0";
            return s.PadLeft(3, '0');
        }

        public static int BandingSandi(string a, string b)
        {
            long x, y;
            bool ax = long.TryParse(a, out x), by = long.TryParse(b, out y);
            if (ax && by) return x.CompareTo(y);
            if (ax != by) return ax ? -1 : 1;
            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Daftar KC yang dicentang di Master untuk menghitung grup ini.</summary>
        public static string[] KCGrup(string kodeGrup) { return KCGrup(Aktif, kodeGrup); }

        public static string[] KCGrup(Mode m, string kodeGrup)
        {
            if (m == null || m.Utama)
                return (kodeGrup ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            return (string[])SemuaKC.Clone();
        }

        /// <summary>Teks kode grup untuk ditampilkan, mis. "KC0600, KC0700" atau "sandi 001, 002".</summary>
        public static string LabelKode(string kodeGrup) { return LabelKode(Aktif, kodeGrup); }

        public static string LabelKode(Mode m, string kodeGrup)
        {
            string k = (kodeGrup ?? "").Replace(",", ", ");
            if (m == null || m.Utama) return k;
            if (k.Length == 0) return "semua " + m.NamaUnit;
            if (m.Id == Penggunaan)
            {
                var nama = new List<string>();
                foreach (var x in (kodeGrup ?? "").Split(',')) if (x.Length > 0) nama.Add(NamaPenggunaan(x));
                return string.Join(", ", nama.ToArray());
            }
            return m.NamaUnit + " " + k;
        }

        // =================================================================
        // Saringan lingkup di Master!H8
        // =================================================================
        /// <summary>Isi mentah Master!H8, mis. "kantor:001,002"; "" bila kosong.</summary>
        public static string BacaFilterMentah(Excel.Worksheet master)
        {
            if (master == null) return "";
            try
            {
                object v = ((Excel.Range)master.Range[SelFilter]).Value2;
                return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
            }
            catch { return ""; }
        }

        /// <summary>Kode unit saringan untuk mode tertentu ("" bila Master!H8 kosong atau untuk mode lain).</summary>
        public static string BacaFilter(Excel.Worksheet master, Mode m)
        {
            if (m == null || m.Utama) return "";
            string s = BacaFilterMentah(master);
            int i = s.IndexOf(':');
            if (i <= 0 || !string.Equals(s.Substring(0, i).Trim(), m.Id, StringComparison.OrdinalIgnoreCase)) return "";
            return NormKode(m, s.Substring(i + 1));
        }

        /// <summary>
        /// Tulis saringan lingkup ke Master!H8 (sel diformat teks agar "001,002" tidak berubah menjadi angka).
        /// kode kosong / mode akad = hapus saringan. Proteksi Master dibuka hanya bila isi berubah.
        /// </summary>
        public static void TulisFilter(Excel.Workbook wb, Mode m, string kode)
        {
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            if (master == null) throw new InvalidOperationException("Sheet 'Master' tidak ditemukan.");
            string isi = m == null || m.Utama || string.IsNullOrEmpty(kode) ? "" : m.Id + ":" + NormKode(m, kode);
            if (BacaFilterMentah(master) == isi) return;
            Protection.UbahMaster(wb, ms =>
            {
                Excel.Range r = (Excel.Range)ms.Range[SelFilter];
                r.NumberFormat = "@";
                r.Value2 = isi;
                try
                {
                    Excel.Range lbl = (Excel.Range)ms.Range["G8"];
                    if (lbl.Value2 == null) lbl.Value2 = "Lingkup mode (diisi panel)";
                }
                catch { }
            });
        }

        /// <summary>Hapus saringan bila mode aktif = akad (Lite / setelah ganti mode) supaya tombol Master menghitung per KC.</summary>
        public static void BersihkanFilterBilaAkad(Excel.Workbook wb)
        {
            if (wb == null || !Aktif.Utama) return;
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            if (master == null || BacaFilterMentah(master).Length == 0) return;
            TulisFilter(wb, Utama, "");
        }

        /// <summary>
        /// Kode grup yang sedang ada di workbook menurut mode aktif:
        ///   akad   → KC yang dicentang di Master ("KC0600,KC0700");
        ///   kantor → sandi di Master!H8 ("001,002"; "" bila tanpa saringan).
        /// </summary>
        public static string KodeGrupWorkbook(Excel.Worksheet master)
        {
            if (master == null) return "";
            Mode m = Aktif;
            if (m.Utama) return string.Join(",", ParameterMaster.BacaKCDicentang(master).ToArray());
            return BacaFilter(master, m);
        }

        // =================================================================
        // Info untuk panel
        // =================================================================
        public static Dictionary<string, object> Info(int tahun)
        {
            Mode a = Aktif;
            Resmi r = tahun > 0 ? ModeResmi(tahun) : new Resmi();
            var daftar = new List<object>();
            foreach (var m in Semua)
                daftar.Add(new Dictionary<string, object>
                {
                    { "id", m.Id }, { "nama", m.Nama }, { "namaPendek", m.NamaPendek }, { "keterangan", m.Keterangan },
                    { "namaUnit", m.NamaUnit }, { "utama", m.Utama }, { "file", m.FileDb },
                    { "adaDatabase", Database.AdaMode(m) }
                });
            string info;
            return new Dictionary<string, object>
            {
                { "aktif", a.Id }, { "namaAktif", a.Nama }, { "namaPendekAktif", a.NamaPendek }, { "namaUnit", a.NamaUnit },
                { "tersimpan", IdTersimpan }, { "lite", Lite }, { "daftar", daftar },
                { "tahun", tahun }, { "resmi", r.Mode }, { "namaResmi", (Cari(r.Mode) ?? Utama).Nama },
                { "resmiTahun", r.SumberTahun }, { "resmiDasar", r.Dasar }, { "resmiPengguna", r.Pengguna }, { "resmiWaktu", r.Waktu },
                { "simulasi", r.Mode != a.Id }, { "bolehMenulis", Database.BolehMenulis(out info) }
            };
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
