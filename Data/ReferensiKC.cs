using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Referensi Kode KC untuk rekening hapus buku KC2900 (Tahap 6d), pengganti kolom M + RefKCBuilder.
    ///
    /// Masalah lama: LGD ER memfilter KC2900 per segmen lewat kolom M (Kode KC) yang diisi RefKCBuilder
    /// dari 3 digit awal nomor rekening. Rekening yang kodenya tidak ada di KC0600–KC1100 mendapat kolom M
    /// kosong, sehingga diam-diam TIDAK ikut LGD ER grup mana pun. Peringatannya dulu hanya di sheet Audit Log.
    ///
    /// Sekarang:
    ///   1. PINDAI (otomatis saat LGD ER bila file tahunan Master!D60:D65 berubah, atau tombol di panel):
    ///      buka keenam file READ-ONLY, catat (a) setiap rekening aktif di KC0600–KC1100 beserta sheet-nya,
    ///      (b) jumlah rekening aktif per awalan kode produk per KC, (c) seluruh baris KC2900.
    ///      Hasil disimpan di database; file sumber TIDAK lagi ditulis/di-save.
    ///   2. TENTUKAN Kode KC tiap rekening KC2900, urutan prioritas:
    ///        manual-rek     isian petugas untuk satu rekening (panel)
    ///        rekening       rekening yang sama ditemukan di sheet KC0600–KC1100 salah satu file tahunan
    ///        manual-produk  isian petugas untuk satu awalan kode produk (panel; awalan terpanjang menang)
    ///        kolom-m        isi kolom M KC2900 di file (isian lama / manual), bila ada
    ///        produk         KC mayoritas untuk awalan kode produk (cara RefKCBuilder)
    ///      Kode "DIKECUALIKAN" = sengaja tidak masuk LGD ER segmen mana pun (tidak dihitung sebagai kekurangan).
    ///   3. SELESAI: daftar rekening yang benar-benar masuk tiap lingkup KC disimpan, sehingga panel bisa
    ///      menunjukkan grup mana yang perlu dihitung ulang setelah petugas melengkapi referensi.
    ///
    /// Nomor rekening dinormalisasi sama seperti LGDExpectedRecoveries.NormNoRek (tanpa apostrof; angka murni
    /// tanpa nol di depan). Awalan kode produk diambil dari nomor yang sudah dinormalisasi.
    /// </summary>
    internal static class ReferensiKC
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static readonly string[] DaftarKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };
        public const string Dikecualikan = "DIKECUALIKAN";
        public const int PanjangDefault = 3;
        /// <summary>Tahap 8: kode untuk rekening kantor lain (bukan KC0600–KC1100 → tidak masuk lingkup mana pun, tanpa peringatan).</summary>
        public const string LuarLingkup = "LUAR-KANTOR";
        /// <summary>Versi isi hasil pindai (2 = dengan sandi kantor).</summary>
        private const int VersiPindai = 3;   // 3 = + jenis penggunaan (Tahap 8c)
        private static readonly int[] PanjangDipindai = { 2, 3, 4, 5 };
        private const int BarisAwalKC = 5, BarisAwalHB = 3, JendelaCohort = 6;

        public const string SManualRek = "manual-rek", SRekening = "rekening", SManualProduk = "manual-produk",
                            SKolomM = "kolom-m", SProduk = "produk", STanpa = "tanpa";

        // =================================================================
        // Struktur data
        // =================================================================
        internal class RekHB
        {
            public string NoRek, RekAsli, Cif, Nama, Produk, TglHB, KolomM, KcRekening, TahunFile;
            /// <summary>Tahap 8: sandi kantor KC2900 kolom B ("" bila kosong/0) dan sandi rekening yang sama di KC0600–KC1100.</summary>
            public string SandiHB, SandiRekening;
            /// <summary>Tahap 8c: jenis penggunaan (1/2/3) rekening yang sama di KC0600–KC1100 file tahunan terbaru.</summary>
            public string GunaRekening;
            /// <summary>Sandi kantor yang dipakai mode per kantor: KC2900 kolom B, bila kosong/0 dari rekening aktif.</summary>
            public string SandiEfektif { get { return !string.IsNullOrEmpty(SandiHB) ? SandiHB : (SandiRekening ?? ""); } }
            public int TahunHB, KcRekeningTahun;
            public double Baki;
        }

        internal class HasilPindai
        {
            public List<RekHB> Rek = new List<RekHB>();
            /// <summary>kunci "panjang|awalan|KC" → jumlah rekening aktif</summary>
            public Dictionary<string, int> Agg = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            /// <summary>Tahap 8c: kunci "panjang|awalan|jenis penggunaan" → jumlah rekening aktif</summary>
            public Dictionary<string, int> AggGuna = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public List<Dictionary<string, object>> File = new List<Dictionary<string, object>>();
        }

        /// <summary>Semua bahan penentuan Kode KC (manual + hasil pindai).</summary>
        internal class Peta
        {
            public int Panjang = PanjangDefault;
            public Dictionary<string, string> ManualRek = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ManualProduk = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ScanRek = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            /// <summary>awalan (panjang terpilih) → KC → jumlah</summary>
            public Dictionary<string, Dictionary<string, int>> Distribusi = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Pemenang = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private List<string> _awalanManual;

            public void Bangun(Dictionary<string, int> agg)
            {
                Distribusi.Clear(); Pemenang.Clear();
                foreach (var kv in agg)
                {
                    string[] p = kv.Key.Split('|');
                    if (p.Length != 3 || p[0] != Panjang.ToString()) continue;
                    Dictionary<string, int> d;
                    if (!Distribusi.TryGetValue(p[1], out d)) Distribusi[p[1]] = d = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    d[p[2]] = (d.ContainsKey(p[2]) ? d[p[2]] : 0) + kv.Value;
                }
                // Sama dengan RefKCBuilder: jumlah terbanyak; bila seri, kode KC terkecil
                foreach (var kv in Distribusi)
                {
                    string best = ""; int n = -1;
                    foreach (var k in kv.Value)
                        if (k.Value > n || (k.Value == n && string.Compare(k.Key, best, StringComparison.OrdinalIgnoreCase) < 0)) { best = k.Key; n = k.Value; }
                    if (best.Length > 0) Pemenang[kv.Key] = best;
                }
                _awalanManual = new List<string>(ManualProduk.Keys);
                _awalanManual.Sort((a, b) => b.Length.CompareTo(a.Length));
            }

            // ---- Tahap 8c: jenis penggunaan rekening hapus buku (mode per jenis penggunaan) ----
            public Dictionary<string, string> GunaManualRek = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> GunaManualProduk = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> GunaScanRek = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, Dictionary<string, int>> GunaDistribusi = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> GunaPemenang = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private List<string> _awalanGunaManual;

            public void BangunGuna(Dictionary<string, int> agg)
            {
                GunaDistribusi.Clear(); GunaPemenang.Clear();
                foreach (var kv in agg ?? new Dictionary<string, int>())
                {
                    string[] p = kv.Key.Split('|');
                    if (p.Length != 3 || p[0] != Panjang.ToString()) continue;
                    Dictionary<string, int> d;
                    if (!GunaDistribusi.TryGetValue(p[1], out d)) GunaDistribusi[p[1]] = d = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    d[p[2]] = (d.ContainsKey(p[2]) ? d[p[2]] : 0) + kv.Value;
                }
                foreach (var kv in GunaDistribusi)
                {
                    string best = ""; int n = -1;
                    foreach (var k in kv.Value)
                        if (k.Value > n || (k.Value == n && string.CompareOrdinal(k.Key, best) < 0)) { best = k.Key; n = k.Value; }
                    if (best.Length > 0) GunaPemenang[kv.Key] = best;
                }
                _awalanGunaManual = new List<string>(GunaManualProduk.Keys);
                _awalanGunaManual.Sort((a, b) => b.Length.CompareTo(a.Length));
            }

            public string AwalanGunaManual(string norm)
            {
                if (_awalanGunaManual == null) BangunGuna(null);
                foreach (var a in _awalanGunaManual) if (norm.StartsWith(a, StringComparison.OrdinalIgnoreCase)) return a;
                return null;
            }

            /// <summary>Jenis penggunaan rekening hapus buku: manual rekening › rekening aktif › manual kode produk › mayoritas kode produk.</summary>
            public string TentukanGuna(string norm, out string sumber)
            {
                string g;
                if (norm.Length == 0) { sumber = STanpa; return ""; }
                if (GunaManualRek.TryGetValue(norm, out g)) { sumber = SManualRek; return g; }
                if (GunaScanRek.TryGetValue(norm, out g)) { sumber = SRekening; return g; }
                string a = AwalanGunaManual(norm);
                if (a != null) { sumber = SManualProduk; return GunaManualProduk[a]; }
                if (norm.Length >= Panjang && GunaPemenang.TryGetValue(norm.Substring(0, Panjang), out g)) { sumber = SProduk; return g; }
                sumber = STanpa;
                return "";
            }

            public string AwalanManual(string norm)
            {
                if (_awalanManual == null) Bangun(new Dictionary<string, int>());
                foreach (var a in _awalanManual) if (norm.StartsWith(a, StringComparison.OrdinalIgnoreCase)) return a;
                return null;
            }

            public string Tentukan(string norm, string kolomM, out string sumber)
            {
                string kc;
                if (norm.Length == 0) { sumber = STanpa; return ""; }
                if (ManualRek.TryGetValue(norm, out kc)) { sumber = SManualRek; return kc; }
                if (ScanRek.TryGetValue(norm, out kc)) { sumber = SRekening; return kc; }
                string a = AwalanManual(norm);
                if (a != null) { sumber = SManualProduk; return ManualProduk[a]; }
                if (!string.IsNullOrEmpty(kolomM)) { sumber = SKolomM; return kolomM; }
                if (norm.Length >= Panjang && Pemenang.TryGetValue(norm.Substring(0, Panjang), out kc)) { sumber = SProduk; return kc; }
                sumber = STanpa;
                return "";
            }
        }

        /// <summary>Dipakai LGDExpectedRecoveries: satu objek per perhitungan.</summary>
        internal class Resolver
        {
            internal Peta P;
            internal List<RekHB> Rek = new List<RekHB>();
            /// <summary>Tahap 8: saringan mode per kantor (null = tanpa saringan).</summary>
            internal FilterLingkup Filter;
            private Dictionary<string, RekHB> _peta;
            private readonly HashSet<string> _luarKantor = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _tanpaSandi = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private double _bakiTanpaSandi;
            internal int CurrYear;
            internal bool DariCache, Tersimpan;
            internal string Catatan;
            private readonly Dictionary<string, int> _sumber = new Dictionary<string, int>();

            /// <summary>Kode KC untuk satu baris KC2900 (rekening mentah kolom H, isi kolom M). "" = tidak terpetakan.</summary>
            public string Tentukan(string rekMentah, string kolomMMentah)
            {
                string s;
                string norm = NormRek(rekMentah);
                string kc = P.Tentukan(norm, NormKC(kolomMMentah), out s);
                _sumber[s] = (_sumber.ContainsKey(s) ? _sumber[s] : 0) + 1;
                // Tahap 8: mode per kantor — rekening kantor lain tidak masuk lingkup (kode di luar KC0600–KC1100)
                if (Filter != null && kc.Length > 0 && kc != Dikecualikan && !CocokKantor(norm))
                {
                    _luarKantor.Add(norm);
                    return LuarLingkup;
                }
                return kc;
            }

            /// <summary>True bila sandi kantor efektif rekening hapus buku termasuk grup kantor ini.</summary>
            internal bool CocokKantor(string norm)
            {
                if (Filter == null) return true;
                if (_peta == null)
                {
                    _peta = new Dictionary<string, RekHB>(StringComparer.OrdinalIgnoreCase);
                    foreach (var r in Rek) _peta[r.NoRek] = r;
                }
                RekHB x;
                _peta.TryGetValue(norm, out x);
                string sandi = UnitRek(P, x, norm, Filter.ModeId);   // sandi kantor / jenis penggunaan
                if (sandi.Length == 0)
                {
                    if (_tanpaSandi.Add(norm) && x != null) _bakiTanpaSandi += x.Baki;
                    return false;
                }
                return Filter.Cocok(sandi);
            }

            /// <summary>
            /// Dipanggil di akhir LGD ER: catat rekening yang masuk lingkup ini (untuk deteksi "perlu hitung ulang")
            /// dan tulis ringkasan ke log proses. Tidak pernah melempar error.
            /// </summary>
            public void Selesai(IEnumerable<string> lingkupKC)
            {
                try
                {
                    var scope = new HashSet<string>(lingkupKC ?? new string[0], StringComparer.OrdinalIgnoreCase);
                    string kunci = KunciLingkup(scope, Filter);
                    var masuk = new List<string>();
                    int tanpa = 0; double bakiTanpa = 0;
                    var tanpaAwalan = new Dictionary<string, int>();
                    foreach (var r in Rek)
                    {
                        if (!DalamCohort(r, CurrYear)) continue;
                        string s;
                        string kc = P.Tentukan(r.NoRek, NormKC(r.KolomM), out s);
                        if (kc.Length == 0)
                        {
                            tanpa++; bakiTanpa += r.Baki;
                            string a = Awalan(r.NoRek, P.Panjang);
                            tanpaAwalan[a] = (tanpaAwalan.ContainsKey(a) ? tanpaAwalan[a] : 0) + 1;
                        }
                        else if (scope.Contains(kc) && (Filter == null || Filter.Cocok(UnitRek(P, r, r.NoRek, Filter.ModeId)))) masuk.Add(r.NoRek);
                    }
                    masuk.Sort(StringComparer.OrdinalIgnoreCase);
                    if (Filter != null)
                        LogProses.Catat("LGD ER - Referensi KC", "Saringan " + Filter.Label, _tanpaSandi.Count == 0 ? LogProses.OK : LogProses.Peringatan, LogProses.R()
                            .Tambah("Rekening KC2900 di luar grup (tidak dihitung)", _luarKantor.Count - _tanpaSandi.Count)
                            .Tambah(Filter.ModeId == ModeHitung.Penggunaan
                                ? "Rekening tanpa jenis penggunaan (tidak ditemukan di KC0600–KC1100 dan belum diisi di Pengaturan)"
                                : "Rekening tanpa sandi kantor (KC2900 kolom B kosong/0 dan rekening tidak ditemukan di KC0600–KC1100)", _tanpaSandi.Count)
                            .TambahBila(_tanpaSandi.Count > 0, "Baki tanpa " + (Filter.ModeId == ModeHitung.Penggunaan ? "jenis penggunaan" : "sandi kantor"), _bakiTanpaSandi)
                            .TambahBila(_tanpaSandi.Count > 0, "Keterangan", "Rekening ini tidak masuk LGD ER grup mana pun. Lengkapi di ⚙ Pengaturan › Referensi hapus buku."));

                    string alasan;
                    if (Tersimpan && Database.BolehMenulis(out alasan))
                    {
                        using (var con = Database.BukaBersama(true))
                        using (var tx = con.BeginTransaction())
                        {
                            Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_lingkup(lingkup,curr_year,rekening,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4)",
                                kunci, CurrYear, _json.Serialize(masuk), Environment.UserName, Database.Sekarang());
                            SimpanInfo(con, "dipakai", _json.Serialize(new Dictionary<string, object>
                            {
                                { "waktu", Database.Sekarang() }, { "pengguna", Environment.UserName }, { "lingkup", kunci },
                                { "currYear", CurrYear }, { "tanpa", tanpa }, { "bakiTanpa", bakiTanpa }
                            }));
                            tx.Commit();
                        }
                    }

                    var ringkas = new List<string>();
                    foreach (var kv in tanpaAwalan) ringkas.Add(kv.Key + " (" + kv.Value + ")");
                    LogProses.Catat("LGD ER - Referensi KC", "Lingkup " + kunci, tanpa == 0 ? LogProses.OK : LogProses.Peringatan, LogProses.R()
                        .Tambah("Sumber referensi", DariCache ? "database (file tahunan tidak berubah)" : "pindai ulang file tahunan")
                        .TambahBila(!string.IsNullOrEmpty(Catatan), "Catatan", Catatan)
                        .Tambah("Rek masuk lingkup (cohort " + (CurrYear - JendelaCohort + 1) + "-" + CurrYear + ")", masuk.Count)
                        .Tambah("Baris dari referensi manual rekening", Hit(SManualRek))
                        .Tambah("Baris cocok rekening KC0600-KC1100", Hit(SRekening))
                        .Tambah("Baris dari referensi manual kode produk", Hit(SManualProduk))
                        .Tambah("Baris dari kolom M file", Hit(SKolomM))
                        .Tambah("Baris dari KC mayoritas kode produk", Hit(SProduk))
                        .Tambah("Rek cohort TANPA Kode KC", tanpa)
                        .TambahBila(tanpa > 0, "Baki tanpa Kode KC", bakiTanpa)
                        .TambahBila(tanpa > 0, "Awalan kode produk tanpa KC", string.Join(", ", ringkas.ToArray()))
                        .TambahBila(tanpa > 0, "Tindakan", "Lengkapi di Panel CKPN › Pengaturan › Referensi KC hapus buku, lalu hitung ulang grup"));
                }
                catch (Exception ex)
                {
                    CatatanLog.Tulis("Referensi KC: gagal mencatat hasil LGD ER: " + ex.Message);
                }
            }

            private int Hit(string s) { return _sumber.ContainsKey(s) ? _sumber[s] : 0; }
        }

        // =================================================================
        // Dipanggil LGDExpectedRecoveries.Hitung (thread Excel)
        // =================================================================
        public static Resolver Siapkan(Excel.Application app, Dictionary<int, string> filePaths, int currYear)
        {
            var r = new Resolver { CurrYear = currYear, P = new Peta() };
            string alasan;
            bool boleh = Database.BolehMenulis(out alasan);
            Dictionary<string, int> agg = null, aggGuna = null;
            List<RekHB> rek = null;

            if (Database.AdaBersama || boleh)
            {
                try
                {
                    using (var con = Database.BukaBersama(boleh))
                    {
                        BacaManual(con, r.P);
                        if (CacheCocok(con, filePaths))
                        {
                            rek = BacaRek(con);
                            agg = BacaAgg(con);
                            aggGuna = BacaAggGuna(con);
                            r.DariCache = true;
                            r.Tersimpan = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    r.Catatan = "Database tidak dapat dibaca (" + ex.Message + "); referensi manual tidak dipakai.";
                    CatatanLog.Tulis("Referensi KC: " + r.Catatan);
                }
            }

            if (rek == null)
            {
                var h = Pindai(app, filePaths);
                rek = h.Rek; agg = h.Agg; aggGuna = h.AggGuna;
                if (boleh)
                {
                    try { SimpanPindai(h, filePaths, currYear); r.Tersimpan = true; }
                    catch (Exception ex)
                    {
                        r.Catatan = "Hasil pindai tidak tersimpan: " + ex.Message;
                        CatatanLog.Tulis("Referensi KC: " + r.Catatan);
                    }
                }
                else r.Catatan = "User bukan pengirim — hasil pindai hanya dipakai untuk perhitungan ini.";
            }

            r.Rek = rek;
            foreach (var x in rek) if (!string.IsNullOrEmpty(x.KcRekening)) r.P.ScanRek[x.NoRek] = x.KcRekening;
            r.P.Bangun(agg);
            foreach (var x in rek) if (!string.IsNullOrEmpty(x.GunaRekening)) r.P.GunaScanRek[x.NoRek] = x.GunaRekening;
            r.P.BangunGuna(aggGuna);
            r.Filter = FilterLingkup.Dari(app);   // Tahap 8: mode per kantor (null = semua kantor)
            return r;
        }

        // =================================================================
        // Pindai file tahunan
        // =================================================================
        public static HasilPindai Pindai(Excel.Application app, Dictionary<int, string> filePaths)
        {
            var h = new HasilPindai();
            var tahun = new List<int>(filePaths.Keys);
            tahun.Sort();

            // rekening aktif: norm → [KC, tahun, nama]; file lebih baru menimpa
            var aktif = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
            // KC2900: norm → data per tahun file
            var hb = new Dictionary<string, RekHB>(StringComparer.OrdinalIgnoreCase);
            var bakiPerTahun = new Dictionary<string, Dictionary<int, double>>(StringComparer.OrdinalIgnoreCase);

            bool layar = true, alert = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; app.ScreenUpdating = false; app.DisplayAlerts = false; } catch { }
            try
            {
                int urut = 0;
                foreach (int y in tahun)
                {
                    string path = filePaths[y];
                    try { Pemberitahu.Progres("Referensi KC", ++urut, tahun.Count, "Pindai file tahun " + y); } catch { }
                    if (!File.Exists(path)) throw new FileNotFoundException("File tahun " + y + " tidak ditemukan: " + path);

                    Excel.Workbook wb = null;
                    bool dibukaDiSini = false;
                    int nAktif = 0, nHB = 0;
                    try
                    {
                        string full = Path.GetFullPath(path);
                        foreach (Excel.Workbook w in app.Workbooks)
                            if (string.Equals(w.FullName, full, StringComparison.OrdinalIgnoreCase)) { wb = w; break; }
                        if (wb == null) { wb = app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true); dibukaDiSini = true; }

                        foreach (string kc in DaftarKC)
                        {
                            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kc);
                            if (ws == null) continue;
                            object[,] d = BacaArea(ws);
                            if (d == null) continue;
                            for (int i = BarisAwalKC; i <= d.GetUpperBound(0); i++)
                            {
                                string nama = Teks(Sel(d, i, 4));
                                string norm = NormRek(Teks(Sel(d, i, 10)));
                                if (nama.Length == 0 || norm.Length == 0) continue;
                                nAktif++;
                                // Tahap 8c: jenis penggunaan kolom S; dalam satu file, baris pembiayaan pertama kontrak yang menentukan
                                string guna = ModeHitung.NormPenggunaan(Sel(d, i, 19));
                                if (guna.Length == 0) guna = ModeHitung.PenggunaanCadangan;
                                string[] lama;
                                if (aktif.TryGetValue(norm, out lama) && lama[1] == y.ToString()) guna = lama[4];
                                aktif[norm] = new[] { kc, y.ToString(), nama, SandiSah(Sel(d, i, 2)), guna };   // Tahap 8: + sandi kantor kolom B
                                foreach (int pg in PanjangDipindai)
                                {
                                    if (norm.Length < pg) continue;
                                    string kg = pg + "|" + norm.Substring(0, pg) + "|" + guna;
                                    h.AggGuna[kg] = (h.AggGuna.ContainsKey(kg) ? h.AggGuna[kg] : 0) + 1;
                                }
                                foreach (int p in PanjangDipindai)
                                {
                                    if (norm.Length < p) continue;
                                    string k = p + "|" + norm.Substring(0, p) + "|" + kc;
                                    h.Agg[k] = (h.Agg.ContainsKey(k) ? h.Agg[k] : 0) + 1;
                                }
                            }
                        }

                        Excel.Worksheet wsHB = ParameterMaster.CariSheet(wb, "KC2900");
                        object[,] dh = wsHB == null ? null : BacaArea(wsHB);
                        if (dh != null)
                            for (int i = BarisAwalHB; i <= dh.GetUpperBound(0); i++)
                            {
                                string mentah = Teks(Sel(dh, i, 8));
                                string norm = NormRek(mentah);
                                if (norm.Length == 0) continue;
                                nHB++;
                                double baki = Angka(Sel(dh, i, 12));
                                RekHB r;
                                if (!hb.TryGetValue(norm, out r))
                                {
                                    string tgl = ProfilTemplate.Tanggal(Sel(dh, i, 9));
                                    r = new RekHB
                                    {
                                        NoRek = norm, RekAsli = mentah, Cif = Teks(Sel(dh, i, 3)), Produk = Teks(Sel(dh, i, 7)),
                                        TglHB = tgl, TahunHB = TahunDari(tgl, y), KolomM = NormKC(Teks(Sel(dh, i, 13))),
                                        SandiHB = SandiSah(Sel(dh, i, 2))
                                    };
                                    hb[norm] = r;
                                    bakiPerTahun[norm] = new Dictionary<int, double>();
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(r.KolomM)) r.KolomM = NormKC(Teks(Sel(dh, i, 13)));
                                    if (string.IsNullOrEmpty(r.SandiHB)) r.SandiHB = SandiSah(Sel(dh, i, 2));
                                }
                                var bt = bakiPerTahun[norm];
                                bt[y] = (bt.ContainsKey(y) ? bt[y] : 0) + baki;
                            }
                    }
                    finally
                    {
                        if (dibukaDiSini && wb != null) try { wb.Close(false); } catch { }
                    }
                    h.File.Add(new Dictionary<string, object>
                    {
                        { "tahun", y }, { "path", path }, { "cap", CapFile(path) }, { "aktif", nAktif }, { "kc2900", nHB }
                    });
                }
            }
            finally
            {
                try { app.DisplayAlerts = alert; app.ScreenUpdating = layar; } catch { }
            }

            foreach (var kv in hb)
            {
                var r = kv.Value;
                var bt = bakiPerTahun[kv.Key];
                var th = new List<int>(bt.Keys);
                th.Sort();
                r.TahunFile = string.Join(",", th.ConvertAll(t => t.ToString()).ToArray());
                // Baki saat hapus buku = file tahun hapus buku; bila tidak ada, file tertua yang memuatnya
                r.Baki = bt.ContainsKey(r.TahunHB) ? bt[r.TahunHB] : bt[th[0]];
                string[] a;
                if (aktif.TryGetValue(kv.Key, out a)) { r.KcRekening = a[0]; r.KcRekeningTahun = int.Parse(a[1]); r.Nama = a[2]; r.SandiRekening = a[3]; r.GunaRekening = a[4]; }
                h.Rek.Add(r);
            }
            return h;
        }

        private static void SimpanPindai(HasilPindai h, Dictionary<int, string> filePaths, int currYear)
        {
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                Database.Exec(con, "DELETE FROM ref_kc_hapusbuku");
                Database.Exec(con, "DELETE FROM ref_kc_produk");
                Database.Exec(con, "DELETE FROM ref_guna_produk");
                using (var cmd = Database.Cmd(con,
                    "INSERT OR REPLACE INTO ref_kc_hapusbuku(no_rek,rek_asli,cif,nama,jenis_instrumen,tgl_hb,tahun_hb,baki,tahun_file,kolom_m,kc_rekening,kc_rekening_tahun,sandi_hb,sandi_rekening,guna_rekening) " +
                    "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13,@p14)",
                    "", "", "", "", "", "", 0, 0.0, "", "", "", 0, "", "", ""))
                {
                    foreach (var r in h.Rek)
                    {
                        object[] v = { r.NoRek, r.RekAsli, r.Cif, r.Nama, r.Produk, r.TglHB, r.TahunHB, r.Baki, r.TahunFile, r.KolomM,
                                       r.KcRekening, r.KcRekening == null ? (object)null : r.KcRekeningTahun,
                                       string.IsNullOrEmpty(r.SandiHB) ? null : r.SandiHB, string.IsNullOrEmpty(r.SandiRekening) ? null : r.SandiRekening,
                                       string.IsNullOrEmpty(r.GunaRekening) ? null : r.GunaRekening };
                        for (int i = 0; i < v.Length; i++) cmd.Parameters["@p" + i].Value = v[i] ?? DBNull.Value;
                        cmd.ExecuteNonQuery();
                    }
                }
                using (var cmd = Database.Cmd(con, "INSERT INTO ref_kc_produk(panjang,awalan,kode_kc,jumlah) VALUES(@p0,@p1,@p2,@p3)", 0, "", "", 0))
                {
                    foreach (var kv in h.Agg)
                    {
                        string[] p = kv.Key.Split('|');
                        cmd.Parameters["@p0"].Value = int.Parse(p[0]);
                        cmd.Parameters["@p1"].Value = p[1];
                        cmd.Parameters["@p2"].Value = p[2];
                        cmd.Parameters["@p3"].Value = kv.Value;
                        cmd.ExecuteNonQuery();
                    }
                }
                using (var cmd = Database.Cmd(con, "INSERT INTO ref_guna_produk(panjang,awalan,kode,jumlah) VALUES(@p0,@p1,@p2,@p3)", 0, "", "", 0))
                {
                    foreach (var kv in h.AggGuna)
                    {
                        string[] pg = kv.Key.Split('|');
                        cmd.Parameters["@p0"].Value = int.Parse(pg[0]);
                        cmd.Parameters["@p1"].Value = pg[1];
                        cmd.Parameters["@p2"].Value = pg[2];
                        cmd.Parameters["@p3"].Value = kv.Value;
                        cmd.ExecuteNonQuery();
                    }
                }
                SimpanInfo(con, "pindai", _json.Serialize(new Dictionary<string, object>
                {
                    { "waktu", Database.Sekarang() }, { "pengguna", Environment.UserName }, { "currYear", currYear }, { "file", h.File },
                    { "versi", VersiPindai }
                }));
                Database.CatatAktivitas(con, "", "referensi-kc", "Pindai " + h.File.Count + " file tahunan: " + h.Rek.Count + " rekening KC2900");
                tx.Commit();
            }
            int tanpaRek = 0;
            foreach (var r in h.Rek) if (string.IsNullOrEmpty(r.KcRekening)) tanpaRek++;
            LogProses.Catat("LGD ER - Referensi KC", "Pindai file tahunan", LogProses.OK, LogProses.R()
                .Tambah("Jumlah file", h.File.Count)
                .Tambah("Rekening KC2900", h.Rek.Count)
                .Tambah("Cocok rekening di KC0600-KC1100", h.Rek.Count - tanpaRek)
                .Tambah("Tidak ditemukan di KC0600-KC1100", tanpaRek)
                .Tambah("Keterangan", "File sumber dibuka read-only; kolom M tidak lagi ditulis"));
        }

        // =================================================================
        // Panel: data, simpan, hapus, pindai ulang
        // =================================================================
        /// <param name="grup">daftar grup periode aktif: nama + kode KC (untuk status "perlu hitung ulang")</param>
        public static Dictionary<string, object> Data(List<KeyValuePair<string, string[]>> grup, Dictionary<string, string> unitGrup = null)
        {
            string alasan;
            bool boleh = Database.BolehMenulis(out alasan);
            var hasil = new Dictionary<string, object>
            {
                { "ada", false }, { "bolehMenulis", boleh }, { "infoPengirim", alasan }, { "daftarKC", DaftarKC },
                { "panjang", PanjangDefault }, { "dikecualikan", Dikecualikan }
            };
            if (!Database.AdaBersama) return hasil;

            var p = new Peta();
            List<RekHB> rek;
            Dictionary<string, int> agg, aggGuna;
            var manual = new List<object>();
            var manualGuna = new List<object>();
            var lingkup = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, object> pindai = null, dipakai = null;
            using (var con = Database.BukaBersama(false))
            {
                pindai = BacaInfoJson(con, "pindai");
                dipakai = BacaInfoJson(con, "dipakai");
                BacaManual(con, p);
                rek = BacaRek(con);
                agg = BacaAgg(con);
                aggGuna = BacaAggGuna(con);
                try
                {
                    using (var cmd = Database.Cmd(con, "SELECT jenis,kunci,kode,alasan,pengguna,waktu FROM ref_guna_manual ORDER BY jenis,kunci"))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read())
                            manualGuna.Add(new Dictionary<string, object>
                            {
                                { "jenis", rd.GetString(0) }, { "kunci", rd.GetString(1) }, { "kode", rd.GetString(2) },
                                { "nama", ModeHitung.NamaPenggunaan(rd.GetString(2)) },
                                { "alasan", rd.IsDBNull(3) ? "" : rd.GetString(3) }, { "pengguna", rd.IsDBNull(4) ? "" : rd.GetString(4) },
                                { "waktu", rd.IsDBNull(5) ? "" : rd.GetString(5) }
                            });
                }
                catch { }
                using (var cmd = Database.Cmd(con, "SELECT jenis,kunci,kode_kc,alasan,pengguna,waktu FROM ref_kc_manual ORDER BY jenis,kunci"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        manual.Add(new Dictionary<string, object>
                        {
                            { "jenis", rd.GetString(0) }, { "kunci", rd.GetString(1) }, { "kodeKC", rd.GetString(2) },
                            { "alasan", rd.IsDBNull(3) ? "" : rd.GetString(3) }, { "pengguna", rd.IsDBNull(4) ? "" : rd.GetString(4) },
                            { "waktu", rd.IsDBNull(5) ? "" : rd.GetString(5) }
                        });
                using (var cmd = Database.Cmd(con, "SELECT lingkup,curr_year,rekening,pengguna,waktu FROM ref_kc_lingkup"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        lingkup[rd.GetString(0)] = new Dictionary<string, object>
                        {
                            { "currYear", rd.GetInt32(1) }, { "rekening", rd.GetString(2) },
                            { "pengguna", rd.IsDBNull(3) ? "" : rd.GetString(3) }, { "waktu", rd.IsDBNull(4) ? "" : rd.GetString(4) }
                        };
            }
            int panjangSimpan = p.Panjang;
            foreach (var x in rek) if (!string.IsNullOrEmpty(x.KcRekening)) p.ScanRek[x.NoRek] = x.KcRekening;
            p.Bangun(agg);
            foreach (var x in rek) if (!string.IsNullOrEmpty(x.GunaRekening)) p.GunaScanRek[x.NoRek] = x.GunaRekening;
            p.BangunGuna(aggGuna);

            int currYear = pindai != null && pindai.ContainsKey("currYear") ? Convert.ToInt32(pindai["currYear"]) : DateTime.Today.Year;
            hasil["ada"] = pindai != null;
            hasil["panjang"] = panjangSimpan;
            hasil["currYear"] = currYear;
            hasil["dipakai"] = dipakai;

            // Status file: berubah sejak dipindai?
            if (pindai != null)
            {
                bool berubah = false;
                var files = pindai["file"] as System.Collections.IList;
                if (files != null)
                    foreach (Dictionary<string, object> f in files)
                    {
                        string path = Convert.ToString(f["path"]);
                        bool beda = Convert.ToString(f["cap"]) != CapFile(path);
                        f["berubah"] = beda;
                        f.Remove("cap");
                        berubah |= beda;
                    }
                pindai["berubah"] = berubah;
            }
            hasil["pindai"] = pindai;

            // Baris per rekening + ringkasan per sumber
            var perSumber = new Dictionary<string, double[]>();
            foreach (var s in new[] { SManualRek, SRekening, SManualProduk, SKolomM, SProduk, STanpa }) perSumber[s] = new double[4];
            var tanpa = new List<object>();
            var bedaM = new List<object>();
            var dikecualikan = new List<object>();
            int nBedaM = 0, nCohort = 0;
            var hbPerAwalan = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var kcRek = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var rekMap = new Dictionary<string, RekHB>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rek)
            {
                string s;
                string kc = p.Tentukan(r.NoRek, r.KolomM, out s);
                bool cohort = DalamCohort(r, currYear);
                kcRek[r.NoRek] = kc;
                rekMap[r.NoRek] = r;
                if (cohort) nCohort++;
                var ps = perSumber[s];
                ps[0]++; ps[1] += r.Baki;
                if (cohort) { ps[2]++; ps[3] += r.Baki; }
                string aw = Awalan(r.NoRek, p.Panjang);
                hbPerAwalan[aw] = (hbPerAwalan.ContainsKey(aw) ? hbPerAwalan[aw] : 0) + 1;
                if (kc.Length == 0) tanpa.Add(Baris(r, kc, s, cohort, p));
                else if (kc == Dikecualikan) dikecualikan.Add(Baris(r, kc, s, cohort, p));
                if (!string.IsNullOrEmpty(r.KolomM) && kc.Length > 0 && !string.Equals(r.KolomM, kc, StringComparison.OrdinalIgnoreCase))
                {
                    nBedaM++;
                    if (bedaM.Count < 200) bedaM.Add(Baris(r, kc, s, cohort, p));
                }
            }
            var ringkas = new Dictionary<string, object>();
            foreach (var kv in perSumber)
                ringkas[kv.Key] = new Dictionary<string, object> { { "n", kv.Value[0] }, { "baki", kv.Value[1] }, { "nCohort", kv.Value[2] }, { "bakiCohort", kv.Value[3] } };
            hasil["ringkas"] = ringkas;
            hasil["total"] = rek.Count;
            hasil["totalCohort"] = nCohort;
            hasil["tanpa"] = tanpa;
            hasil["dikecualikanRek"] = dikecualikan;
            hasil["bedaM"] = bedaM;
            hasil["jumlahBedaM"] = nBedaM;

            // Hitungan rekening per isian manual
            foreach (Dictionary<string, object> m in manual)
            {
                string k = Convert.ToString(m["kunci"]);
                int n = 0;
                if (Convert.ToString(m["jenis"]) == "rek") n = rekMap.ContainsKey(k) ? 1 : 0;
                else foreach (var r in rek) if (r.NoRek.StartsWith(k, StringComparison.OrdinalIgnoreCase) && p.AwalanManual(r.NoRek) == k && !p.ManualRek.ContainsKey(r.NoRek) && !p.ScanRek.ContainsKey(r.NoRek)) n++;
                m["jumlahRek"] = n;
            }
            hasil["manual"] = manual;

            // Kode produk (awalan panjang terpilih)
            var produk = new List<object>();
            var semuaAwalan = new SortedSet<string>(p.Distribusi.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var a in hbPerAwalan.Keys) semuaAwalan.Add(a);
            foreach (var a in semuaAwalan)
            {
                var dist = new List<object>();
                int total = 0, maks = 0;
                Dictionary<string, int> d;
                if (p.Distribusi.TryGetValue(a, out d))
                {
                    var kunci = new List<string>(d.Keys);
                    kunci.Sort((x, y) => d[y].CompareTo(d[x]));
                    foreach (var k in kunci) { dist.Add(new Dictionary<string, object> { { "kc", k }, { "n", d[k] } }); total += d[k]; maks = Math.Max(maks, d[k]); }
                }
                string pem;
                p.Pemenang.TryGetValue(a, out pem);
                string man;
                p.ManualProduk.TryGetValue(a, out man);
                produk.Add(new Dictionary<string, object>
                {
                    { "awalan", a }, { "distribusi", dist }, { "pemenang", pem }, { "manual", man },
                    { "kemurnian", total == 0 ? (object)null : (double)maks / total },
                    { "hapusBuku", hbPerAwalan.ContainsKey(a) ? hbPerAwalan[a] : 0 }
                });
            }
            hasil["produk"] = produk;

            // Status per grup: rekening yang masuk lingkup sekarang vs saat LGD ER terakhir dihitung
            var statusGrup = new List<object>();
            if (grup != null)
                foreach (var g in grup)
                {
                    var scope = new HashSet<string>(g.Value, StringComparer.OrdinalIgnoreCase);
                    // Tahap 8: mode per kantor — lingkup juga disaring sandi kantor grup
                    FilterLingkup f = null;
                    string unit;
                    if (!ModeHitung.Aktif.Utama && unitGrup != null && unitGrup.TryGetValue(g.Key, out unit) && !string.IsNullOrEmpty(unit))
                        f = FilterLingkup.Buat(ModeHitung.Aktif, unit);
                    string kunci = KunciLingkup(scope, f);
                    var sekarang = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var r in rek)
                        if (DalamCohort(r, currYear) && scope.Contains(kcRek[r.NoRek]) &&
                            (f == null || f.Cocok(UnitRek(p, r, r.NoRek, f.ModeId)))) sekarang.Add(r.NoRek);
                    var item = new Dictionary<string, object> { { "nama", g.Key }, { "lingkup", kunci }, { "jumlah", sekarang.Count } };
                    Dictionary<string, object> l;
                    if (!lingkup.TryGetValue(kunci, out l) || Convert.ToInt32(l["currYear"]) != currYear)
                        item["status"] = "belum";
                    else
                    {
                        var dulu = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var o in (System.Collections.IEnumerable)_json.DeserializeObject(Convert.ToString(l["rekening"]))) dulu.Add(Convert.ToString(o));
                        int tambah = 0, kurang = 0; double bTambah = 0, bKurang = 0;
                        foreach (var k in sekarang) if (!dulu.Contains(k)) { tambah++; bTambah += rekMap[k].Baki; }
                        foreach (var k in dulu) if (!sekarang.Contains(k)) { kurang++; RekHB r; if (rekMap.TryGetValue(k, out r)) bKurang += r.Baki; }
                        item["status"] = tambah + kurang == 0 ? "ok" : "ulang";
                        item["tambah"] = tambah; item["kurang"] = kurang; item["bakiTambah"] = bTambah; item["bakiKurang"] = bKurang;
                        item["waktu"] = l["waktu"]; item["pengguna"] = l["pengguna"];
                    }
                    statusGrup.Add(item);
                }
            hasil["grup"] = statusGrup;
            hasil["guna"] = DataGuna(p, rek, manualGuna, currYear);
            return hasil;
        }

        /// <summary>
        /// Tahap 8c: jenis penggunaan rekening hapus buku untuk ⚙ Pengaturan — asal penentuan, rekening cohort yang belum
        /// punya jenis penggunaan, sebaran per kode produk, dan isian manual.
        /// </summary>
        private static Dictionary<string, object> DataGuna(Peta p, List<RekHB> rek, List<object> manual, int currYear)
        {
            var perSumber = new Dictionary<string, double[]>();
            foreach (var s in new[] { SManualRek, SRekening, SManualProduk, SProduk, STanpa }) perSumber[s] = new double[2];
            var perKode = new Dictionary<string, double[]>();
            var tanpa = new List<object>();
            var hbPerAwalan = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rek)
            {
                if (!DalamCohort(r, currYear)) continue;
                string sumber;
                string g = p.TentukanGuna(r.NoRek, out sumber);
                perSumber[sumber][0]++; perSumber[sumber][1] += r.Baki;
                if (g.Length > 0)
                {
                    double[] k;
                    if (!perKode.TryGetValue(g, out k)) perKode[g] = k = new double[2];
                    k[0]++; k[1] += r.Baki;
                }
                else tanpa.Add(new Dictionary<string, object>
                {
                    { "kunci", r.NoRek }, { "rek", r.RekAsli }, { "cif", r.Cif }, { "nama", r.Nama }, { "produk", r.Produk },
                    { "tahunHB", r.TahunHB }, { "baki", r.Baki }, { "awalan", Awalan(r.NoRek, p.Panjang) }
                });
                string aw = Awalan(r.NoRek, p.Panjang);
                hbPerAwalan[aw] = (hbPerAwalan.ContainsKey(aw) ? hbPerAwalan[aw] : 0) + 1;
            }
            var ringkas = new Dictionary<string, object>();
            foreach (var kv in perSumber) ringkas[kv.Key] = new Dictionary<string, object> { { "n", kv.Value[0] }, { "baki", kv.Value[1] } };
            var kode = new List<object>();
            for (int i = 0; i < ModeHitung.JenisPenggunaan.GetLength(0); i++)
            {
                double[] k;
                perKode.TryGetValue(ModeHitung.JenisPenggunaan[i, 0], out k);
                kode.Add(new Dictionary<string, object>
                {
                    { "kode", ModeHitung.JenisPenggunaan[i, 0] }, { "nama", ModeHitung.JenisPenggunaan[i, 1] },
                    { "n", k == null ? 0 : k[0] }, { "baki", k == null ? 0 : k[1] }
                });
            }
            var produk = new List<object>();
            var semua = new SortedSet<string>(hbPerAwalan.Keys, StringComparer.OrdinalIgnoreCase);
            foreach (var a in semua)
            {
                var dist = new List<object>();
                Dictionary<string, int> d;
                if (p.GunaDistribusi.TryGetValue(a, out d))
                    for (int i = 0; i < ModeHitung.JenisPenggunaan.GetLength(0); i++)
                    {
                        int n;
                        if (d.TryGetValue(ModeHitung.JenisPenggunaan[i, 0], out n) && n > 0)
                            dist.Add(new Dictionary<string, object> { { "kode", ModeHitung.JenisPenggunaan[i, 0] }, { "n", n } });
                    }
                string pem, man;
                p.GunaPemenang.TryGetValue(a, out pem);
                p.GunaManualProduk.TryGetValue(a, out man);
                produk.Add(new Dictionary<string, object>
                {
                    { "awalan", a }, { "distribusi", dist }, { "pemenang", pem }, { "manual", man }, { "hapusBuku", hbPerAwalan[a] }
                });
            }
            tanpa.Sort((x, y) => ((double)((Dictionary<string, object>)y)["baki"]).CompareTo((double)((Dictionary<string, object>)x)["baki"]));
            return new Dictionary<string, object>
            {
                { "ringkas", ringkas }, { "perKode", kode }, { "tanpa", tanpa }, { "produk", produk }, { "manual", manual },
                { "panjang", p.Panjang }
            };
        }

        private static Dictionary<string, object> Baris(RekHB r, string kc, string sumber, bool cohort, Peta p)
        {
            return new Dictionary<string, object>
            {
                { "kunci", r.NoRek }, { "rek", r.RekAsli }, { "cif", r.Cif }, { "nama", r.Nama }, { "produk", r.Produk },
                { "tglHB", r.TglHB }, { "tahunHB", r.TahunHB }, { "baki", r.Baki }, { "tahunFile", r.TahunFile },
                { "kolomM", r.KolomM }, { "kc", kc }, { "sumber", sumber }, { "cohort", cohort }, { "awalan", Awalan(r.NoRek, p.Panjang) }
            };
        }

        /// <summary>Unit lingkup rekening hapus buku menurut mode: sandi kantor efektif atau jenis penggunaan ("" = tidak diketahui).</summary>
        internal static string UnitRek(Peta p, RekHB r, string norm, string modeId)
        {
            if (modeId == ModeHitung.Penggunaan) { string s; return p.TentukanGuna(norm, out s); }
            return r == null ? "" : r.SandiEfektif;
        }

        /// <summary>
        /// Tahap 8c: isian manual jenis penggunaan rekening hapus buku. item: jenis 'rek'|'produk', kunci, kode ("1"/"2"/"3"; "" = hapus).
        /// </summary>
        public static Dictionary<string, object> SimpanManualGuna(System.Collections.IList item, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Dasar / alasan pengisian referensi wajib diisi.");
            if (item == null || item.Count == 0) throw new InvalidOperationException("Tidak ada isian yang disimpan.");
            var catatan = new List<string>();
            Database.Cadangkan();
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                foreach (Dictionary<string, object> x in item)
                {
                    string jenis = Convert.ToString(x["jenis"]);
                    string kunci = Convert.ToString(x["kunci"] ?? "");
                    string kode = ModeHitung.NormPenggunaan(x.ContainsKey("kode") ? x["kode"] : "");
                    if (jenis != "rek" && jenis != "produk") throw new InvalidOperationException("Jenis referensi tidak dikenal: " + jenis);
                    kunci = jenis == "rek" ? NormRek(kunci.Replace(" ", "")) : NormAwalan(kunci);
                    if (kunci.Length == 0) throw new InvalidOperationException(jenis == "rek" ? "Nomor rekening kosong." : "Awalan kode produk kosong.");
                    if (kode.Length == 0)
                    {
                        Database.Exec(con, "DELETE FROM ref_guna_manual WHERE jenis=@p0 AND kunci=@p1", jenis, kunci);
                        catatan.Add(jenis + " " + kunci + " dihapus");
                        continue;
                    }
                    Database.Exec(con, "INSERT OR REPLACE INTO ref_guna_manual(jenis,kunci,kode,alasan,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4,@p5)",
                        jenis, kunci, kode, alasan.Trim(), Environment.UserName, Database.Sekarang());
                    catatan.Add(jenis + " " + kunci + " → " + ModeHitung.NamaPenggunaan(kode));
                }
                SimpanInfo(con, "diubah", Database.Sekarang());
                Database.CatatAktivitas(con, "", "referensi-penggunaan", string.Join("; ", catatan.ToArray()) + " · dasar: " + alasan.Trim());
                tx.Commit();
            }
            LogProses.CatatPanel(DateTime.Today.ToString("yyyy-MM-dd"), "Referensi hapus buku", "Jenis penggunaan manual", LogProses.OK, LogProses.R()
                .Tambah("Perubahan", string.Join("; ", catatan.ToArray())).Tambah("Dasar", alasan.Trim()));
            return new Dictionary<string, object> { { "jumlah", catatan.Count } };
        }

        /// <summary>Simpan isian manual. item: jenis 'rek'|'produk', kunci, kodeKC (KC0600–KC1100 / DIKECUALIKAN / "" = hapus).</summary>
        public static Dictionary<string, object> SimpanManual(System.Collections.IList item, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Dasar / alasan pengisian referensi wajib diisi.");
            if (item == null || item.Count == 0) throw new InvalidOperationException("Tidak ada isian yang disimpan.");
            var valid = new HashSet<string>(DaftarKC, StringComparer.OrdinalIgnoreCase) { Dikecualikan };
            var catatan = new List<string>();
            Database.Cadangkan();
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                foreach (Dictionary<string, object> x in item)
                {
                    string jenis = Convert.ToString(x["jenis"]);
                    string kunci = Convert.ToString(x["kunci"] ?? "");
                    string kc = Convert.ToString(x.ContainsKey("kodeKC") ? x["kodeKC"] : "").Trim().ToUpperInvariant();
                    if (jenis != "rek" && jenis != "produk") throw new InvalidOperationException("Jenis referensi tidak dikenal: " + jenis);
                    kunci = jenis == "rek" ? NormRek(kunci.Replace(" ", "")) : NormAwalan(kunci);
                    if (kunci.Length == 0) throw new InvalidOperationException(jenis == "rek" ? "Nomor rekening kosong." : "Awalan kode produk kosong.");
                    if (kc.Length == 0)
                    {
                        Database.Exec(con, "DELETE FROM ref_kc_manual WHERE jenis=@p0 AND kunci=@p1", jenis, kunci);
                        catatan.Add(jenis + " " + kunci + " dihapus");
                        continue;
                    }
                    if (!valid.Contains(kc)) throw new InvalidOperationException("Kode KC tidak valid: " + kc);
                    Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_manual(jenis,kunci,kode_kc,alasan,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4,@p5)",
                        jenis, kunci, kc, alasan.Trim(), Environment.UserName, Database.Sekarang());
                    catatan.Add(jenis + " " + kunci + " → " + kc);
                }
                SimpanInfo(con, "diubah", Database.Sekarang());
                Database.CatatAktivitas(con, "", "referensi-kc", string.Join("; ", catatan.ToArray()) + " · dasar: " + alasan.Trim());
                tx.Commit();
            }
            LogProses.CatatPanel(DateTime.Today.ToString("yyyy-MM-dd"), "Referensi KC", "Isian manual", LogProses.OK, LogProses.R()
                .Tambah("Perubahan", string.Join("; ", catatan.ToArray())).Tambah("Dasar", alasan.Trim()));
            return new Dictionary<string, object> { { "jumlah", catatan.Count } };
        }

        public static Dictionary<string, object> AturPanjang(int panjang)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (Array.IndexOf(PanjangDipindai, panjang) < 0) throw new InvalidOperationException("Panjang awalan kode produk harus 2–5 digit.");
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                SimpanInfo(con, "panjang", panjang.ToString());
                SimpanInfo(con, "diubah", Database.Sekarang());
                Database.CatatAktivitas(con, "", "referensi-kc", "Panjang awalan kode produk = " + panjang + " digit");
                tx.Commit();
            }
            return new Dictionary<string, object> { { "panjang", panjang } };
        }

        /// <summary>Pindai ulang file tahunan Master!D60:D65 (tahun F60:F65) dari panel.</summary>
        public static Dictionary<string, object> PindaiUlang(Excel.Application app)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak ditemukan.");

            // Baca Master dengan aturan yang SAMA dengan perhitungan LGD ER (ParameterMaster.BacaLgdEr):
            // tahun F60:F65 boleh angka, teks angka, atau hasil rumus; path D60:D65; currYear = YEAR(EDATE(C4,-11)).
            ParameterCKPN prm = ParameterMaster.Baca(wb);
            List<string> masalah;
            if (prm.Error.TryGetValue("lgder", out masalah) && masalah.Count > 0)
                throw new InvalidOperationException("Parameter LGD ER di Master belum lengkap:\n- " + string.Join("\n- ", masalah.ToArray()));
            if (prm.LgdEr == null || string.IsNullOrEmpty(prm.LgdEr.FilePathsStr))
                throw new InvalidOperationException("Path file LGD ER (Master!D60:D65) tidak terbaca.");

            int currYear = prm.LgdEr.CurrYear;
            var files = new Dictionary<int, string>();
            foreach (var bagian in prm.LgdEr.FilePathsStr.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int i = bagian.IndexOf('=');
                int th;
                if (i > 0 && int.TryParse(bagian.Substring(0, i).Trim(), out th)) files[th] = bagian.Substring(i + 1).Trim();
            }
            if (files.Count == 0) throw new InvalidOperationException("Path file LGD ER (Master!D60:D65) tidak terbaca.");
            var h = Pindai(app, files);
            SimpanPindai(h, files, currYear);
            return new Dictionary<string, object> { { "rekening", h.Rek.Count }, { "file", h.File.Count } };
        }

        // =================================================================
        // Database
        // =================================================================
        private static void BacaManual(SQLiteConnection con, Peta p)
        {
            using (var cmd = Database.Cmd(con, "SELECT jenis,kunci,kode_kc FROM ref_kc_manual"))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    (rd.GetString(0) == "rek" ? p.ManualRek : p.ManualProduk)[rd.GetString(1)] = rd.GetString(2);
            // Tahap 8c: isian manual jenis penggunaan (tabel baru; database lama tanpa tabel → abaikan)
            try
            {
                using (var cmd = Database.Cmd(con, "SELECT jenis,kunci,kode FROM ref_guna_manual"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        (rd.GetString(0) == "rek" ? p.GunaManualRek : p.GunaManualProduk)[rd.GetString(1)] = rd.GetString(2);
            }
            catch { }
            string pj = BacaInfo(con, "panjang");
            int n;
            if (pj != null && int.TryParse(pj, out n) && Array.IndexOf(PanjangDipindai, n) >= 0) p.Panjang = n;
        }

        private static List<RekHB> BacaRek(SQLiteConnection con)
        {
            var hasil = new List<RekHB>();
            using (var cmd = Database.Cmd(con,
                "SELECT no_rek,rek_asli,cif,nama,jenis_instrumen,tgl_hb,tahun_hb,baki,tahun_file,kolom_m,kc_rekening,kc_rekening_tahun,sandi_hb,sandi_rekening,guna_rekening FROM ref_kc_hapusbuku"))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    hasil.Add(new RekHB
                    {
                        NoRek = rd.GetString(0), RekAsli = S(rd, 1), Cif = S(rd, 2), Nama = S(rd, 3), Produk = S(rd, 4), TglHB = S(rd, 5),
                        TahunHB = rd.IsDBNull(6) ? 0 : rd.GetInt32(6), Baki = rd.IsDBNull(7) ? 0 : rd.GetDouble(7), TahunFile = S(rd, 8),
                        KolomM = S(rd, 9), KcRekening = rd.IsDBNull(10) ? null : rd.GetString(10), KcRekeningTahun = rd.IsDBNull(11) ? 0 : rd.GetInt32(11),
                        SandiHB = S(rd, 12), SandiRekening = S(rd, 13), GunaRekening = S(rd, 14)
                    });
            return hasil;
        }

        /// <summary>Tahap 8c: hitungan rekening aktif per awalan kode produk per jenis penggunaan.</summary>
        private static Dictionary<string, int> BacaAggGuna(SQLiteConnection con)
        {
            var hasil = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var cmd = Database.Cmd(con, "SELECT panjang,awalan,kode,jumlah FROM ref_guna_produk"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) hasil[rd.GetInt32(0) + "|" + rd.GetString(1) + "|" + rd.GetString(2)] = rd.GetInt32(3);
            }
            catch { }
            return hasil;
        }

        private static Dictionary<string, int> BacaAgg(SQLiteConnection con)
        {
            var hasil = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = Database.Cmd(con, "SELECT panjang,awalan,kode_kc,jumlah FROM ref_kc_produk"))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read()) hasil[rd.GetInt32(0) + "|" + rd.GetString(1) + "|" + rd.GetString(2)] = rd.GetInt32(3);
            return hasil;
        }

        /// <summary>True bila hasil pindai tersimpan berasal dari file yang sama persis (path, ukuran, waktu ubah).</summary>
        private static bool CacheCocok(SQLiteConnection con, Dictionary<int, string> filePaths)
        {
            var info = BacaInfoJson(con, "pindai");
            if (info == null) return false;
            // Tahap 8: hasil pindai lama belum memuat sandi kantor → pindai ulang sekali
            object vp;
            if (!info.TryGetValue("versi", out vp) || Convert.ToInt32(vp) < VersiPindai) return false;
            var files = info["file"] as System.Collections.IList;
            if (files == null || files.Count != filePaths.Count) return false;
            foreach (Dictionary<string, object> f in files)
            {
                int t = Convert.ToInt32(f["tahun"]);
                string path;
                if (!filePaths.TryGetValue(t, out path)) return false;
                if (!string.Equals(SafeFull(path), SafeFull(Convert.ToString(f["path"])), StringComparison.OrdinalIgnoreCase)) return false;
                if (Convert.ToString(f["cap"]) != CapFile(path)) return false;
            }
            return true;
        }

        private static string BacaInfo(SQLiteConnection con, string kunci)
        {
            object v = Database.Scalar(con, "SELECT nilai FROM ref_kc_info WHERE kunci=@p0", kunci);
            return v == null || v is DBNull ? null : Convert.ToString(v);
        }

        private static Dictionary<string, object> BacaInfoJson(SQLiteConnection con, string kunci)
        {
            string s = BacaInfo(con, kunci);
            if (string.IsNullOrEmpty(s)) return null;
            try { return _json.DeserializeObject(s) as Dictionary<string, object>; } catch { return null; }
        }

        private static void SimpanInfo(SQLiteConnection con, string kunci, string nilai)
        {
            Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_info(kunci,nilai) VALUES(@p0,@p1)", kunci, nilai);
        }

        // =================================================================
        // Util
        // =================================================================
        /// <summary>Sama dengan LGDExpectedRecoveries.NormNoRek.</summary>
        internal static string NormRek(string raw)
        {
            if (raw == null) return "";
            string s = raw.Trim();
            if (s.StartsWith("'")) s = s.Substring(1).Trim();
            if (s.Length == 0) return "";
            foreach (char c in s) if (!char.IsDigit(c)) return s;
            string t = s.TrimStart('0');
            return t.Length == 0 ? "0" : t;
        }

        /// <summary>Awalan kode produk dari isian petugas: tanpa apostrof/spasi; angka tanpa nol di depan.</summary>
        private static string NormAwalan(string raw)
        {
            string s = (raw ?? "").Trim().TrimStart('\'').Replace(" ", "");
            foreach (char c in s) if (!char.IsDigit(c)) return s;
            return s.TrimStart('0');
        }

        /// <summary>Sama dengan LGDExpectedRecoveries.NormKodeKC.</summary>
        internal static string NormKC(string raw)
        {
            if (raw == null) return "";
            string s = raw.Trim().ToUpperInvariant();
            if (s.Length == 0 || s == "FALSE" || s == "TRUE" || s == "0" || s.StartsWith("#")) return "";
            return s;
        }

        private static string Awalan(string norm, int panjang)
        {
            return norm.Length >= panjang ? norm.Substring(0, panjang) : norm;
        }

        internal static bool DalamCohort(RekHB r, int currYear)
        {
            return r.TahunHB >= currYear - JendelaCohort + 1 && r.TahunHB <= currYear;
        }

        /// <summary>Tahap 8: lingkup mode per kantor = KC + saringan, mis. "KC0600,…,KC1100|kantor:001,002".</summary>
        private static string KunciLingkup(IEnumerable<string> kc, FilterLingkup f)
        {
            string k = KunciLingkup(kc);
            return f == null ? k : k + "|" + f.ModeId + ":" + f.Kode;
        }

        /// <summary>Sandi kantor kanonik; kosong bila sel kosong atau "0"/"000" (kode kantor tidak diisi).</summary>
        internal static string SandiSah(object v)
        {
            string s = ModeHitung.NormSandi(v);
            return s.Length == 0 || s.TrimStart('0').Length == 0 ? "" : s;
        }

        private static string KunciLingkup(IEnumerable<string> kc)
        {
            var l = new List<string>();
            foreach (var k in kc) if (!string.IsNullOrWhiteSpace(k)) l.Add(k.Trim().ToUpperInvariant());
            l.Sort(StringComparer.Ordinal);
            return string.Join(",", l.ToArray());
        }

        private static int TahunDari(string tgl, int cadangan)
        {
            DateTime t;
            if (tgl != null && tgl.Length >= 10 && DateTime.TryParse(tgl.Substring(0, 10), System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out t)) return t.Year;
            return cadangan;
        }

        internal static string CapFile(string path)
        {
            try
            {
                var fi = new FileInfo(path);
                return fi.Exists ? fi.Length + "|" + fi.LastWriteTimeUtc.Ticks : "";
            }
            catch { return ""; }
        }

        private static string SafeFull(string p)
        {
            try { return Path.GetFullPath(p); } catch { return p ?? ""; }
        }

        private static object[,] BacaArea(Excel.Worksheet ws)
        {
            Excel.Range used = ws.UsedRange;
            int baris = used.Row + used.Rows.Count - 1, kolom = used.Column + used.Columns.Count - 1;
            if (baris < 2) return null;
            if (kolom < 13) kolom = 13;
            return ((Excel.Range)ws.Range[ws.Cells[1, 1], ws.Cells[baris, kolom]]).Value2 as object[,];
        }

        private static object Sel(object[,] d, int baris, int kolom)
        {
            if (baris > d.GetUpperBound(0) || kolom > d.GetUpperBound(1)) return null;
            return d[baris, kolom];
        }

        private static string Teks(object v) { return v == null ? "" : ProfilTemplate.Norm(v); }

        private static double Angka(object v)
        {
            if (v is double) return (double)v;
            double d;
            return double.TryParse(Convert.ToString(v ?? ""), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out d) ? d : 0;
        }

        private static string S(SQLiteDataReader rd, int i) { return rd.IsDBNull(i) ? null : Convert.ToString(rd.GetValue(i)); }
    }
}
