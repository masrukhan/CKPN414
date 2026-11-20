using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Mode per produk — kode produk pembiayaan yang dibaca dari nomor rekening/kontrak.
    ///
    /// Nomor kontrak = nomor rekening: kolom J sheet KC0600–KC1100 (terisi di baris pembiayaan DAN baris agunan,
    /// sehingga setiap baris bisa disaring langsung tanpa melihat baris di atasnya) dan kolom H sheet KC2900.
    /// Kode produk = potongan nomor rekening menurut <see cref="Pola"/> (posisi awal + panjang, bawaan 1 + 3:
    /// 4510102454 → 451). Nol di depan diabaikan (sel angka tidak menyimpannya), sama dengan pencocokan rekening
    /// di seluruh aplikasi (ReferensiKC.NormRek).
    ///
    /// Hasil pindai template disimpan di database utama (tabel kode_produk) dan menjadi bahan susunan grup:
    ///   • kode dengan rekening di bulan berjalan → wajib masuk grup;
    ///   • kode yang hanya ada di data hapus buku (KC2900, produk lama) → juga wajib dipetakan user ke salah satu
    ///     grup, karena LGD Expected Recoveries memakai rekening hapus buku per grup.
    /// Pipeline menolak menghitung bila ada kode di template (pembiayaan maupun hapus buku) yang belum masuk grup,
    /// supaya tidak ada saldo yang diam-diam hilang dari konsolidasi.
    /// </summary>
    internal static class KodeProduk
    {
        private const int BarisDataKC = 5;     // KC0600–KC1100: header baris 2–4
        private const int BarisDataHB = 3;     // KC2900: header baris 1–2
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // ================================================================
        // Pola identifikasi (posisi awal 1-based + panjang), disimpan di ref_kc_info 'produk_pola'
        // ================================================================
        internal sealed class PolaProduk
        {
            public int Posisi = 1, Panjang = 3;
            public string Teks { get { return Posisi + "," + Panjang; } }
        }

        private static PolaProduk _pola;
        private static DateTime _waktuPola;

        public static PolaProduk Pola
        {
            get
            {
                // Dibaca sangat sering (setiap baris) → cache 30 detik
                if (_pola != null && (DateTime.Now - _waktuPola).TotalSeconds < 30) return _pola;
                var p = new PolaProduk();
                try
                {
                    if (Database.AdaBersama)
                        using (var con = Database.BukaBersama(false))
                        {
                            object v = Database.Scalar(con, "SELECT nilai FROM ref_kc_info WHERE kunci='produk_pola'");
                            if (v != null && !(v is DBNull))
                            {
                                var d = _json.Deserialize<Dictionary<string, object>>(Convert.ToString(v));
                                p.Posisi = Math.Max(1, Convert.ToInt32(d["posisi"]));
                                p.Panjang = Math.Max(1, Convert.ToInt32(d["panjang"]));
                            }
                        }
                }
                catch (Exception ex) { CatatanLog.Tulis("Baca pola kode produk: " + ex.Message); }
                _pola = p; _waktuPola = DateTime.Now;
                return p;
            }
        }

        /// <summary>
        /// Ubah pola. Kode lama kehilangan arti → daftar kode produk dikosongkan (pindai ulang); susunan grup
        /// yang sudah ada perlu disusun ulang (panel memberi peringatan).
        /// </summary>
        /// <returns>true bila pola berubah</returns>
        public static bool AturPola(int posisi, int panjang)
        {
            if (posisi < 1 || posisi > 20) throw new InvalidOperationException("Posisi awal kode produk harus 1–20.");
            if (panjang < 1 || panjang > 10) throw new InvalidOperationException("Panjang kode produk harus 1–10 digit.");
            var lama = Pola;
            if (lama.Posisi == posisi && lama.Panjang == panjang) return false;
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_info(kunci,nilai) VALUES('produk_pola',@p0)",
                    _json.Serialize(new Dictionary<string, object> { { "posisi", posisi }, { "panjang", panjang } }));
                Database.Exec(con, "DELETE FROM kode_produk");
                Database.Exec(con, "DELETE FROM ref_kc_info WHERE kunci='produk_pindai'");
                Database.CatatAktivitas(con, "", "pola-produk", lama.Teks + " → " + posisi + "," + panjang);
                tx.Commit();
            }
            _pola = new PolaProduk { Posisi = posisi, Panjang = panjang }; _waktuPola = DateTime.Now;
            CatatanLog.Tulis("Pola kode produk diganti " + lama.Teks + " → " + _pola.Teks + " oleh " + Environment.UserName);
            return true;
        }

        // ================================================================
        // Kode dari isian sel
        // ================================================================
        /// <summary>
        /// Kode produk dari isian sel nomor rekening (teks atau angka). "" bila kosong atau lebih pendek dari pola.
        /// 4510102454 / "4510102454" / "'4510102454" → "451" (pola 1,3).
        /// </summary>
        public static string Dari(object v)
        {
            string s = SumberLingkup.NormRek(v);
            if (s.Length == 0) return "";
            var p = Pola;
            if (s.Length < p.Posisi - 1 + p.Panjang) return "";
            return s.Substring(p.Posisi - 1, p.Panjang).ToUpperInvariant();
        }

        /// <summary>Kode produk sebagai bagian kode grup ("451", " 456 ", "'013") — tidak dipotong lagi.</summary>
        public static string NormKode(string kode)
        {
            return (kode ?? "").Trim().Trim('\'', '"').Replace(" ", "").ToUpperInvariant();
        }

        // ================================================================
        // Pindai template
        // ================================================================
        internal class HasilPindai
        {
            public string File, Cap, Tanggal, Pola;
            public readonly Dictionary<string, Unit> Kode = new Dictionary<string, Unit>(StringComparer.OrdinalIgnoreCase);
            /// <summary>Baris pembiayaan yang nomor rekeningnya kosong / terlalu pendek untuk pola.</summary>
            public int Kosong;
            /// <summary>Rekening hapus buku yang nomor rekeningnya kosong / terlalu pendek.</summary>
            public int HbKosong;
        }

        internal class Unit
        {
            public readonly Dictionary<string, int> Rincian = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            public readonly HashSet<string> Rek = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public readonly HashSet<string> RekHb = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public double BakiHb;
            /// <summary>Rincian rekening hapus buku: {rek, cif, tanggal HB, baki} — daftar tindak lanjut di panel.</summary>
            public readonly List<Dictionary<string, object>> HbRinci = new List<Dictionary<string, object>>();
            public string Contoh = "";
            public int Rekening { get { return Rek.Count; } }
        }

        /// <summary>
        /// Baca kolom J (rekening) + D (nama) sheet KC0600–KC1100 dan kolom H (rekening) + L (baki debet) KC2900.
        /// Rekening dihitung dari baris pembiayaan (nama terisi); baris agunan tidak menambah jumlah.
        /// Template dibuka read-only bila belum terbuka.
        /// </summary>
        public static HasilPindai Pindai(Excel.Application app, string path, string tanggal)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                throw new FileNotFoundException("File template APOLLO (Master!D14) tidak ditemukan: " + path);
            var h = new HasilPindai { File = path, Cap = SandiKantor.Cap(path), Tanggal = tanggal ?? "", Pola = Pola.Teks };

            Excel.Workbook wb = null;
            bool dibukaDiSini = false, layar = true, alert = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; app.ScreenUpdating = false; app.DisplayAlerts = false; } catch { }
            try
            {
                string full = Path.GetFullPath(path);
                foreach (Excel.Workbook w in app.Workbooks)
                    if (string.Equals(w.FullName, full, StringComparison.OrdinalIgnoreCase)) { wb = w; break; }
                if (wb == null) { wb = app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true); dibukaDiSini = true; }

                // ---- pembiayaan bulan berjalan ----
                foreach (var kc in ModeHitung.SemuaKC)
                {
                    Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kc);
                    if (ws == null) continue;
                    int akhir = BarisAkhir(ws, 10);
                    if (akhir < BarisDataKC) continue;
                    // D (4) s.d. J (10)
                    object[,] d = ((Excel.Range)ws.Range[ws.Cells[BarisDataKC, 4], ws.Cells[akhir, 10]]).Value2 as object[,];
                    if (d == null) continue;
                    int c0 = d.GetLowerBound(1);
                    for (int r = d.GetLowerBound(0); r <= d.GetUpperBound(0); r++)
                    {
                        object nama = d[r, c0];
                        if (nama == null || Convert.ToString(nama).Trim().Length == 0) continue;   // baris agunan / kosong
                        object rekSel = d[r, c0 + 6];
                        string rek = SumberLingkup.NormRek(rekSel);
                        string kode = Dari(rekSel);
                        if (kode.Length == 0) { if (rek.Length > 0 || nama != null) h.Kosong++; continue; }
                        Unit u = AmbilUnit(h, kode);
                        if (u.Rek.Add(rek)) u.Rincian[kc] = (u.Rincian.ContainsKey(kc) ? u.Rincian[kc] : 0) + 1;
                        if (u.Contoh.Length == 0) u.Contoh = rek;
                    }
                }

                // ---- hapus buku (KC2900): kode produk lama ----
                Excel.Worksheet hb = ParameterMaster.CariSheet(wb, "KC2900");
                if (hb != null)
                {
                    int akhir = BarisAkhir(hb, 8);
                    if (akhir >= BarisDataHB)
                    {
                        // C (3, CIF) s.d. L (12): H = rekening, I = tanggal hapus buku, L = baki debet
                        object[,] d = ((Excel.Range)hb.Range[hb.Cells[BarisDataHB, 3], hb.Cells[akhir, 12]]).Value2 as object[,];
                        if (d != null)
                        {
                            int c0 = d.GetLowerBound(1);
                            for (int r = d.GetLowerBound(0); r <= d.GetUpperBound(0); r++)
                            {
                                object rekSel = d[r, c0 + 5];
                                string rek = SumberLingkup.NormRek(rekSel);
                                if (rek.Length == 0) continue;
                                string kode = Dari(rekSel);
                                if (kode.Length == 0) { h.HbKosong++; continue; }
                                Unit u = AmbilUnit(h, kode);
                                double baki = Angka(d[r, c0 + 9]);
                                if (u.RekHb.Add(rek))
                                {
                                    u.BakiHb += baki;
                                    u.HbRinci.Add(new Dictionary<string, object>
                                    {
                                        { "rek", rek }, { "cif", SumberLingkup.NormRek(d[r, c0]) },
                                        { "tanggal", Convert.ToString(d[r, c0 + 6], CultureInfo.InvariantCulture) ?? "" }, { "baki", baki }
                                    });
                                }
                                if (u.Contoh.Length == 0) u.Contoh = rek;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (dibukaDiSini && wb != null) try { wb.Close(false); } catch { }
                try { app.DisplayAlerts = alert; app.ScreenUpdating = layar; } catch { }
            }
            return h;
        }

        private static Unit AmbilUnit(HasilPindai h, string kode)
        {
            Unit u;
            if (!h.Kode.TryGetValue(kode, out u)) { u = new Unit(); h.Kode[kode] = u; }
            return u;
        }

        private static int BarisAkhir(Excel.Worksheet ws, int kolom)
        {
            try { return ((Excel.Range)ws.Cells[ws.Rows.Count, kolom]).End[Excel.XlDirection.xlUp].Row; }
            catch { return 0; }
        }

        private static double Angka(object v)
        {
            if (v is double) return (double)v;
            double d;
            return v != null && double.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? d : 0;
        }

        /// <summary>Simpan hasil pindai (nama produk yang sudah diisi user dipertahankan).</summary>
        public static void Simpan(HasilPindai h)
        {
            string info;
            if (!Database.BolehMenulis(out info)) return;
            string file = Path.GetFileName(h.File), sekarang = Database.Sekarang();
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                foreach (var kv in h.Kode)
                {
                    var u = kv.Value;
                    Database.Exec(con,
                        "INSERT INTO kode_produk(kode,nama,rekening,rincian,os,hb,baki_hb,contoh,tanggal,file,waktu) " +
                        "VALUES(@p0,NULL,@p1,@p2,NULL,@p3,@p4,@p5,@p6,@p7,@p8) " +
                        "ON CONFLICT(kode) DO UPDATE SET rekening=@p1, rincian=@p2, hb=@p3, baki_hb=@p4, contoh=@p5, tanggal=@p6, file=@p7, waktu=@p8",
                        kv.Key, u.Rekening, _json.Serialize(u.Rincian), u.RekHb.Count, u.BakiHb, u.Contoh, h.Tanggal, file, sekarang);
                }
                // Kode yang tidak lagi muncul tetap disimpan (mungkin dipakai susunan tahun lalu), jumlahnya dinolkan
                var hilang = new List<string>();
                using (var cmd = Database.Cmd(con, "SELECT kode FROM kode_produk"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) if (!h.Kode.ContainsKey(rd.GetString(0))) hilang.Add(rd.GetString(0));
                foreach (var k in hilang)
                    Database.Exec(con, "UPDATE kode_produk SET rekening=0, rincian='{}', hb=0, baki_hb=0, tanggal=@p1, file=@p2, waktu=@p3 WHERE kode=@p0",
                        k, h.Tanggal, file, sekarang);

                var pemb = new List<string>(); var lama = new List<string>();
                foreach (var kv in h.Kode)
                {
                    if (kv.Value.Rekening > 0) pemb.Add(kv.Key);
                    else if (kv.Value.RekHb.Count > 0) lama.Add(kv.Key);
                }
                // Rincian rekening hapus buku per kode (untuk daftar tindak lanjut produk yang belum masuk grup)
                var rinciHb = new Dictionary<string, object>();
                foreach (var kv in h.Kode) if (kv.Value.HbRinci.Count > 0) rinciHb[kv.Key] = kv.Value.HbRinci;
                Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_info(kunci,nilai) VALUES('produk_hb',@p0)", _json.Serialize(rinciHb));
                Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_info(kunci,nilai) VALUES('produk_pindai',@p0)", _json.Serialize(new Dictionary<string, object>
                {
                    { "file", h.File }, { "cap", h.Cap }, { "pola", h.Pola }, { "tanggal", h.Tanggal }, { "waktu", sekarang },
                    { "pengguna", Environment.UserName }, { "kode", pemb }, { "hapusBuku", lama },
                    { "kosong", h.Kosong }, { "hbKosong", h.HbKosong }
                }));
                tx.Commit();
            }
        }

        /// <summary>
        /// Kode produk di template periode (pindai ulang bila file atau pola berubah):
        /// kodePembiayaan = kode dengan rekening pembiayaan; kodeHapusBuku = kode yang HANYA ada di KC2900.
        /// </summary>
        public static void KodeTemplate(Excel.Application app, string path, string tanggal,
                                        out List<string> kodePembiayaan, out List<string> kodeHapusBuku, out int kosong)
        {
            kodePembiayaan = new List<string>(); kodeHapusBuku = new List<string>(); kosong = 0;
            var info = InfoPindai();
            if (info != null && Convert.ToString(Ambil(info, "file")) == path && Convert.ToString(Ambil(info, "cap")) == SandiKantor.Cap(path) &&
                Convert.ToString(Ambil(info, "pola")) == Pola.Teks)
            {
                Isi(kodePembiayaan, Ambil(info, "kode"));
                Isi(kodeHapusBuku, Ambil(info, "hapusBuku"));
                object k = Ambil(info, "kosong");
                if (k != null) kosong = Convert.ToInt32(k);
                return;
            }
            var h = Pindai(app, path, tanggal);
            try { Simpan(h); } catch (Exception ex) { CatatanLog.Tulis("Simpan hasil pindai kode produk: " + ex.Message); }
            kosong = h.Kosong;
            foreach (var kv in h.Kode)
            {
                if (kv.Value.Rekening > 0) kodePembiayaan.Add(kv.Key);
                else if (kv.Value.RekHb.Count > 0) kodeHapusBuku.Add(kv.Key);
            }
            kodePembiayaan.Sort(ModeHitung.BandingSandi);
            kodeHapusBuku.Sort(ModeHitung.BandingSandi);
        }

        private static void Isi(List<string> l, object arr)
        {
            var e = arr as System.Collections.IEnumerable;
            if (e != null) foreach (var x in e) l.Add(Convert.ToString(x));
        }

        public static Dictionary<string, object> InfoPindai()
        {
            try
            {
                if (!Database.AdaBersama) return null;
                using (var con = Database.BukaBersama(false))
                {
                    object v = Database.Scalar(con, "SELECT nilai FROM ref_kc_info WHERE kunci='produk_pindai'");
                    if (v == null || v is DBNull) return null;
                    return _json.Deserialize<Dictionary<string, object>>(Convert.ToString(v));
                }
            }
            catch { return null; }
        }

        /// <summary>Daftar kode produk tersimpan untuk editor susunan grup (urut numerik).</summary>
        public static List<Dictionary<string, object>> Daftar()
        {
            var hasil = new List<Dictionary<string, object>>();
            try
            {
                if (!Database.AdaBersama) return hasil;
                using (var con = Database.BukaBersama(false))
                using (var cmd = Database.Cmd(con, "SELECT kode, nama, rekening, rincian, hb, baki_hb, contoh, tanggal, file FROM kode_produk"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                    {
                        Dictionary<string, object> rinci = null;
                        try { rinci = _json.Deserialize<Dictionary<string, object>>(Str(rd[3]) == "" ? "{}" : Str(rd[3])); } catch { }
                        int rek = rd[2] is DBNull ? 0 : Convert.ToInt32(rd[2]);
                        int hb = rd[4] is DBNull ? 0 : Convert.ToInt32(rd[4]);
                        hasil.Add(new Dictionary<string, object>
                        {
                            { "kode", rd.GetString(0) }, { "nama", Str(rd[1]) }, { "rekening", rek },
                            { "rincian", rinci ?? new Dictionary<string, object>() },
                            { "hb", hb }, { "bakiHb", rd[5] is DBNull ? 0.0 : Convert.ToDouble(rd[5]) },
                            { "contoh", Str(rd[6]) }, { "tanggal", Str(rd[7]) }, { "file", Str(rd[8]) },
                            // Produk berjalan ATAU produk lama yang masih ada di data hapus buku → wajib dipetakan ke grup
                            { "wajibIsi", rek > 0 || hb > 0 }, { "lama", rek == 0 && hb > 0 }
                        });
                    }
            }
            catch (Exception ex) { CatatanLog.Tulis("Baca daftar kode produk: " + ex.Message); }
            hasil.Sort((a, b) => ModeHitung.BandingSandi((string)a["kode"], (string)b["kode"]));
            return hasil;
        }

        /// <summary>Rekening hapus buku (KC2900 template terakhir) untuk satu kode produk.</summary>
        public static List<object> RekeningHapusBuku(string kode)
        {
            var hasil = new List<object>();
            try
            {
                if (!Database.AdaBersama) return hasil;
                using (var con = Database.BukaBersama(false))
                {
                    object v = Database.Scalar(con, "SELECT nilai FROM ref_kc_info WHERE kunci='produk_hb'");
                    if (v == null || v is DBNull) return hasil;
                    var d = _json.Deserialize<Dictionary<string, object>>(Convert.ToString(v));
                    object l;
                    if (d != null && d.TryGetValue(NormKode(kode), out l) && l is System.Collections.IEnumerable)
                        foreach (var x in (System.Collections.IEnumerable)l) hasil.Add(x);
                }
            }
            catch (Exception ex) { CatatanLog.Tulis("Baca rekening hapus buku kode produk: " + ex.Message); }
            return hasil;
        }

        /// <summary>Nama produk (opsional, untuk tampilan). Kosong = hapus nama.</summary>
        public static void SimpanNama(Dictionary<string, string> nama)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (nama == null || nama.Count == 0) return;
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                foreach (var kv in nama)
                {
                    string k = NormKode(kv.Key);
                    if (k.Length == 0) continue;
                    string n = (kv.Value ?? "").Trim();
                    if (n.Length > 80) n = n.Substring(0, 80);
                    Database.Exec(con,
                        "INSERT INTO kode_produk(kode,nama,rekening,waktu) VALUES(@p0,@p1,0,@p2) ON CONFLICT(kode) DO UPDATE SET nama=@p1",
                        k, n.Length == 0 ? null : n, Database.Sekarang());
                }
                tx.Commit();
            }
        }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
