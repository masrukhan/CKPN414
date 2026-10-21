using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Tahap 8 — daftar sandi lokasi kantor dari template APOLLO (kolom B sheet KC0600–KC1100).
    ///
    /// Hasil pindai disimpan di database utama (tabel sandi_kantor) dan menjadi bahan susunan grup
    /// mode per kantor. Setiap kali grup dihitung pada mode kantor, template periode itu diperiksa:
    /// sandi yang ada di template tetapi tidak masuk grup mana pun membuat perhitungan ditolak,
    /// supaya tidak ada pembiayaan yang terlewat dari konsolidasi.
    ///
    /// Hanya baris pembiayaan (nama nasabah terisi) yang dihitung sebagai rekening; baris agunan
    /// (nama kosong) tidak menambah jumlah.
    /// </summary>
    internal static class SandiKantor
    {
        private const int BarisDataPertama = 5;   // header template baris 2–4
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        internal class HasilPindai
        {
            public string File, Cap, Tanggal;
            /// <summary>sandi → (KC → jumlah rekening)</summary>
            public Dictionary<string, Dictionary<string, int>> Sandi =
                new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            public int Kosong;   // baris pembiayaan tanpa sandi kantor
        }

        /// <summary>Penanda versi file (ukuran + waktu ubah) — pindai ulang hanya bila berubah.</summary>
        public static string Cap(string path)
        {
            try { var fi = new FileInfo(path); return fi.Length + "@" + fi.LastWriteTimeUtc.Ticks; }
            catch { return ""; }
        }

        /// <summary>Baca kolom B (sandi) dan D (nama) sheet KC0600–KC1100. Template dibuka read-only bila belum terbuka.</summary>
        public static HasilPindai Pindai(Excel.Application app, string path, string tanggal)
        {
            if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path))
                throw new FileNotFoundException("File template APOLLO (Master!D14) tidak ditemukan: " + path);
            var h = new HasilPindai { File = path, Cap = Cap(path), Tanggal = tanggal ?? "" };

            Excel.Workbook wb = null;
            bool dibukaDiSini = false, layar = true, alert = true;
            try { layar = app.ScreenUpdating; alert = app.DisplayAlerts; app.ScreenUpdating = false; app.DisplayAlerts = false; } catch { }
            try
            {
                string full = Path.GetFullPath(path);
                foreach (Excel.Workbook w in app.Workbooks)
                    if (string.Equals(w.FullName, full, StringComparison.OrdinalIgnoreCase)) { wb = w; break; }
                if (wb == null) { wb = app.Workbooks.Open(path, UpdateLinks: 0, ReadOnly: true); dibukaDiSini = true; }

                foreach (var kc in ModeHitung.SemuaKC)
                {
                    Excel.Worksheet ws = ParameterMaster.CariSheet(wb, kc);
                    if (ws == null) continue;
                    Excel.Range used = ws.UsedRange;
                    int akhir = used.Row + used.Rows.Count - 1;
                    if (akhir < BarisDataPertama) continue;
                    object[,] d = ((Excel.Range)ws.Range[ws.Cells[BarisDataPertama, 2], ws.Cells[akhir, 4]]).Value2 as object[,];
                    if (d == null) continue;
                    for (int r = d.GetLowerBound(0); r <= d.GetUpperBound(0); r++)
                    {
                        object nama = d[r, d.GetLowerBound(1) + 2];
                        if (nama == null || Convert.ToString(nama).Trim().Length == 0) continue;   // baris agunan / kosong
                        string sandi = ModeHitung.NormSandi(d[r, d.GetLowerBound(1)]);
                        if (sandi.Length == 0) { h.Kosong++; continue; }
                        Dictionary<string, int> per;
                        if (!h.Sandi.TryGetValue(sandi, out per)) { per = new Dictionary<string, int>(); h.Sandi[sandi] = per; }
                        per[kc] = (per.ContainsKey(kc) ? per[kc] : 0) + 1;
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

        /// <summary>Simpan hasil pindai ke database utama (nama kantor yang sudah diisi tetap dipertahankan).</summary>
        public static void Simpan(HasilPindai h)
        {
            string info;
            if (!Database.BolehMenulis(out info)) return;
            using (var con = Database.BukaBersama(true))
            using (var tx = con.BeginTransaction())
            {
                string sekarang = Database.Sekarang();
                foreach (var kv in h.Sandi)
                {
                    int n = 0;
                    foreach (var x in kv.Value.Values) n += x;
                    Database.Exec(con,
                        "INSERT INTO sandi_kantor(sandi,nama,rekening,rincian,tanggal,file,waktu) VALUES(@p0,NULL,@p1,@p2,@p3,@p4,@p5) " +
                        "ON CONFLICT(sandi) DO UPDATE SET rekening=@p1, rincian=@p2, tanggal=@p3, file=@p4, waktu=@p5",
                        kv.Key, n, _json.Serialize(kv.Value), h.Tanggal, Path.GetFileName(h.File), sekarang);
                }
                // Sandi yang tidak lagi muncul tetap disimpan (mungkin dipakai susunan tahun lalu), jumlahnya dinolkan
                var ada = new List<string>(h.Sandi.Keys);
                using (var cmd = Database.Cmd(con, "SELECT sandi FROM sandi_kantor"))
                using (var rd = cmd.ExecuteReader())
                {
                    var hilang = new List<string>();
                    while (rd.Read()) if (!h.Sandi.ContainsKey(rd.GetString(0))) hilang.Add(rd.GetString(0));
                    rd.Close();
                    foreach (var s in hilang)
                        Database.Exec(con, "UPDATE sandi_kantor SET rekening=0, rincian='{}', tanggal=@p1, file=@p2, waktu=@p3 WHERE sandi=@p0",
                            s, h.Tanggal, Path.GetFileName(h.File), sekarang);
                }
                Database.Exec(con, "INSERT OR REPLACE INTO ref_kc_info(kunci,nilai) VALUES('sandi_pindai',@p0)", _json.Serialize(new Dictionary<string, object>
                {
                    { "file", h.File }, { "cap", h.Cap }, { "tanggal", h.Tanggal }, { "waktu", sekarang },
                    { "pengguna", Environment.UserName }, { "sandi", ada }, { "kosong", h.Kosong }
                }));
                tx.Commit();
            }
        }

        /// <summary>Pindai template bila belum pernah / file berubah, lalu kembalikan daftar sandi di template itu.</summary>
        public static List<string> SandiTemplate(Excel.Application app, string path, string tanggal, out int kosong)
        {
            kosong = 0;
            var info = InfoPindai();
            if (info != null && Convert.ToString(Ambil(info, "file")) == path && Convert.ToString(Ambil(info, "cap")) == Cap(path))
            {
                var l = new List<string>();
                var arr = Ambil(info, "sandi") as System.Collections.IEnumerable;
                if (arr != null) foreach (var x in arr) l.Add(Convert.ToString(x));
                object k = Ambil(info, "kosong");
                if (k != null) kosong = Convert.ToInt32(k);
                return l;
            }
            var h = Pindai(app, path, tanggal);
            try { Simpan(h); } catch (Exception ex) { CatatanLog.Tulis("Simpan hasil pindai sandi kantor: " + ex.Message); }
            kosong = h.Kosong;
            var hasil = new List<string>(h.Sandi.Keys);
            hasil.Sort(ModeHitung.BandingSandi);
            return hasil;
        }

        public static Dictionary<string, object> InfoPindai()
        {
            try
            {
                if (!Database.AdaBersama) return null;
                using (var con = Database.BukaBersama(false))
                {
                    object v = Database.Scalar(con, "SELECT nilai FROM ref_kc_info WHERE kunci='sandi_pindai'");
                    if (v == null || v is DBNull) return null;
                    return _json.Deserialize<Dictionary<string, object>>(Convert.ToString(v));
                }
            }
            catch { return null; }
        }

        /// <summary>Daftar sandi tersimpan untuk editor susunan grup (urut numerik).</summary>
        public static List<Dictionary<string, object>> Daftar()
        {
            var hasil = new List<Dictionary<string, object>>();
            try
            {
                if (!Database.AdaBersama) return hasil;
                using (var con = Database.BukaBersama(false))
                using (var cmd = Database.Cmd(con, "SELECT sandi, nama, rekening, rincian, tanggal, file FROM sandi_kantor"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                    {
                        Dictionary<string, object> rinci = null;
                        try { rinci = _json.Deserialize<Dictionary<string, object>>(Str(rd[3]) == "" ? "{}" : Str(rd[3])); } catch { }
                        hasil.Add(new Dictionary<string, object>
                        {
                            { "kode", rd.GetString(0) }, { "nama", Str(rd[1]) },
                            { "rekening", rd[2] is DBNull ? 0 : Convert.ToInt32(rd[2]) },
                            { "rincian", rinci ?? new Dictionary<string, object>() },
                            { "tanggal", Str(rd[4]) }, { "file", Str(rd[5]) }
                        });
                    }
            }
            catch (Exception ex) { CatatanLog.Tulis("Baca daftar sandi kantor: " + ex.Message); }
            hasil.Sort((a, b) => ModeHitung.BandingSandi((string)a["kode"], (string)b["kode"]));
            return hasil;
        }

        /// <summary>Nama kantor (opsional, untuk tampilan). Kosong = hapus nama.</summary>
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
                    string s = ModeHitung.NormSandi(kv.Key);
                    if (s.Length == 0) continue;
                    string n = (kv.Value ?? "").Trim();
                    if (n.Length > 80) n = n.Substring(0, 80);
                    Database.Exec(con,
                        "INSERT INTO sandi_kantor(sandi,nama,rekening,waktu) VALUES(@p0,@p1,0,@p2) ON CONFLICT(sandi) DO UPDATE SET nama=@p1",
                        s, n.Length == 0 ? null : n, Database.Sekarang());
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
