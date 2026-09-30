using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Alur kerja satu periode laporan (Tahap 3b):
    ///
    ///   1. Susunan grup      — kode segmen KC0600..KC1100 dikelompokkan (setahun sekali, sesuai SOP)
    ///   2. Hitung & simpan   — tiap grup dihitung, direview/diedit, lalu disimpan (bisa berulang → versi)
    ///   3. Konsolidasi       — setelah SEMUA grup tersimpan: CKPN (satu metode untuk semua grup,
    ///                          ditetapkan di susunan tahunan) vs PPKA OJK + usulan jurnal CKPN
    ///   4. Final             — periode dikunci (tidak bisa simpan/hapus versi sampai dibuka kuncinya)
    ///
    /// Tahap 3c: metode konsolidasi (Net Flow / Migration) ditetapkan SEKALI per tahun
    /// bersama susunan grup (tabel susunan_tahun).
    /// Tahap 3d: usulan jurnal = satu jurnal atas selisih total CKPN (termasuk ABA) vs PPKA.
    ///
    /// Semua data tulis melewati Database.BolehMenulis (daftar pengirim).
    /// </summary>
    internal static class Periode
    {
        public static readonly string[] SemuaKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // =================================================================
        // KONTEKS WORKBOOK (lokal per user): grup apa yang terakhir dihitung
        // di PC ini, dan apakah sudah disimpan.
        // =================================================================
        public static void CatatHitung(string tanggal, string kodeKC, IEnumerable<string> langkah)
        {
            var k = BacaKonteks() ?? new Dictionary<string, object>();
            k["tanggal"] = tanggal;
            k["kodeKC"] = kodeKC;
            k["waktuHitung"] = Database.Sekarang();
            k["langkah"] = new List<string>(langkah);
            k["tersimpanVersi"] = null;
            k["waktuSimpan"] = null;
            SimpanKonteks(k);
        }

        public static void CatatSimpan(string tanggal, string kodeKC, int versi)
        {
            var k = BacaKonteks() ?? new Dictionary<string, object>();
            if (Convert.ToString(k.ContainsKey("tanggal") ? k["tanggal"] : "") != tanggal ||
                Convert.ToString(k.ContainsKey("kodeKC") ? k["kodeKC"] : "") != kodeKC)
            {
                k["tanggal"] = tanggal; k["kodeKC"] = kodeKC;
                k["waktuHitung"] = null; k["langkah"] = new List<string>();
            }
            k["tersimpanVersi"] = versi;
            k["waktuSimpan"] = Database.Sekarang();
            SimpanKonteks(k);
        }

        /// <summary>
        /// Setelah data periode dihapus dari database (Tahap 3d): tanda "sudah disimpan"
        /// di konteks lokal dikosongkan. tanggal null = semua periode.
        /// </summary>
        public static void ResetKonteksSimpan(string tanggal)
        {
            var k = BacaKonteks();
            if (k == null) return;
            if (tanggal != null && Convert.ToString(k.ContainsKey("tanggal") ? k["tanggal"] : "") != tanggal) return;
            k["tersimpanVersi"] = null;
            k["waktuSimpan"] = null;
            SimpanKonteks(k);
        }

        public static Dictionary<string, object> BacaKonteks()
        {
            try
            {
                string f = Path.Combine(AppPaths.FolderDasar, "konteks.json");
                if (!File.Exists(f)) return null;
                return _json.Deserialize<Dictionary<string, object>>(File.ReadAllText(f));
            }
            catch { return null; }
        }

        private static void SimpanKonteks(Dictionary<string, object> k)
        {
            try
            {
                Directory.CreateDirectory(AppPaths.FolderDasar);
                File.WriteAllText(Path.Combine(AppPaths.FolderDasar, "konteks.json"), _json.Serialize(k));
            }
            catch (Exception ex) { CatatanLog.Tulis("Simpan konteks gagal: " + ex.Message); }
        }

        // =================================================================
        // SUSUNAN GRUP
        // =================================================================
        public class Grup
        {
            public int Urut;
            public string Nama, KodeKC, Dasar;
            public int? TopN;                      // Top-N debitur CKPN Individu grup ini (Tahap 4c)
            public List<string> KC = new List<string>();
        }

        /// <summary>
        /// Susunan grup untuk tahun tertentu. Bila tahun itu belum ditetapkan,
        /// dipakai susunan tahun terakhir sebelumnya (sumberTahun menunjukkan asalnya).
        /// Bila belum pernah ditetapkan sama sekali → null.
        /// </summary>
        public static List<Grup> Susunan(SQLiteConnection con, int tahun, out int sumberTahun)
        {
            sumberTahun = 0;
            object t = Database.Scalar(con, "SELECT MAX(tahun) FROM susunan_grup WHERE tahun<=@p0", tahun);
            if (t == null || t is DBNull) return null;
            sumberTahun = Convert.ToInt32(t);

            var hasil = new List<Grup>();
            using (var cmd = Database.Cmd(con, "SELECT urut, nama, kode_kc, dasar, top_n FROM susunan_grup WHERE tahun=@p0 ORDER BY urut", sumberTahun))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                {
                    var g = new Grup
                    {
                        Urut = rd.GetInt32(0), Nama = rd.GetString(1), KodeKC = rd.GetString(2), Dasar = Str(rd[3]),
                        TopN = rd[4] is DBNull ? (int?)null : Convert.ToInt32(rd[4])
                    };
                    g.KC.AddRange(g.KodeKC.Split(','));
                    hasil.Add(g);
                }
            return hasil;
        }

        // ---- Pengaturan konsolidasi per tahun (Tahap 3c) ----
        public class PengaturanTahun
        {
            public int Tahun;
            public string Metode;                  // "nf" | "mig"
            public string KebijakanSaldo = "ckpn"; // "ckpn" | "maks"
            public int? TopN;                      // Top-N debitur CKPN Individu (Tahap 4b)
            public string Pengguna, Waktu;
        }

        /// <summary>
        /// Metode & kebijakan saldo yang berlaku untuk tahun tertentu. Seperti susunan grup,
        /// bila tahun itu belum ditetapkan dipakai pengaturan tahun terakhir sebelumnya.
        /// </summary>
        public static PengaturanTahun Pengaturan(SQLiteConnection con, int tahun)
        {
            object t = Database.Scalar(con, "SELECT MAX(tahun) FROM susunan_tahun WHERE tahun<=@p0 AND metode IS NOT NULL", tahun);
            if (t == null || t is DBNull) return null;
            using (var cmd = Database.Cmd(con,
                "SELECT tahun, metode, kebijakan_saldo, pengguna, waktu, top_n FROM susunan_tahun WHERE tahun=@p0", Convert.ToInt32(t)))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) return null;
                return new PengaturanTahun
                {
                    Tahun = rd.GetInt32(0), Metode = Str(rd[1]),
                    KebijakanSaldo = Str(rd[2]) == "maks" ? "maks" : "ckpn",
                    Pengguna = Str(rd[3]), Waktu = Str(rd[4]),
                    TopN = rd[5] is DBNull ? (int?)null : Convert.ToInt32(rd[5])
                };
            }
        }

        /// <summary>
        /// Top-N yang ditetapkan untuk grup (kombinasi kode KC) pada susunan tahun tersebut.
        /// Null bila database/susunan belum ada, kombinasi KC tidak ada di susunan, atau Top-N
        /// grup belum diisi. Dipakai pipeline (tulis Master!C10 sebelum hitung grup) dan Simpan grup.
        /// </summary>
        public static int? TopNGrup(int tahun, string kodeKC, out string namaGrup)
        {
            namaGrup = null;
            try
            {
                if (!Database.Ada) return null;
                using (var con = Database.Buka(false))
                {
                    int sumber;
                    var susunan = Susunan(con, tahun, out sumber);
                    var g = susunan == null ? null : susunan.Find(x => string.Equals(x.KodeKC, kodeKC, StringComparison.OrdinalIgnoreCase));
                    if (g == null) return null;
                    namaGrup = g.Nama;
                    return g.TopN;
                }
            }
            catch (Exception ex) { CatatanLog.Tulis("Baca Top-N grup: " + ex.Message); return null; }
        }

        public static string NamaMetode(string m)
        {
            return m == "nf" ? "Net Flow" : m == "mig" ? "Migration" : "—";
        }

        /// <summary>Usulan awal bila belum ada susunan: satu grup per kode segmen.</summary>
        public static List<Grup> SusunanBawaan()
        {
            var hasil = new List<Grup>();
            for (int i = 0; i < SemuaKC.Length; i++)
                hasil.Add(new Grup { Urut = i + 1, Nama = SemuaKC[i], KodeKC = SemuaKC[i], KC = new List<string> { SemuaKC[i] } });
            return hasil;
        }

        /// <summary>
        /// peta: kode KC → nama grup. Validasi: 6 kode tercakup tepat sekali.
        /// metode: "nf" | "mig" — satu metode konsolidasi untuk semua grup sepanjang tahun.
        /// kebijakanSaldo: "ckpn" | "maks" — dasar saldo CKPN yang dibukukan (lihat HitungJurnal).
        /// </summary>
        public static void SimpanSusunan(int tahun, Dictionary<string, string> peta, string dasar,
                                         string metode, string kebijakanSaldo, Dictionary<string, int> topNGrup)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(dasar)) throw new InvalidOperationException("Dasar kebijakan/SOP wajib diisi.");
            if (metode != "nf" && metode != "mig")
                throw new InvalidOperationException("Pilih metode konsolidasi (Net Flow atau Migration) untuk tahun " + tahun + ".");
            if (kebijakanSaldo != "maks") kebijakanSaldo = "ckpn";

            var grup = new List<Grup>();
            var index = new Dictionary<string, Grup>(StringComparer.OrdinalIgnoreCase);
            foreach (var kc in SemuaKC)
            {
                string nama;
                if (!peta.TryGetValue(kc, out nama) || string.IsNullOrWhiteSpace(nama))
                    throw new InvalidOperationException("Kode " + kc + " belum masuk grup mana pun.");
                nama = nama.Trim();
                Grup g;
                if (!index.TryGetValue(nama, out g))
                {
                    g = new Grup { Nama = nama, Urut = grup.Count + 1 };
                    index[nama] = g;
                    grup.Add(g);
                }
                g.KC.Add(kc);   // urutan kanonik KC0600..KC1100 terjaga
            }

            // Top-N per grup (kunci = nama grup), wajib 1–1000 sesuai SOP
            foreach (var g in grup)
            {
                int n;
                if (topNGrup == null || !topNGrup.TryGetValue(g.Nama, out n) || n < 1 || n > 1000)
                    throw new InvalidOperationException("Top-N debitur CKPN Individu grup '" + g.Nama + "' harus diisi (1–1000) sesuai SOP.");
                g.TopN = n;
                g.KodeKC = string.Join(",", g.KC.ToArray());
            }

            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                // Metode tidak boleh berganti di tengah tahun setelah ada periode yang dikunci
                // dengan metode lain. Buka kunci periode tersebut dulu bila memang harus diganti.
                var bentrok = new List<string>();
                using (var cmd = Database.Cmd(con,
                    "SELECT tanggal, metode, IFNULL(kebijakan_saldo,'ckpn') FROM periode " +
                    "WHERE status='Final' AND substr(tanggal,1,4)=@p0 AND metode IS NOT NULL ORDER BY tanggal", tahun.ToString()))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        if (rd.GetString(1) != metode)
                            bentrok.Add(rd.GetString(0) + " (" + NamaMetode(rd.GetString(1)) + ")");
                // Top-N per grup juga dikunci: tidak boleh berubah bila sudah ada periode Final di tahun itu
                bool adaFinal = Convert.ToInt32(Database.Scalar(con,
                    "SELECT COUNT(*) FROM periode WHERE status='Final' AND substr(tanggal,1,4)=@p0", tahun.ToString())) > 0;
                if (adaFinal)
                {
                    var berubah = new List<string>();
                    foreach (var g in grup)
                    {
                        object lama = Database.Scalar(con, "SELECT top_n FROM susunan_grup WHERE tahun=@p0 AND kode_kc=@p1", tahun, g.KodeKC);
                        if (lama != null && !(lama is DBNull) && Convert.ToInt32(lama) != g.TopN)
                            berubah.Add(g.Nama + " (" + Convert.ToInt32(lama) + " → " + g.TopN + ")");
                    }
                    if (berubah.Count > 0)
                        throw new InvalidOperationException("Top-N grup " + string.Join(", ", berubah.ToArray()) +
                            " tidak dapat diubah karena sudah ada periode Final di tahun " + tahun +
                            ". Buka kunci periode tersebut terlebih dahulu bila perubahan memang disetujui.");
                }

                if (bentrok.Count > 0)
                    throw new InvalidOperationException(
                        "Metode tahun " + tahun + " tidak dapat diubah karena sudah dipakai di periode Final: " +
                        string.Join(", ", bentrok.ToArray()) + ". Buka kunci periode tersebut terlebih dahulu bila perubahan memang disetujui.");

                Database.Exec(con,
                    "INSERT OR REPLACE INTO susunan_tahun(tahun,metode,kebijakan_saldo,dasar,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4,@p5)",
                    tahun, metode, kebijakanSaldo, dasar.Trim(), Environment.UserName, Database.Sekarang());

                Database.Exec(con, "DELETE FROM susunan_grup WHERE tahun=@p0", tahun);
                foreach (var g in grup)
                {
                    Database.Exec(con,
                        "INSERT INTO susunan_grup(tahun,urut,nama,kode_kc,dasar,pengguna,waktu,top_n) VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7)",
                        tahun, g.Urut, g.Nama, g.KodeKC, dasar.Trim(), Environment.UserName, Database.Sekarang(), g.TopN);
                }
                var ringkas = new List<string>();
                foreach (var g in grup) ringkas.Add(g.Nama + "=" + g.KodeKC + " (Top-N " + g.TopN + ")");
                Database.CatatAktivitas(con, tahun.ToString(), "susunan-grup",
                    string.Join("; ", ringkas.ToArray()) + " | metode: " + NamaMetode(metode) +
                    " | saldo: " + (kebijakanSaldo == "maks" ? "maks(CKPN, PPKA)" : "CKPN") +
                    " | dasar: " + dasar.Trim());
                tx.Commit();
            }
        }

        // =================================================================
        // STATUS PERIODE — dasar tampilan alur di panel
        // =================================================================
        /// <param name="tanggal">"yyyy-MM-dd"</param>
        /// <param name="kodeKCMaster">kode KC yang sedang dicentang di Master (workbook saat ini)</param>
        /// <param name="tanggalMaster">periode Master!C4</param>
        public static Dictionary<string, object> Status(string tanggal, string kodeKCMaster, string tanggalMaster)
        {
            int tahun = int.Parse(tanggal.Substring(0, 4));
            var hasil = new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "tanggalMaster", tanggalMaster },
                { "periodeMaster", tanggal == tanggalMaster },
                { "kodeKCMaster", kodeKCMaster }, { "tahun", tahun },
                { "adaDatabase", Database.Ada }, { "status", "Terbuka" }
            };
            string info;
            hasil["bolehMenulis"] = Database.BolehMenulis(out info);
            hasil["infoPengirim"] = info;

            var konteks = BacaKonteks();
            hasil["konteks"] = konteks;

            List<Grup> susunan = null;
            PengaturanTahun peng = null;
            int sumberTahun = 0;
            var runs = new List<Dictionary<string, object>>();
            var daftarPeriode = new List<object>();

            if (Database.Ada)
            {
                using (var con = Database.Buka(false))
                {
                    susunan = Susunan(con, tahun, out sumberTahun);
                    peng = Pengaturan(con, tahun);

                    using (var cmd = Database.Cmd(con,
                        "SELECT status, dikunci_oleh, dikunci_waktu, catatan_kunci FROM periode WHERE tanggal=@p0", tanggal))
                    using (var rd = cmd.ExecuteReader())
                        if (rd.Read())
                        {
                            hasil["status"] = rd.GetString(0);
                            hasil["dikunciOleh"] = Str(rd[1]);
                            hasil["dikunciWaktu"] = Str(rd[2]);
                            hasil["catatanKunci"] = Str(rd[3]);
                        }

                    using (var cmd = Database.Cmd(con,
                        "SELECT p.tanggal, p.status, COUNT(g.id) FROM periode p " +
                        "LEFT JOIN run_grup g ON g.periode_id=p.id AND g.aktif=1 AND g.dihapus=0 " +
                        "GROUP BY p.id ORDER BY p.tanggal DESC"))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read())
                            daftarPeriode.Add(new Dictionary<string, object>
                            {
                                { "tanggal", rd.GetString(0) }, { "status", rd.GetString(1) }, { "jumlahGrup", Convert.ToInt32(rd[2]) }
                            });

                    runs = DaftarRun(con, tanggal);
                }
            }
            hasil["daftarPeriode"] = daftarPeriode;

            bool susunanDitetapkan = susunan != null;
            if (susunan == null) susunan = SusunanBawaan();
            hasil["susunanDitetapkan"] = susunanDitetapkan;
            hasil["susunanTahun"] = sumberTahun;
            hasil["susunanDasar"] = susunan.Count > 0 ? susunan[0].Dasar : "";
            hasil["metode"] = peng == null ? null : peng.Metode;
            hasil["metodeTahun"] = peng == null ? 0 : peng.Tahun;
            hasil["kebijakanSaldo"] = peng == null ? "ckpn" : peng.KebijakanSaldo;

            // ---- status per grup ----
            var grupList = new List<object>();
            var kodeGrup = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int tersimpan = 0;
            foreach (var g in susunan)
            {
                kodeGrup.Add(g.KodeKC);
                var versi = runs.FindAll(r => (string)r["kodeKC"] == g.KodeKC);
                var aktif = versi.Find(r => (bool)r["aktif"] && !(bool)r["dihapus"]);
                if (aktif != null) tersimpan++;

                bool diWorkbook = string.Equals(g.KodeKC, kodeKCMaster, StringComparison.OrdinalIgnoreCase) && tanggal == tanggalMaster;
                bool dihitungDiPC = konteks != null && diWorkbook &&
                                    Convert.ToString(Ambil(konteks, "tanggal")) == tanggal &&
                                    Convert.ToString(Ambil(konteks, "kodeKC")) == g.KodeKC &&
                                    Ambil(konteks, "waktuHitung") != null;
                bool belumDisimpanSetelahHitung = dihitungDiPC && Ambil(konteks, "tersimpanVersi") == null;

                string status = aktif != null
                    ? (belumDisimpanSetelahHitung ? "dihitung-ulang" : "tersimpan")
                    : (dihitungDiPC ? "dihitung" : "belum");

                grupList.Add(new Dictionary<string, object>
                {
                    { "urut", g.Urut }, { "nama", g.Nama }, { "kodeKC", g.KodeKC }, { "kc", g.KC }, { "topN", g.TopN },
                    { "status", status }, { "aktif", aktif }, { "versi", versi },
                    { "diWorkbook", diWorkbook }, { "dihitungDiPC", dihitungDiPC },
                    { "waktuHitung", dihitungDiPC ? Ambil(konteks, "waktuHitung") : null }
                });
            }
            hasil["grup"] = grupList;
            hasil["jumlahGrup"] = susunan.Count;
            hasil["grupTersimpan"] = tersimpan;

            // Kiriman yang KC-nya tidak cocok dengan susunan (mis. dihitung sebelum susunan ditetapkan)
            var luar = runs.FindAll(r => !kodeGrup.Contains((string)r["kodeKC"]) && !(bool)r["dihapus"]);
            hasil["runLuarSusunan"] = luar;

            hasil["siapKonsolidasi"] = susunanDitetapkan && tersimpan == susunan.Count && susunan.Count > 0;
            hasil["awalTransisi"] = AwalTransisi;
            return hasil;
        }

        internal static List<Dictionary<string, object>> DaftarRun(SQLiteConnection con, string tanggal)
        {
            var runs = new List<Dictionary<string, object>>();
            using (var cmd = Database.Cmd(con,
                "SELECT g.id, g.kode_kc, g.versi, g.aktif, g.dihapus, g.pengguna, g.waktu, g.catatan, g.snapshot, " +
                "g.alasan_hapus, g.pengguna_hapus, g.waktu_hapus " +
                "FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE p.tanggal=@p0 ORDER BY g.kode_kc, g.versi DESC", tanggal))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    runs.Add(new Dictionary<string, object>
                    {
                        { "id", rd.GetInt64(0) }, { "kodeKC", rd.GetString(1) }, { "versi", rd.GetInt32(2) },
                        { "aktif", rd.GetInt32(3) == 1 }, { "dihapus", rd.GetInt32(4) == 1 },
                        { "pengguna", Str(rd[5]) }, { "waktu", Str(rd[6]) }, { "catatan", Str(rd[7]) },
                        { "adaSnapshot", !string.IsNullOrEmpty(Str(rd[8])) && File.Exists(Str(rd[8])) },
                        { "alasanHapus", Str(rd[9]) }, { "penggunaHapus", Str(rd[10]) }, { "waktuHapus", Str(rd[11]) }
                    });

            foreach (var run in runs)
            {
                var nilai = new Dictionary<string, object>();
                using (var cmd = Database.Cmd(con, "SELECT kunci, nilai FROM ringkasan WHERE run_id=@p0", run["id"]))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) nilai[rd.GetString(0)] = Database.Dbl(rd[1]);
                run["ringkasan"] = nilai;
                run["ppkaGrup"] = PpkaGrup(nilai, ((string)run["kodeKC"]).Split(','));
            }
            return runs;
        }

        /// <summary>PPKA OJK untuk KC dalam grup (Summary!H3:H8). PPKA ABA (H2) dihitung terpisah.</summary>
        private static double PpkaGrup(Dictionary<string, object> nilai, string[] kc)
        {
            double t = 0;
            foreach (var k in kc)
            {
                object v;
                if (nilai.TryGetValue("ppka_" + k.Trim(), out v) && v != null) t += Convert.ToDouble(v);
            }
            return t;
        }

        // =================================================================
        // HAPUS VERSI (soft delete)
        // =================================================================
        public static void HapusVersi(long runId, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan penghapusan wajib diisi.");

            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                long periodeId; string kode, tanggal, status; int versi; bool aktif;
                using (var cmd = Database.Cmd(con,
                    "SELECT g.periode_id, g.kode_kc, g.versi, g.aktif, p.tanggal, p.status FROM run_grup g " +
                    "JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0 AND g.dihapus=0", runId))
                using (var rd = cmd.ExecuteReader())
                {
                    if (!rd.Read()) throw new InvalidOperationException("Versi tidak ditemukan atau sudah dihapus.");
                    periodeId = rd.GetInt64(0); kode = rd.GetString(1); versi = rd.GetInt32(2);
                    aktif = rd.GetInt32(3) == 1; tanggal = rd.GetString(4); status = rd.GetString(5);
                }
                if (status == "Final") throw new InvalidOperationException("Periode " + tanggal + " sudah dikunci. Buka kunci terlebih dahulu.");

                Database.Exec(con,
                    "UPDATE run_grup SET dihapus=1, aktif=0, alasan_hapus=@p0, pengguna_hapus=@p1, waktu_hapus=@p2 WHERE id=@p3",
                    alasan.Trim(), Environment.UserName, Database.Sekarang(), runId);

                // Versi aktif dihapus → versi terakhir yang tersisa menjadi aktif kembali
                if (aktif)
                    Database.Exec(con,
                        "UPDATE run_grup SET aktif=1 WHERE id=(SELECT id FROM run_grup WHERE periode_id=@p0 AND kode_kc=@p1 " +
                        "AND dihapus=0 ORDER BY versi DESC LIMIT 1)", periodeId, kode);

                Database.CatatAktivitas(con, tanggal, "hapus-versi", kode + " v" + versi + " | " + alasan.Trim());
                tx.Commit();
            }
        }

        public static void PastikanTerbuka(SQLiteConnection con, string tanggal)
        {
            object s = Database.Scalar(con, "SELECT status FROM periode WHERE tanggal=@p0", tanggal);
            if (s != null && Convert.ToString(s) == "Final")
                throw new InvalidOperationException("Periode " + tanggal + " sudah dikunci (Final). Buka kunci terlebih dahulu di tab Periode.");
        }

        // =================================================================
        // KONSOLIDASI vs PPKA  (Tahap 3c: satu metode + usulan jurnal)
        // =================================================================

        /// <summary>
        /// Posisi pertama yang dianggap transisi PSAK 414: CKPN posisi Desember 2026
        /// menjadi saldo pembukaan 1 Januari 2027. Periode sebelumnya = simulasi (parallel run).
        /// </summary>
        public const string AwalTransisi = "2026-12-01";

        public static Dictionary<string, object> Konsolidasi(string tanggal)
        {
            if (!Database.Ada) throw new InvalidOperationException("Database belum ada.");
            int tahun = int.Parse(tanggal.Substring(0, 4));

            using (var con = Database.Buka(false))
            {
                int sumber;
                var susunan = Susunan(con, tahun, out sumber);
                if (susunan == null) throw new InvalidOperationException("Susunan grup tahun " + tahun + " belum ditetapkan.");
                var runs = DaftarRun(con, tanggal);

                // ---- status periode & metode ----
                long periodeId = 0;
                string status = "Terbuka", metodeFinal = null, kebijakanFinal = null;
                using (var cmd = Database.Cmd(con, "SELECT id, status, metode, kebijakan_saldo FROM periode WHERE tanggal=@p0", tanggal))
                using (var rd = cmd.ExecuteReader())
                    if (rd.Read())
                    {
                        periodeId = rd.GetInt64(0); status = rd.GetString(1);
                        metodeFinal = Str(rd[2]); kebijakanFinal = Str(rd[3]);
                    }
                bool final = status == "Final";
                var peng = Pengaturan(con, tahun);

                string metode = final && metodeFinal != "" ? metodeFinal : (peng == null ? null : peng.Metode);
                string kebijakan = final && kebijakanFinal != "" ? kebijakanFinal : (peng == null ? "ckpn" : peng.KebijakanSaldo);

                // Periode yang dikunci di Tahap 3b menyimpan metode per grup (bisa campuran).
                var keputusan = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (final)
                    using (var cmd = Database.Cmd(con, "SELECT kode_kc, metode FROM keputusan_metode WHERE periode_id=@p0", periodeId))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) keputusan[rd.GetString(0)] = rd.GetString(1);

                // ---- per grup ----
                var baris = new List<object>();
                var belum = new List<string>();
                Dictionary<string, object> ringkasTerbaru = null;
                string waktuTerbaru = "";
                double totNf = 0, totMig = 0, totPpka = 0, totDipakai = 0;
                bool adaDipakai = true;

                foreach (var g in susunan)
                {
                    var aktif = runs.Find(r => (string)r["kodeKC"] == g.KodeKC && (bool)r["aktif"] && !(bool)r["dihapus"]);
                    if (aktif == null) { belum.Add(g.Nama); continue; }
                    var rg = (Dictionary<string, object>)aktif["ringkasan"];
                    double nf = Nilai(rg, "nf_total"), mig = Nilai(rg, "mig_total");
                    double ppka = Convert.ToDouble(aktif["ppkaGrup"]);

                    string dipakai;
                    if (!(final && keputusan.TryGetValue(g.KodeKC, out dipakai))) dipakai = metode;
                    double? ckpn = dipakai == "nf" ? nf : dipakai == "mig" ? mig : (double?)null;
                    if (ckpn == null) adaDipakai = false; else totDipakai += ckpn.Value;
                    totNf += nf; totMig += mig; totPpka += ppka;

                    baris.Add(new Dictionary<string, object>
                    {
                        { "nama", g.Nama }, { "kodeKC", g.KodeKC }, { "runId", aktif["id"] }, { "versi", aktif["versi"] },
                        { "nf", nf }, { "mig", mig }, { "ppka", ppka }, { "dipakai", dipakai }, { "ckpn", ckpn }
                    });
                    string w = Convert.ToString(aktif["waktu"]);
                    if (string.CompareOrdinal(w, waktuTerbaru) > 0) { waktuTerbaru = w; ringkasTerbaru = rg; }
                }

                // CKPN & PPKA antar bank (KC0500) sama untuk semua grup — ambil dari kiriman terbaru
                double abaCkpn = ringkasTerbaru == null ? 0 : Nilai(ringkasTerbaru, "aba_ckpn");
                double abaPpka = ringkasTerbaru == null ? 0 : Nilai(ringkasTerbaru, "ppka_KC0500");
                bool abaAda = ringkasTerbaru != null && ringkasTerbaru.ContainsKey("aba_ckpn");
                var aba = new Dictionary<string, object> { { "ckpn", abaCkpn }, { "ppka", abaPpka }, { "tersedia", abaAda } };

                // Pembanding total bank (informasi; metode tetap dari susunan tahunan)
                var pembanding = new Dictionary<string, object>
                {
                    { "nf", totNf + abaCkpn }, { "mig", totMig + abaCkpn }, { "ppka", totPpka + abaPpka },
                    { "terendah", totNf <= totMig ? "nf" : "mig" }
                };

                var hasil = new Dictionary<string, object>
                {
                    { "tanggal", tanggal }, { "status", status }, { "grup", baris }, { "belum", belum }, { "aba", aba },
                    { "lengkap", belum.Count == 0 }, { "metode", metode }, { "namaMetode", NamaMetode(metode) },
                    { "metodeTahun", final && metodeFinal != "" ? tahun : (peng == null ? 0 : peng.Tahun) },
                    { "metodeCampuran", final && metodeFinal == "" && keputusan.Count > 0 },
                    { "kebijakanSaldo", kebijakan }, { "pembanding", pembanding },
                    { "total", new Dictionary<string, object>
                        {
                            { "ckpn", adaDipakai ? totDipakai + abaCkpn : (double?)null },
                            { "ppka", totPpka + abaPpka }
                        }
                    }
                };

                // ---- usulan jurnal: satu jurnal atas total CKPN (pembiayaan + ABA) vs total PPKA ----
                if (final && periodeId > 0 && AdaJurnalTersimpan(con, periodeId))
                {
                    hasil["jurnal"] = JurnalTersimpan(con, periodeId);
                    hasil["jurnalTersimpan"] = true;
                }
                else if (!final && adaDipakai && belum.Count == 0)
                {
                    hasil["jurnal"] = new List<object> { HitungJurnal(con, tanggal, totDipakai + abaCkpn, totPpka + abaPpka) };
                    hasil["jurnalTersimpan"] = false;
                }
                return hasil;
            }
        }

        /// <summary>
        /// transisi : jurnal 1x — posisi pertama ≥ Desember 2026 (saldo awal 2027) yang belum punya
        ///            jurnal transisi Final sebelumnya. Periode sebelum Desember 2026 juga dihitung
        ///            sebagai transisi dengan tanda "simulasi" (parallel run, tidak dibukukan).
        /// reguler  : periode sesudah jurnal transisi.
        /// </summary>
        private static string JenisJurnal(SQLiteConnection con, string tanggal, out bool simulasi)
        {
            simulasi = string.CompareOrdinal(tanggal, AwalTransisi) < 0;
            if (simulasi) return "transisi";
            object n = Database.Scalar(con,
                "SELECT COUNT(*) FROM jurnal_ckpn j JOIN periode p ON p.id=j.periode_id " +
                "WHERE j.jenis='transisi' AND p.status='Final' AND p.tanggal>=@p0 AND p.tanggal<@p1", AwalTransisi, tanggal);
            return Convert.ToInt32(n) > 0 ? "reguler" : "transisi";
        }

        // Nama akun sesuai kebijakan jurnal bank (Tahap 3d)
        private const string AkunLabaDitahan = "Laba ditahan";
        private const string AkunCkpn        = "CKPN";
        private const string AkunBiayaCkpn   = "Biaya CKPN";
        private const string AkunCadangan    = "Cadangan CKPN";
        private const string AkunBiayaAtauPendapatan = "Biaya CKPN / Pendapatan";

        /// <summary>
        /// Selisih = CKPN (metode tahunan, termasuk ABA) − PPKA.
        ///   CKPN > PPKA : transisi Db. Laba ditahan – Kr. CKPN     | reguler Db. Biaya CKPN – Kr. CKPN
        ///   PPKA > CKPN : transisi Db. CKPN – Kr. Laba ditahan     | reguler Db. Cadangan CKPN – Kr. Biaya CKPN / Pendapatan
        /// </summary>
        private static Dictionary<string, object> HitungJurnal(SQLiteConnection con, string tanggal, double ckpn, double ppka)
        {
            bool simulasi;
            string jenis = JenisJurnal(con, tanggal, out simulasi);
            double selisih = Math.Round(ckpn - ppka, 0);

            string debit = null, kredit = null, ket;
            if (Math.Abs(selisih) < 1)
            {
                selisih = 0;
                ket = "CKPN sama dengan PPKA — tidak perlu jurnal.";
            }
            else if (jenis == "transisi")
            {
                if (selisih > 0) { debit = AkunLabaDitahan; kredit = AkunCkpn; ket = "PPKA lebih kecil dari CKPN — bentuk tambahan CKPN."; }
                else { debit = AkunCkpn; kredit = AkunLabaDitahan; ket = "PPKA lebih besar dari CKPN — kelebihan dikembalikan ke laba ditahan."; }
            }
            else
            {
                if (selisih > 0) { debit = AkunBiayaCkpn; kredit = AkunCkpn; ket = "CKPN kurang — PPKA lebih kecil dari CKPN."; }
                else { debit = AkunCadangan; kredit = AkunBiayaAtauPendapatan; ket = "CKPN lebih — PPKA lebih besar dari CKPN."; }
            }

            var j = new Dictionary<string, object>
            {
                { "komponen", "total" }, { "label", "CKPN (pembiayaan + ABA)" }, { "jenis", jenis }, { "simulasi", simulasi },
                { "ckpn", ckpn }, { "ppka", ppka }, { "selisihPpka", ckpn - ppka },
                { "nominal", Math.Abs(selisih) }, { "arah", Math.Sign(selisih) },
                { "debit", debit }, { "kredit", kredit }, { "keterangan", ket }
            };

            // Nominal = selisih penuh bulan berjalan. Setelah jurnal dibukukan ke CBS, PPKA = CKPN;
            // bulan berikutnya PPKA OJK dihitung ulang dan selisihnya dicek lagi dari awal,
            // sehingga tidak ada perbandingan dengan jurnal bulan sebelumnya.
            return j;
        }

        private static bool AdaJurnalTersimpan(SQLiteConnection con, long periodeId)
        {
            return Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM jurnal_ckpn WHERE periode_id=@p0", periodeId)) > 0;
        }

        private static List<object> JurnalTersimpan(SQLiteConnection con, long periodeId)
        {
            var hasil = new List<object>();
            using (var cmd = Database.Cmd(con,
                "SELECT j.komponen, j.jenis, j.ckpn, j.ppka, j.nominal, j.akun_debit, j.akun_kredit, j.keterangan, p.tanggal " +
                "FROM jurnal_ckpn j JOIN periode p ON p.id=j.periode_id WHERE j.periode_id=@p0 ORDER BY j.komponen DESC", periodeId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                {
                    double ckpn = Database.Dbl(rd[2]) ?? 0, ppka = Database.Dbl(rd[3]) ?? 0;
                    string komponen = rd.GetString(0);
                    hasil.Add(new Dictionary<string, object>
                    {
                        { "komponen", komponen },
                        { "label", komponen == "aba" ? "Penempatan pada bank lain (ABA)" :
                                   komponen == "pembiayaan" ? "Pembiayaan" : "CKPN (pembiayaan + ABA)" },
                        { "jenis", rd.GetString(1) },
                        { "simulasi", string.CompareOrdinal(rd.GetString(8), AwalTransisi) < 0 },
                        { "ckpn", ckpn }, { "ppka", ppka }, { "selisihPpka", ckpn - ppka },
                        { "nominal", Database.Dbl(rd[4]) ?? 0 }, { "arah", Math.Sign(Math.Round(ckpn - ppka, 0)) },
                        { "debit", Str(rd[5]) == "" ? null : Str(rd[5]) }, { "kredit", Str(rd[6]) == "" ? null : Str(rd[6]) },
                        { "keterangan", Str(rd[7]) }
                    });
                }
            return hasil;
        }

        public static void Tetapkan(string tanggal, string catatan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            var k = Konsolidasi(tanggal);
            if (!(bool)k["lengkap"])
                throw new InvalidOperationException("Belum semua grup tersimpan: " + string.Join(", ", ((List<string>)k["belum"]).ToArray()));
            string metode = k["metode"] as string;
            if (metode != "nf" && metode != "mig")
                throw new InvalidOperationException("Metode konsolidasi tahun " + tanggal.Substring(0, 4) +
                    " belum ditetapkan. Tetapkan di Atur susunan grup terlebih dahulu.");
            const string kebijakan = "ckpn";   // Tahap 3d: jurnal selalu CKPN vs PPKA
            var total = (Dictionary<string, object>)k["total"];
            var jurnal = k.ContainsKey("jurnal") ? (List<object>)k["jurnal"] : new List<object>();

            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                PastikanTerbuka(con, tanggal);
                long periodeId = Convert.ToInt64(Database.Scalar(con, "SELECT id FROM periode WHERE tanggal=@p0", tanggal));
                string sekarang = Database.Sekarang();

                Database.Exec(con, "DELETE FROM keputusan_metode WHERE periode_id=@p0", periodeId);
                foreach (Dictionary<string, object> b in (List<object>)k["grup"])
                    Database.Exec(con,
                        "INSERT INTO keputusan_metode(periode_id,kode_kc,nama_grup,run_id,metode,ckpn_nf,ckpn_mig,ppka,pengguna,waktu) " +
                        "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9)",
                        periodeId, b["kodeKC"], b["nama"], b["runId"], metode, b["nf"], b["mig"], b["ppka"],
                        Environment.UserName, sekarang);

                Database.Exec(con, "DELETE FROM jurnal_ckpn WHERE periode_id=@p0", periodeId);
                var ringkasJurnal = new List<string>();
                foreach (Dictionary<string, object> j in jurnal)
                {
                    Database.Exec(con,
                        "INSERT INTO jurnal_ckpn(periode_id,komponen,jenis,ckpn,ppka,saldo_target,saldo_awal,nominal," +
                        "akun_debit,akun_kredit,keterangan,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12)",
                        periodeId, j["komponen"], j["jenis"], j["ckpn"], j["ppka"], j["ckpn"], j["ppka"], j["nominal"],
                        j["debit"], j["kredit"], j["keterangan"], Environment.UserName, sekarang);
                    ringkasJurnal.Add(j["komponen"] + " " + j["jenis"] + " " +
                        (j["debit"] == null ? "tanpa jurnal" : "Db " + j["debit"] + " / Kr " + j["kredit"] + " " +
                         Convert.ToDouble(j["nominal"]).ToString("#,##0", CultureInfo.InvariantCulture)));
                }

                Database.Exec(con,
                    "UPDATE periode SET status='Final', dikunci_oleh=@p0, dikunci_waktu=@p1, catatan_kunci=@p2, " +
                    "metode=@p3, kebijakan_saldo=@p4, ckpn_total=@p5, ppka_total=@p6 WHERE id=@p7",
                    Environment.UserName, sekarang, catatan ?? "", metode, kebijakan, total["ckpn"], total["ppka"], periodeId);
                Database.CatatAktivitas(con, tanggal, "kunci-periode",
                    "metode " + NamaMetode(metode) + " | " + string.Join("; ", ringkasJurnal.ToArray()) + " | " + (catatan ?? ""));
                tx.Commit();
            }
        }

        public static void BukaKunci(string tanggal, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan membuka kunci wajib diisi.");
            Database.Cadangkan();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                Database.Exec(con,
                    "UPDATE periode SET status='Terbuka', dikunci_oleh=NULL, dikunci_waktu=NULL, catatan_kunci=NULL WHERE tanggal=@p0",
                    tanggal);
                // Keputusan metode lama tetap disimpan sampai konsolidasi ditetapkan ulang
                Database.CatatAktivitas(con, tanggal, "buka-kunci", alasan.Trim());
                tx.Commit();
            }
        }

        // =================================================================
        internal static double Nilai(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) && v != null ? Convert.ToDouble(v) : 0;
        }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
