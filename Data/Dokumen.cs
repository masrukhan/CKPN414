using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Dokumentasi perhitungan CKPN (Tahap 6) — tab Dokumen di panel.
    ///
    /// Tiga jenis dokumen per periode (tabel memo, skema v9):
    ///   ringkasan  (kunci '')        — catatan &amp; penandatangan Laporan Ringkasan CKPN bulanan
    ///   individu   (kunci = CIF)     — memo penilaian penurunan nilai individual per debitur:
    ///                                  profil, bukti objektif, penilaian agunan, estimasi biaya penjualan
    ///   lgdcs      (kunci = no. rek) — memo penetapan nilai realisasi agunan (LGD Collateral Shortfall)
    ///
    /// Isi memo (isian petugas) disimpan sebagai JSON; angka hasil perhitungan dibaca dari versi grup
    /// aktif periode itu (hasil_individu / hasil_lgdcs), sehingga dokumen selalu sesuai angka tersimpan.
    /// Status Draf → Final (ditandatangani). Memo Final hanya bisa diubah setelah dibuka kembali (beralasan).
    /// Nilai agunan &amp; biaya di memo menjadi acuan perhitungan berikutnya lewat penyesuaian tersimpan
    /// (diterapkan dari panel dengan tombol "Terapkan ke penyesuaian").
    /// </summary>
    internal static class Dokumen
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private const int BatasData = 400 * 1024;
        private static readonly string[] Jenis = { "ringkasan", "individu", "lgdcs" };

        // =================================================================
        // Data untuk tab Dokumen (satu periode)
        // =================================================================
        public static Dictionary<string, object> Data(string tanggal)
        {
            string info;
            var h = new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "bolehMenulis", Database.BolehMenulis(out info) }, { "infoPengirim", info },
                { "individu", new List<object>() }, { "lgdcs", new List<object>() },
                { "memo", KosongMemo() }, { "memoLalu", KosongMemo() },
                { "profil", new Dictionary<string, object>() }, { "profilInfo", null }, { "kode", CKPNLibrary.Panel.ProfilTemplate.BacaKode() },
                { "penyesuaian", new Dictionary<string, object> {
                    { "individu", new Dictionary<string, object>() }, { "lgdcs", new Dictionary<string, object>() } } }
            };
            if (string.IsNullOrEmpty(tanggal) || !Database.Ada) return h;

            using (var con = Database.Buka(false))
            {
                bool adaBukti = KolomAda(con, "hasil_individu", "kualitas");
                bool adaMemo = TabelAda(con, "memo");
                h["skemaLengkap"] = adaBukti && adaMemo;

                // ---- CKPN Individu: semua kontrak dari versi aktif periode ini ----
                var ind = (List<object>)h["individu"];
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id, g.kode_kc, g.versi, h.urut, h.kc, h.cif, h.nama, h.no_kontrak, h.os, h.ada_pn, h.jaminan, h.biaya_jual, " +
                    "h.penurunan_nilai, h.disesuaikan" +
                    (adaBukti ? ", h.kualitas, h.hari_tunggakan, h.restruktur, h.dasar_pn " : " ") +
                    "FROM hasil_individu h JOIN run_grup g ON g.id=h.run_id JOIN periode p ON p.id=g.periode_id " +
                    "WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0 ORDER BY g.kode_kc, h.urut, h.no_kontrak", tanggal))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                    {
                        var b = new Dictionary<string, object>
                        {
                            { "runId", rd.GetInt64(0) }, { "kodeKC", rd.GetString(1) }, { "versi", rd.GetInt32(2) },
                            { "urut", rd[3] is DBNull ? 0 : Convert.ToInt32(rd[3]) }, { "kc", Str(rd[4]) }, { "cif", Str(rd[5]) },
                            { "nama", Str(rd[6]) }, { "kontrak", Str(rd[7]) }, { "os", Database.Dbl(rd[8]) }, { "adaPN", Str(rd[9]) },
                            { "jaminan", Database.Dbl(rd[10]) }, { "biaya", Database.Dbl(rd[11]) }, { "ckpn", Database.Dbl(rd[12]) },
                            { "disesuaikan", !(rd[13] is DBNull) && Convert.ToInt32(rd[13]) == 1 }
                        };
                        if (adaBukti)
                        {
                            b["kualitas"] = rd[14] is DBNull ? null : (object)Convert.ToInt32(rd[14]);
                            b["hari"] = Database.Dbl(rd[15]);
                            b["restruktur"] = rd[16] is DBNull ? null : (object)(Convert.ToInt32(rd[16]) == 1);
                            b["dasarPN"] = rd[17] is DBNull ? null : rd.GetString(17);
                        }
                        ind.Add(b);
                    }

                // ---- LGD CS: baris dari versi aktif periode ini ----
                var cs = (List<object>)h["lgdcs"];
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id, g.kode_kc, g.versi, c.no_rek, c.nama, c.pokok_awal, c.nilai_agunan, c.thn_diserahkan, c.thn_eksekusi, " +
                    "c.recovery, c.shortfall, c.sumber " +
                    "FROM hasil_lgdcs c JOIN run_grup g ON g.id=c.run_id JOIN periode p ON p.id=g.periode_id " +
                    "WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0 ORDER BY g.kode_kc, c.no_rek", tanggal))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        cs.Add(new Dictionary<string, object>
                        {
                            { "runId", rd.GetInt64(0) }, { "kodeKC", rd.GetString(1) }, { "versi", rd.GetInt32(2) },
                            { "rek", Str(rd[3]) }, { "nama", Str(rd[4]) }, { "pokok", Database.Dbl(rd[5]) }, { "agunan", Database.Dbl(rd[6]) },
                            { "thnSerah", Str(rd[7]) }, { "thnEks", Str(rd[8]) }, { "recovery", Database.Dbl(rd[9]) },
                            { "shortfall", Database.Dbl(rd[10]) }, { "sumber", Str(rd[11]) }
                        });

                // ---- Memo periode ini + memo terakhir sebelumnya (untuk "salin dari bulan lalu") ----
                if (adaMemo)
                {
                    var memo = (Dictionary<string, object>)h["memo"];
                    using (var cmd = Database.Cmd(con,
                        "SELECT id, jenis, kunci, data, status, dibuat_oleh, dibuat_waktu, diubah_oleh, diubah_waktu, final_oleh, final_waktu " +
                        "FROM memo WHERE tanggal=@p0", tanggal))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read())
                        {
                            var m = BacaMemo(rd);
                            string jenis = rd.GetString(1);
                            if (jenis == "ringkasan") memo["ringkasan"] = m;
                            else ((Dictionary<string, object>)memo[jenis])[rd.GetString(2)] = m;
                        }

                    var lalu = (Dictionary<string, object>)h["memoLalu"];
                    using (var cmd = Database.Cmd(con,
                        "SELECT m.jenis, m.kunci, m.tanggal, m.status, m.data FROM memo m " +
                        "WHERE m.jenis IN ('individu','lgdcs','ringkasan') AND m.tanggal < @p0 AND m.tanggal = " +
                        "(SELECT MAX(x.tanggal) FROM memo x WHERE x.jenis=m.jenis AND x.kunci=m.kunci AND x.tanggal < @p0)", tanggal))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read())
                        {
                            var m = new Dictionary<string, object>
                            {
                                { "tanggal", rd.GetString(2) }, { "status", rd.GetString(3) }, { "data", Urai(Str(rd[4])) }
                            };
                            string jenis = rd.GetString(0);
                            if (jenis == "ringkasan") lalu["ringkasan"] = m;
                            else ((Dictionary<string, object>)lalu[jenis])[rd.GetString(1)] = m;
                        }
                }

                // ---- Nama BPR (kop dokumen) dari snapshot overview terdekat ----
                try
                {
                    if (TabelAda(con, "overview_periode"))
                    {
                        object ov = Database.Scalar(con,
                            "SELECT data FROM overview_periode ORDER BY (tanggal=@p0) DESC, tanggal DESC LIMIT 1", tanggal);
                        var d = ov == null ? null : Urai(Convert.ToString(ov)) as Dictionary<string, object>;
                        object nama;
                        if (d != null && d.TryGetValue("namaBPR", out nama) && nama != null) h["namaBPR"] = Convert.ToString(nama);
                    }
                }
                catch { }

                // ---- Tahap 6b: profil rekening dari template APOLLO (tersimpan per periode) ----
                if (TabelAda(con, "profil_template")) CKPNLibrary.Panel.ProfilTemplate.BacaTersimpan(con, tanggal, h);

                // ---- Penyesuaian tersimpan saat ini (acuan perhitungan berikutnya) ----
                var pnyInd = (Dictionary<string, object>)((Dictionary<string, object>)h["penyesuaian"])["individu"];
                using (var cmd = Database.Cmd(con, "SELECT no_kontrak, jaminan, biaya_jual, alasan, pengguna, waktu FROM penyesuaian_individu"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        pnyInd[rd.GetString(0)] = new Dictionary<string, object>
                        {
                            { "jaminan", Database.Dbl(rd[1]) }, { "biaya", Database.Dbl(rd[2]) }, { "alasan", Str(rd[3]) },
                            { "pengguna", Str(rd[4]) }, { "waktu", Str(rd[5]) }
                        };
                var pnyCs = (Dictionary<string, object>)((Dictionary<string, object>)h["penyesuaian"])["lgdcs"];
                using (var cmd = Database.Cmd(con, "SELECT no_rek, jenis, nilai_agunan, recovery, alasan, pengguna, waktu FROM penyesuaian_lgdcs"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        pnyCs[rd.GetString(0)] = new Dictionary<string, object>
                        {
                            { "jenis", Str(rd[1]) }, { "nilaiAgunan", Database.Dbl(rd[2]) }, { "recovery", Database.Dbl(rd[3]) },
                            { "alasan", Str(rd[4]) }, { "pengguna", Str(rd[5]) }, { "waktu", Str(rd[6]) }
                        };
            }
            return h;
        }

        // =================================================================
        // Simpan / final / buka kembali / hapus
        // =================================================================
        /// <param name="status">"Draf" atau "Final"</param>
        public static Dictionary<string, object> Simpan(string jenis, string tanggal, string kunci, object data, string status)
        {
            CekJenis(jenis);
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrEmpty(tanggal)) throw new InvalidOperationException("Periode belum dipilih.");
            kunci = jenis == "ringkasan" ? "" : (kunci ?? "").Trim();
            if (jenis != "ringkasan" && kunci.Length == 0) throw new InvalidOperationException("Kunci memo (CIF / no. rekening) kosong.");
            if (status != "Final") status = "Draf";
            string isi = _json.Serialize(data ?? new Dictionary<string, object>());
            if (isi.Length > BatasData) throw new InvalidOperationException("Isi memo terlalu besar (maks. " + (BatasData / 1024) + " KB).");

            string now = Database.Sekarang(), user = Environment.UserName;
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                object st = Database.Scalar(con, "SELECT status FROM memo WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2", jenis, tanggal, kunci);
                if (st != null && Convert.ToString(st) == "Final")
                    throw new InvalidOperationException("Memo sudah Final. Buka kembali terlebih dahulu (beralasan) bila perlu diubah.");
                if (st == null)
                    Database.Exec(con,
                        "INSERT INTO memo(jenis,tanggal,kunci,data,status,dibuat_oleh,dibuat_waktu,diubah_oleh,diubah_waktu,final_oleh,final_waktu) " +
                        "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p5,@p6,@p7,@p8)",
                        jenis, tanggal, kunci, isi, status, user, now,
                        status == "Final" ? (object)user : DBNull.Value, status == "Final" ? (object)now : DBNull.Value);
                else
                    Database.Exec(con,
                        "UPDATE memo SET data=@p3, status=@p4, diubah_oleh=@p5, diubah_waktu=@p6, final_oleh=@p7, final_waktu=@p8 " +
                        "WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2",
                        jenis, tanggal, kunci, isi, status, user, now,
                        status == "Final" ? (object)user : DBNull.Value, status == "Final" ? (object)now : DBNull.Value);
                if (status == "Final")
                    Database.CatatAktivitas(con, tanggal, "memo-final", jenis + (kunci.Length > 0 ? " " + kunci : ""));
                tx.Commit();
            }
            LogProses.CatatPanel(tanggal, "Dokumen", NamaJenis(jenis) + (kunci.Length > 0 ? " " + kunci : ""),
                status == "Final" ? LogProses.OK : LogProses.Info, LogProses.R().Tambah("Status", status));
            return Satu(jenis, tanggal, kunci);
        }

        public static Dictionary<string, object> BukaKembali(string jenis, string tanggal, string kunci, string alasan)
        {
            CekJenis(jenis);
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan membuka kembali wajib diisi.");
            kunci = jenis == "ringkasan" ? "" : (kunci ?? "").Trim();
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                Database.Exec(con,
                    "UPDATE memo SET status='Draf', diubah_oleh=@p3, diubah_waktu=@p4 WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2",
                    jenis, tanggal, kunci, Environment.UserName, Database.Sekarang());
                Database.CatatAktivitas(con, tanggal, "memo-buka", jenis + (kunci.Length > 0 ? " " + kunci : "") + " · " + alasan.Trim());
                tx.Commit();
            }
            LogProses.CatatPanel(tanggal, "Dokumen", NamaJenis(jenis) + (kunci.Length > 0 ? " " + kunci : ""), LogProses.Info,
                LogProses.R().Tambah("Status", "dibuka kembali").Tambah("Alasan", alasan.Trim()));
            return Satu(jenis, tanggal, kunci);
        }

        public static void Hapus(string jenis, string tanggal, string kunci)
        {
            CekJenis(jenis);
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            kunci = jenis == "ringkasan" ? "" : (kunci ?? "").Trim();
            using (var con = Database.Buka(true))
            {
                object st = Database.Scalar(con, "SELECT status FROM memo WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2", jenis, tanggal, kunci);
                if (st != null && Convert.ToString(st) == "Final") throw new InvalidOperationException("Memo Final tidak dapat dihapus. Buka kembali terlebih dahulu.");
                Database.Exec(con, "DELETE FROM memo WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2", jenis, tanggal, kunci);
            }
        }

        private static Dictionary<string, object> Satu(string jenis, string tanggal, string kunci)
        {
            using (var con = Database.Buka(false))
            using (var cmd = Database.Cmd(con,
                "SELECT id, jenis, kunci, data, status, dibuat_oleh, dibuat_waktu, diubah_oleh, diubah_waktu, final_oleh, final_waktu " +
                "FROM memo WHERE jenis=@p0 AND tanggal=@p1 AND kunci=@p2", jenis, tanggal, kunci))
            using (var rd = cmd.ExecuteReader())
                return rd.Read() ? BacaMemo(rd) : null;
        }

        // =================================================================
        // Util
        // =================================================================
        private static Dictionary<string, object> KosongMemo()
        {
            return new Dictionary<string, object>
            {
                { "ringkasan", null }, { "individu", new Dictionary<string, object>() }, { "lgdcs", new Dictionary<string, object>() }
            };
        }

        private static Dictionary<string, object> BacaMemo(SQLiteDataReader rd)
        {
            return new Dictionary<string, object>
            {
                { "id", rd.GetInt64(0) }, { "kunci", Str(rd[2]) }, { "data", Urai(Str(rd[3])) }, { "status", Str(rd[4]) },
                { "dibuatOleh", Str(rd[5]) }, { "dibuatWaktu", Str(rd[6]) }, { "diubahOleh", Str(rd[7]) }, { "diubahWaktu", Str(rd[8]) },
                { "finalOleh", Str(rd[9]) }, { "finalWaktu", Str(rd[10]) }
            };
        }

        private static object Urai(string json)
        {
            if (string.IsNullOrEmpty(json)) return new Dictionary<string, object>();
            try { return _json.DeserializeObject(json); } catch { return new Dictionary<string, object>(); }
        }

        private static void CekJenis(string jenis)
        {
            if (Array.IndexOf(Jenis, jenis) < 0) throw new ArgumentException("Jenis dokumen tidak dikenal: " + jenis);
        }

        private static string NamaJenis(string jenis)
        {
            return jenis == "individu" ? "Memo CKPN Individu" : jenis == "lgdcs" ? "Memo LGD CS" : "Ringkasan CKPN";
        }

        private static bool TabelAda(SQLiteConnection con, string tabel)
        {
            return Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@p0", tabel)) > 0;
        }

        private static bool KolomAda(SQLiteConnection con, string tabel, string kolom)
        {
            using (var cmd = Database.Cmd(con, "PRAGMA table_info(" + tabel + ")"))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    if (string.Equals(Convert.ToString(rd["name"]), kolom, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private static string Str(object v) { return v == null || v is DBNull ? "" : Convert.ToString(v); }
    }
}
