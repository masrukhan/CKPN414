using System;
using System.Collections.Generic;
using System.IO;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Pengelolaan isi database (Tahap 3d) — terutama untuk masa simulasi / parallel run:
    ///   - InfoDatabase      : ringkasan isi per periode
    ///   - HapusDataPeriode  : hapus permanen seluruh data satu periode (versi, hasil, konsolidasi, jurnal)
    ///   - KosongkanDatabase : hapus seluruh data perhitungan (opsional penyesuaian & susunan grup)
    ///
    /// Pengaman:
    ///   - hanya pengirim (config\pengirim.txt);
    ///   - konfirmasi harus diketik (tanggal periode / "HAPUS SEMUA");
    ///   - salinan arsip "arsip_*.db" dibuat lebih dulu di library\backup dan tidak ikut dirotasi;
    ///   - log_aktivitas dan log_penyesuaian TIDAK dihapus (jejak audit tetap ada).
    /// </summary>
    internal static class KelolaData
    {
        public const string KonfirmasiSemua = "HAPUS SEMUA";

        public static Dictionary<string, object> InfoDatabase()
        {
            string info;
            bool boleh = Database.BolehMenulis(out info);
            var hasil = new Dictionary<string, object>
            {
                { "adaDatabase", Database.Ada }, { "fileDatabase", AppPaths.FileDatabase },
                { "bolehMenulis", boleh }, { "infoPengirim", info }, { "konfirmasiSemua", KonfirmasiSemua }
            };
            if (!Database.Ada) return hasil;
            hasil["ukuranKB"] = new FileInfo(AppPaths.FileDatabase).Length / 1024;

            using (var con = Database.Buka(false))
            {
                var periode = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT p.tanggal, p.status, p.metode, " +
                    " (SELECT COUNT(*) FROM run_grup g WHERE g.periode_id=p.id), " +
                    " (SELECT COUNT(*) FROM run_grup g WHERE g.periode_id=p.id AND g.aktif=1 AND g.dihapus=0), " +
                    " (SELECT COUNT(*) FROM jurnal_ckpn j WHERE j.periode_id=p.id), " +
                    " (SELECT COUNT(*) FROM penyesuaian_individu x WHERE x.periode=p.tanggal) + " +
                    " (SELECT COUNT(*) FROM penyesuaian_lgdcs y WHERE y.periode=p.tanggal) " +
                    "FROM periode p ORDER BY p.tanggal DESC"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        periode.Add(new Dictionary<string, object>
                        {
                            { "tanggal", rd.GetString(0) }, { "status", rd.GetString(1) },
                            { "metode", rd[2] is DBNull ? "" : rd.GetString(2) },
                            { "jumlahVersi", Convert.ToInt32(rd[3]) }, { "grupAktif", Convert.ToInt32(rd[4]) },
                            { "adaJurnal", Convert.ToInt32(rd[5]) > 0 }, { "penyesuaianPeriode", Convert.ToInt32(rd[6]) }
                        });
                hasil["periode"] = periode;
                hasil["penyesuaianIndividu"] = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM penyesuaian_individu"));
                hasil["penyesuaianLgdCs"] = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM penyesuaian_lgdcs"));

                var susunan = new List<object>();
                using (var cmd = Database.Cmd(con,
                    "SELECT g.tahun, COUNT(*), IFNULL(t.metode,'') FROM susunan_grup g LEFT JOIN susunan_tahun t ON t.tahun=g.tahun " +
                    "GROUP BY g.tahun ORDER BY g.tahun DESC"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        susunan.Add(new Dictionary<string, object>
                        {
                            { "tahun", rd.GetInt32(0) }, { "jumlahGrup", Convert.ToInt32(rd[1]) }, { "metode", rd.GetString(2) }
                        });
                hasil["susunan"] = susunan;
            }
            return hasil;
        }

        /// <param name="konfirmasi">harus sama dengan tanggal periode (yyyy-MM-dd)</param>
        /// <param name="hapusPenyesuaian">hapus juga penyesuaian yang tercatat dibuat/diubah pada periode ini</param>
        /// <param name="hapusSnapshot">hapus juga file snapshot .xlsx periode ini di folder staging</param>
        public static Dictionary<string, object> HapusDataPeriode(string tanggal, string konfirmasi, string alasan,
                                                                  bool hapusPenyesuaian, bool hapusSnapshot)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrEmpty(tanggal)) throw new InvalidOperationException("Periode belum dipilih.");
            if ((konfirmasi ?? "").Trim() != tanggal)
                throw new InvalidOperationException("Konfirmasi tidak cocok. Ketik " + tanggal + " untuk menghapus periode ini.");
            alasan = string.IsNullOrWhiteSpace(alasan) ? "Hapus data simulasi" : alasan.Trim();

            string arsip = Database.Arsipkan("sebelum-hapus-" + tanggal);
            var snapshot = new List<string>();
            int nVersi, nPenyesuaian = 0;

            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                object idObj = Database.Scalar(con, "SELECT id FROM periode WHERE tanggal=@p0", tanggal);
                if (idObj == null || idObj is DBNull) throw new InvalidOperationException("Periode " + tanggal + " tidak ada di database.");
                long pid = Convert.ToInt64(idObj);

                using (var cmd = Database.Cmd(con, "SELECT snapshot FROM run_grup WHERE periode_id=@p0 AND snapshot IS NOT NULL", pid))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) if (!(rd[0] is DBNull)) snapshot.Add(Convert.ToString(rd[0]));

                nVersi = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM run_grup WHERE periode_id=@p0", pid));
                const string runs = "(SELECT id FROM run_grup WHERE periode_id=@p0)";
                Database.Exec(con, "DELETE FROM keputusan_metode WHERE periode_id=@p0", pid);
                Database.Exec(con, "DELETE FROM jurnal_ckpn WHERE periode_id=@p0", pid);
                Database.Exec(con, "DELETE FROM ringkasan WHERE run_id IN " + runs, pid);
                Database.Exec(con, "DELETE FROM analisis_pd WHERE run_id IN " + runs, pid);
                Database.Exec(con, "DELETE FROM hasil_individu WHERE run_id IN " + runs, pid);
                Database.Exec(con, "DELETE FROM hasil_lgdcs WHERE run_id IN " + runs, pid);
                Database.Exec(con, "DELETE FROM run_grup WHERE periode_id=@p0", pid);
                Database.Exec(con, "DELETE FROM periode WHERE id=@p0", pid);
                Database.Exec(con, "DELETE FROM overview_periode WHERE tanggal=@p0", tanggal);
                Database.Exec(con, "DELETE FROM memo WHERE tanggal=@p0", tanggal);   // Tahap 6: memo & dokumen periode
                Database.Exec(con, "DELETE FROM profil_template WHERE tanggal=@p0", tanggal);
                Database.Exec(con, "DELETE FROM profil_template_info WHERE tanggal=@p0", tanggal);

                if (hapusPenyesuaian)
                {
                    nPenyesuaian = Convert.ToInt32(Database.Scalar(con,
                        "SELECT (SELECT COUNT(*) FROM penyesuaian_individu WHERE periode=@p0) + " +
                        "(SELECT COUNT(*) FROM penyesuaian_lgdcs WHERE periode=@p0)", tanggal));
                    Database.Exec(con, "DELETE FROM penyesuaian_individu WHERE periode=@p0", tanggal);
                    Database.Exec(con, "DELETE FROM penyesuaian_lgdcs WHERE periode=@p0", tanggal);
                    Database.Exec(con,
                        "INSERT INTO log_penyesuaian(waktu,pengguna,periode,modul,kunci,aksi,sebelum,sesudah,alasan) " +
                        "VALUES(@p0,@p1,@p2,'semua','*','hapus-data-periode','',@p3,@p4)",
                        Database.Sekarang(), Environment.UserName, tanggal, nPenyesuaian + " penyesuaian dihapus", alasan);
                }

                Database.CatatAktivitas(con, tanggal, "hapus-data-periode",
                    nVersi + " versi" + (hapusPenyesuaian ? ", " + nPenyesuaian + " penyesuaian" : "") +
                    " | arsip: " + Path.GetFileName(arsip) + " | " + alasan);
                tx.Commit();
            }

            int nFile = hapusSnapshot ? HapusFileStaging(snapshot) : 0;
            Periode.ResetKonteksSimpan(tanggal);
            CatatanLog.Tulis("Data periode " + tanggal + " dihapus (" + nVersi + " versi). Arsip: " + arsip);

            return new Dictionary<string, object>
            {
                { "versi", nVersi }, { "penyesuaian", nPenyesuaian }, { "snapshot", nFile }, { "arsip", arsip }
            };
        }

        public static Dictionary<string, object> KosongkanDatabase(string konfirmasi, string alasan,
                                                                   bool sertakanPenyesuaian, bool sertakanSusunan, bool hapusSnapshot)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if ((konfirmasi ?? "").Trim() != KonfirmasiSemua)
                throw new InvalidOperationException("Konfirmasi tidak cocok. Ketik " + KonfirmasiSemua + " untuk mengosongkan database.");
            alasan = string.IsNullOrWhiteSpace(alasan) ? "Kosongkan database simulasi" : alasan.Trim();

            string arsip = Database.Arsipkan("sebelum-kosongkan");
            var snapshot = new List<string>();
            int nPeriode, nVersi;

            using (var con = Database.Buka(true))
            {
                using (var tx = con.BeginTransaction())
                {
                    using (var cmd = Database.Cmd(con, "SELECT snapshot FROM run_grup WHERE snapshot IS NOT NULL"))
                    using (var rd = cmd.ExecuteReader())
                        while (rd.Read()) if (!(rd[0] is DBNull)) snapshot.Add(Convert.ToString(rd[0]));
                    nPeriode = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM periode"));
                    nVersi = Convert.ToInt32(Database.Scalar(con, "SELECT COUNT(*) FROM run_grup"));

                    foreach (var t in new[] { "keputusan_metode", "jurnal_ckpn", "ringkasan", "analisis_pd", "hasil_individu", "hasil_lgdcs", "run_grup", "periode", "overview_periode", "memo", "profil_template", "profil_template_info" })
                        Database.Exec(con, "DELETE FROM " + t);
                    if (sertakanPenyesuaian)
                    {
                        Database.Exec(con, "DELETE FROM penyesuaian_individu");
                        Database.Exec(con, "DELETE FROM penyesuaian_lgdcs");
                        Database.Exec(con,
                            "INSERT INTO log_penyesuaian(waktu,pengguna,periode,modul,kunci,aksi,sebelum,sesudah,alasan) " +
                            "VALUES(@p0,@p1,'','semua','*','kosongkan-database','','Semua penyesuaian dihapus',@p2)",
                            Database.Sekarang(), Environment.UserName, alasan);
                    }
                    if (sertakanSusunan)
                    {
                        Database.Exec(con, "DELETE FROM susunan_grup");
                        Database.Exec(con, "DELETE FROM susunan_tahun");
                        Database.Exec(con, "DELETE FROM parameter_aba");   // Tahap 5h: ketetapan tahunan juga
                    }
                    Database.CatatAktivitas(con, "", "kosongkan-database",
                        nPeriode + " periode, " + nVersi + " versi" +
                        (sertakanPenyesuaian ? ", penyesuaian" : "") + (sertakanSusunan ? ", susunan grup" : "") +
                        " | arsip: " + Path.GetFileName(arsip) + " | " + alasan);
                    tx.Commit();
                }
                try { Database.Exec(con, "VACUUM"); } catch (Exception ex) { CatatanLog.Tulis("VACUUM: " + ex.Message); }
            }

            int nFile = hapusSnapshot ? HapusFileStaging(snapshot) : 0;
            Periode.ResetKonteksSimpan(null);
            CatatanLog.Tulis("Database dikosongkan (" + nPeriode + " periode). Arsip: " + arsip);

            return new Dictionary<string, object>
            {
                { "periode", nPeriode }, { "versi", nVersi }, { "snapshot", nFile }, { "arsip", arsip }
            };
        }

        /// <summary>Hapus file snapshot — hanya yang berada di dalam folder staging.</summary>
        private static int HapusFileStaging(List<string> daftar)
        {
            int n = 0;
            string akar = Path.GetFullPath(AppPaths.FolderStaging);
            foreach (var f in daftar)
            {
                try
                {
                    string full = Path.GetFullPath(f);
                    if (!full.StartsWith(akar, StringComparison.OrdinalIgnoreCase) || !File.Exists(full)) continue;
                    File.Delete(full);
                    n++;
                    string folder = Path.GetDirectoryName(full);
                    if (Directory.Exists(folder) && Directory.GetFileSystemEntries(folder).Length == 0 &&
                        !string.Equals(Path.GetFullPath(folder), akar, StringComparison.OrdinalIgnoreCase))
                        Directory.Delete(folder);
                }
                catch (Exception ex) { CatatanLog.Tulis("Hapus snapshot " + f + ": " + ex.Message); }
            }
            return n;
        }
    }
}
