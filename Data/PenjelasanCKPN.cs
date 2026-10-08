using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Penjelasan perubahan CKPN antarperiode (Tahap 4g) — tab Analisis.
    ///
    /// Bahan yang dibandingkan antara periode terpilih dan periode sebelumnya:
    ///   1. CKPN (metode tahunan) dipisah Individu vs Kolektif, PPKA, LGD gabungan  → dari Riwayat
    ///   2. OS, EAD, OS Kol 2–5, OS NPF (Kol 3–5), NPF %                              → dari snapshot Overview Data
    ///   3. Kontrak CKPN Individu yang masuk, keluar, dan berubah nilainya          → dari hasil_individu
    ///
    /// Snapshot Overview disimpan per periode (tabel overview_periode): otomatis di awal perhitungan
    /// panel bila belum ada / template berubah (Tahap 5g), saat Ringkasan memuat ulang dari template,
    /// atau diambil dari file template lama lewat tombol di Analisis › Komposisi.
    /// Narasi (naik/turun karena apa) disusun di panel dari angka-angka ini.
    /// </summary>
    internal static class PenjelasanCKPN
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // =================================================================
        // Snapshot Overview per periode
        // =================================================================
        /// <summary>Simpan snapshot Overview. Diam-diam dilewati bila user bukan pengirim.</summary>
        public static bool SimpanOverview(string tanggal, Dictionary<string, object> data, string sumber)
        {
            string info;
            if (string.IsNullOrEmpty(tanggal) || data == null || !Database.BolehMenulis(out info)) return false;
            try
            {
                // Tahap 5g: cap file template (ukuran + waktu ubah) untuk mendeteksi template yang
                // diganti/diperbarui setelah snapshot disimpan (lihat StatusOverview).
                string cap = CapTemplate(data.ContainsKey("pathTemplate") ? Convert.ToString(data["pathTemplate"]) : "");
                if (cap != null) data["capTemplate"] = cap;
                using (var con = Database.Buka(true))
                    Database.Exec(con,
                        "INSERT OR REPLACE INTO overview_periode(tanggal,data,sumber,pengguna,waktu) VALUES(@p0,@p1,@p2,@p3,@p4)",
                        tanggal, _json.Serialize(data), sumber ?? "", Environment.UserName, Database.Sekarang());
                return true;
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Simpan snapshot overview " + tanggal + " gagal: " + ex.Message);
                return false;
            }
        }

        private static Dictionary<string, object> BacaOverview(SQLiteConnection con, string tanggal)
        {
            using (var cmd = Database.Cmd(con, "SELECT data, sumber, waktu, pengguna FROM overview_periode WHERE tanggal=@p0", tanggal))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) return null;
                var d = _json.Deserialize<Dictionary<string, object>>(rd.GetString(0));
                d["sumberSnapshot"] = rd[1] is DBNull ? "" : rd.GetString(1);
                d["waktuSnapshot"] = rd[2] is DBNull ? "" : rd.GetString(2);
                d["penggunaSnapshot"] = rd[3] is DBNull ? "" : rd.GetString(3);
                return d;
            }
        }

        // =================================================================
        // Tahap 5g: snapshot Overview tersimpan → dipakai tanpa membuka template lagi
        // =================================================================
        /// <summary>"ukuran|waktu ubah UTC" file template; null bila path kosong / file tidak ada.</summary>
        internal static string CapTemplate(string path)
        {
            try
            {
                if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path)) return null;
                var fi = new System.IO.FileInfo(path);
                return fi.Length + "|" + fi.LastWriteTimeUtc.Ticks;
            }
            catch { return null; }
        }

        /// <summary>
        /// Snapshot Overview tersimpan untuk satu periode (tab Ringkasan). Tidak membuka Excel.
        /// Hasil: { ada:false } atau data snapshot + sumberSnapshot/waktuSnapshot/penggunaSnapshot
        /// + templateBerubah (file template di path yang sama sudah berubah sejak snapshot disimpan).
        /// </summary>
        public static Dictionary<string, object> OverviewTersimpan(string tanggal)
        {
            if (string.IsNullOrEmpty(tanggal) || !Database.Ada) return new Dictionary<string, object> { { "ada", false } };
            Dictionary<string, object> d;
            using (var con = Database.Buka(false)) d = BacaOverview(con, tanggal);
            if (d == null) return new Dictionary<string, object> { { "ada", false } };
            d["ada"] = true;
            d["templateBerubah"] = TemplateBerubah(d, d.ContainsKey("pathTemplate") ? Convert.ToString(d["pathTemplate"]) : "");
            return d;
        }

        private static bool TemplateBerubah(Dictionary<string, object> snapshot, string pathTemplate)
        {
            string capLama = snapshot.ContainsKey("capTemplate") ? Convert.ToString(snapshot["capTemplate"]) : null;
            string capKini = CapTemplate(pathTemplate);
            // snapshot lama (sebelum Tahap 5g) tidak punya cap → dianggap masih berlaku
            return capLama != null && capKini != null && capLama != capKini;
        }

        /// <summary>
        /// Apakah snapshot Overview periode ini perlu diambil (ulang) dari template Master!D14 sebelum
        /// perhitungan (dipakai CKPNPipeline). Perlu bila: user pengirim, file template ada, dan
        /// (snapshot belum ada, snapshot dari file template lain, atau file template berubah).
        /// </summary>
        public static bool OverviewPerluDiambil(string tanggal, string pathTemplate, out string alasan)
        {
            alasan = null;
            string info;
            if (string.IsNullOrEmpty(tanggal) || string.IsNullOrEmpty(pathTemplate) || !System.IO.File.Exists(pathTemplate)) return false;
            if (!Database.BolehMenulis(out info)) return false;
            if (!Database.Ada) { alasan = "snapshot belum ada"; return true; }
            try
            {
                Dictionary<string, object> d;
                using (var con = Database.Buka(false)) d = BacaOverview(con, tanggal);
                if (d == null) { alasan = "snapshot belum ada"; return true; }
                string pathLama = d.ContainsKey("pathTemplate") ? Convert.ToString(d["pathTemplate"]) : "";
                if (!string.Equals(SafeFull(pathLama), SafeFull(pathTemplate), StringComparison.OrdinalIgnoreCase))
                {
                    alasan = "snapshot dari file lain (" + (d.ContainsKey("fileTemplate") ? d["fileTemplate"] : pathLama) + ")";
                    return true;
                }
                if (TemplateBerubah(d, pathTemplate)) { alasan = "file template berubah sejak snapshot disimpan"; return true; }
                return false;
            }
            catch (Exception ex)
            {
                // tabel overview_periode belum ada di database lama, dsb. → ambil saja
                alasan = "snapshot tidak terbaca (" + ex.Message + ")";
                return true;
            }
        }

        private static string SafeFull(string p)
        {
            try { return string.IsNullOrEmpty(p) ? "" : System.IO.Path.GetFullPath(p); } catch { return p ?? ""; }
        }

        // =================================================================
        // Data penjelasan
        // =================================================================
        /// <param name="tanggal">periode yang dijelaskan; kosong = periode terakhir yang punya kiriman</param>
        public static Dictionary<string, object> Data(string tanggal)
        {
            var rw = Riwayat.Data();
            var daftar = (List<object>)rw["periode"];
            var hasil = new Dictionary<string, object> { { "grup", rw["grup"] } };
            var tanggalList = new List<string>();
            foreach (Dictionary<string, object> p in daftar) tanggalList.Add((string)p["tanggal"]);
            hasil["daftarPeriode"] = tanggalList;
            if (daftar.Count == 0) return hasil;

            int i = string.IsNullOrEmpty(tanggal) ? daftar.Count - 1 : tanggalList.IndexOf(tanggal);
            if (i < 0) i = daftar.Count - 1;
            var kini = (Dictionary<string, object>)daftar[i];
            var lalu = i > 0 ? (Dictionary<string, object>)daftar[i - 1] : null;

            using (var con = Database.Buka(false))
            {
                hasil["kini"] = Rangkai(con, kini);
                hasil["lalu"] = lalu == null ? null : Rangkai(con, lalu);
                if (lalu != null)
                    hasil["individu"] = BandingIndividu(con, (string)lalu["tanggal"], (string)kini["tanggal"]);
            }
            string info;
            hasil["bolehMenulis"] = Database.BolehMenulis(out info);
            return hasil;
        }

        private static Dictionary<string, object> Rangkai(SQLiteConnection con, Dictionary<string, object> p)
        {
            return new Dictionary<string, object>
            {
                { "tanggal", p["tanggal"] }, { "riwayat", p }, { "overview", BacaOverview(con, (string)p["tanggal"]) }
            };
        }

        /// <summary>
        /// Kontrak CKPN Individu dari kiriman aktif dua periode: masuk (baru), keluar, dan berubah.
        /// Nilai = kolom penurunan nilai (CKPN individu per kontrak). kc = segmen KC kontrak.
        /// </summary>
        private static Dictionary<string, object> BandingIndividu(SQLiteConnection con, string tglLalu, string tglKini)
        {
            var a = Kontrak(con, tglLalu);
            var b = Kontrak(con, tglKini);
            var baris = new List<object>();
            foreach (var kv in b)
            {
                Dictionary<string, object> lama;
                a.TryGetValue(kv.Key, out lama);
                double kini = Convert.ToDouble(kv.Value["ckpn"]);
                double sebelum = lama == null ? 0 : Convert.ToDouble(lama["ckpn"]);
                baris.Add(new Dictionary<string, object>
                {
                    { "kontrak", kv.Key }, { "nama", kv.Value["nama"] }, { "kc", kv.Value["kc"] },
                    { "jenis", lama == null ? "masuk" : "tetap" }, { "lalu", sebelum }, { "kini", kini },
                    { "osLalu", lama == null ? 0 : lama["os"] }, { "osKini", kv.Value["os"] }
                });
            }
            foreach (var kv in a)
                if (!b.ContainsKey(kv.Key))
                    baris.Add(new Dictionary<string, object>
                    {
                        { "kontrak", kv.Key }, { "nama", kv.Value["nama"] }, { "kc", kv.Value["kc"] },
                        { "jenis", "keluar" }, { "lalu", kv.Value["ckpn"] }, { "kini", 0.0 },
                        { "osLalu", kv.Value["os"] }, { "osKini", 0.0 }
                    });
            return new Dictionary<string, object> { { "kontrak", baris } };
        }

        private static Dictionary<string, Dictionary<string, object>> Kontrak(SQLiteConnection con, string tanggal)
        {
            var hasil = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = Database.Cmd(con,
                "SELECT h.no_kontrak, h.nama, h.kc, IFNULL(h.penurunan_nilai,0), IFNULL(h.os,0) FROM hasil_individu h " +
                "JOIN run_grup g ON g.id=h.run_id JOIN periode p ON p.id=g.periode_id " +
                "WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0 AND h.no_kontrak IS NOT NULL", tanggal))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                {
                    string k = Convert.ToString(rd[0]).Trim();
                    if (k.Length == 0) continue;
                    Dictionary<string, object> ada;
                    if (hasil.TryGetValue(k, out ada))
                    {
                        // kontrak sama muncul di dua kiriman (seharusnya tidak) → dijumlah
                        ada["ckpn"] = Convert.ToDouble(ada["ckpn"]) + Convert.ToDouble(rd[3]);
                        continue;
                    }
                    hasil[k] = new Dictionary<string, object>
                    {
                        { "nama", rd[1] is DBNull ? "" : Convert.ToString(rd[1]) },
                        { "kc", rd[2] is DBNull ? "" : Convert.ToString(rd[2]) },
                        { "ckpn", Convert.ToDouble(rd[3]) }, { "os", Convert.ToDouble(rd[4]) }
                    };
                }
            return hasil;
        }
    }
}
