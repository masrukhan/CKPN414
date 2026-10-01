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
    /// Snapshot Overview disimpan per periode (tabel overview_periode) setiap kali tab Overview
    /// dimuat oleh pengirim, atau diambil dari file template lama lewat tombol di tab Analisis.
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
            using (var cmd = Database.Cmd(con, "SELECT data, sumber, waktu FROM overview_periode WHERE tanggal=@p0", tanggal))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) return null;
                var d = _json.Deserialize<Dictionary<string, object>>(rd.GetString(0));
                d["sumberSnapshot"] = rd[1] is DBNull ? "" : rd.GetString(1);
                d["waktuSnapshot"] = rd[2] is DBNull ? "" : rd.GetString(2);
                return d;
            }
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
