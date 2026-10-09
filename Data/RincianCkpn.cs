using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Rincian CKPN Individu vs Kolektif per grup (Tahap 6c) — Analisis › Rincian.
    ///
    /// Saat Simpan grup, sheet "B. CKPN - KOL INDV" dibaca apa adanya (nilai, bukan rumus) dan disimpan
    /// sebagai JSON di analisis_pd (jenis 'kolektif', tanpa perubahan skema):
    ///   Net Flow   : blok yang dipakai Summary — bulanan D6:G19 (G20), setahun sekali U6:X19 (X20)
    ///                per bucket: EAD, PD, LGD, CKPN
    ///   Migration  : C38:F42 per kualitas: EAD, PD, LGD, CKPN (F43)
    ///   Individu   : C24 nilai tercatat, D24 estimasi arus kas (agunan − biaya), E24 CKPN
    ///   Acuan PD & LGD (mode setahun): Z3 periode, AA3 label
    /// Panel membandingkan dengan periode sebelumnya grup yang sama untuk kontrol rasio PD/LGD.
    /// </summary>
    internal static class RincianCkpn
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        public const string Jenis = "kolektif";
        private const string SheetKol = "B. CKPN - KOL INDV";

        // =================================================================
        // Baca workbook (dipanggil StagingGrup.Simpan)
        // =================================================================
        public static string BacaWorkbookJson(Excel.Workbook wb)
        {
            try { return _json.Serialize(BacaWorkbook(wb)); }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Rincian kolektif tidak tersimpan: " + ex.Message);
                return null;
            }
        }

        public static Dictionary<string, object> BacaWorkbook(Excel.Workbook wb)
        {
            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetKol);
            if (ws == null) return null;
            Excel.Worksheet master = ParameterMaster.CariSheet(wb, "Master");
            string d17 = master == null ? "" : Convert.ToString(((Excel.Range)master.Range["D17"]).Value2 ?? "");
            bool setahun = d17.IndexOf("setahun", StringComparison.OrdinalIgnoreCase) >= 0;

            // Net Flow: blok yang dipakai Summary B7 (C29 = setahun ? X20 : G20)
            string awal = setahun ? "U6" : "D6", akhir = setahun ? "X19" : "G19";
            object[,] nf = (object[,])((Excel.Range)ws.Range[awal, akhir]).Value2;
            var netflow = new List<object>();
            for (int i = 1; i <= 14; i++)
                netflow.Add(new Dictionary<string, object>
                {
                    { "bucket", AnalisisPD.LabelBucket[i - 1] }, { "ead", Angka(nf[i, 1]) }, { "pd", AngkaNull(nf[i, 2]) },
                    { "lgd", AngkaNull(nf[i, 3]) }, { "ckpn", Angka(nf[i, 4]) }
                });

            object[,] mg = (object[,])((Excel.Range)ws.Range["C38", "F42"]).Value2;
            var migrasi = new List<object>();
            for (int i = 1; i <= 5; i++)
                migrasi.Add(new Dictionary<string, object>
                {
                    { "kol", i }, { "ead", Angka(mg[i, 1]) }, { "pd", AngkaNull(mg[i, 2]) }, { "lgd", AngkaNull(mg[i, 3]) }, { "ckpn", Angka(mg[i, 4]) }
                });

            // Acuan Desember (Tahap 5e) berlaku bila Z3 = Master!C4
            string acuan = null;
            try
            {
                object z3 = ((Excel.Range)ws.Range["Z3"]).Value2, c4 = master == null ? null : ((Excel.Range)master.Range["C4"]).Value2;
                if (setahun && z3 is double && c4 is double && Math.Abs((double)z3 - (double)c4) < 0.5)
                    acuan = Convert.ToString(((Excel.Range)ws.Range["AA3"]).Value2 ?? "");
            }
            catch { }

            return new Dictionary<string, object>
            {
                { "versiData", 1 },
                { "mode", setahun ? "setahun" : "bulanan" }, { "acuan", acuan },
                { "netflow", netflow }, { "nfTotal", Angka(((Excel.Range)ws.Range[setahun ? "X20" : "G20"]).Value2) },
                { "migrasi", migrasi }, { "migTotal", Angka(((Excel.Range)ws.Range["F43"]).Value2) },
                { "individu", new Dictionary<string, object>
                    {
                        { "os", Angka(((Excel.Range)ws.Range["C24"]).Value2) },
                        { "arusKas", Angka(((Excel.Range)ws.Range["D24"]).Value2) },
                        { "ckpn", Angka(((Excel.Range)ws.Range["E24"]).Value2) }
                    } }
            };
        }

        public static void SimpanRun(SQLiteConnection con, long runId, string json)
        {
            if (string.IsNullOrEmpty(json)) return;
            Database.Exec(con, "INSERT OR REPLACE INTO analisis_pd(run_id, jenis, data) VALUES(@p0,@p1,@p2)", runId, Jenis, json);
        }

        /// <summary>Koreksi panel (Tahap 5b): kolektif × faktor, LGD = LGD gabungan baru, CKPN individu baru.</summary>
        public static string SesuaikanKoreksi(string json, double faktor, double? lgdBaru, double? individuBaru)
        {
            try
            {
                var d = _json.Deserialize<Dictionary<string, object>>(json);
                foreach (var blok in new[] { "netflow", "migrasi" })
                {
                    var list = d.ContainsKey(blok) ? d[blok] as System.Collections.ArrayList : null;
                    if (list == null) continue;
                    foreach (Dictionary<string, object> b in list)
                    {
                        b["ckpn"] = Convert.ToDouble(b["ckpn"]) * faktor;
                        if (lgdBaru.HasValue && b["lgd"] != null) b["lgd"] = lgdBaru.Value;
                    }
                }
                foreach (var k in new[] { "nfTotal", "migTotal" })
                    if (d.ContainsKey(k) && d[k] != null) d[k] = Convert.ToDouble(d[k]) * faktor;
                var ind = d.ContainsKey("individu") ? d["individu"] as Dictionary<string, object> : null;
                if (ind != null && individuBaru.HasValue) ind["ckpn"] = individuBaru.Value;
                d["koreksi"] = true;
                return _json.Serialize(d);
            }
            catch { return json; }
        }

        // =================================================================
        // Data panel: semua grup aktif satu periode + pembanding periode sebelumnya
        // =================================================================
        public static Dictionary<string, object> Data(string tanggal)
        {
            var h = new Dictionary<string, object> { { "tanggal", tanggal }, { "grup", new List<object>() } };
            if (string.IsNullOrEmpty(tanggal) || !Database.Ada) return h;
            var list = (List<object>)h["grup"];
            using (var con = Database.Buka(false))
            {
                var runs = new List<Dictionary<string, object>>();
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id, g.kode_kc, g.versi, g.top_n FROM run_grup g JOIN periode p ON p.id=g.periode_id " +
                    "WHERE p.tanggal=@p0 AND g.aktif=1 AND g.dihapus=0 ORDER BY g.kode_kc", tanggal))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        runs.Add(new Dictionary<string, object>
                        {
                            { "runId", rd.GetInt64(0) }, { "kodeKC", rd.GetString(1) }, { "versi", rd.GetInt32(2) },
                            { "topN", rd[3] is DBNull ? null : (object)Convert.ToInt32(rd[3]) }
                        });

                foreach (var r in runs)
                {
                    long id = (long)r["runId"];
                    r["kolektif"] = BacaJson(con, id, Jenis);
                    r["pd"] = BacaJson(con, id, "pd");
                    r["individu"] = StatistikIndividu(con, id);

                    // periode sebelumnya, grup (kombinasi KC) yang sama
                    using (var cmd = Database.Cmd(con,
                        "SELECT g.id, p.tanggal, g.versi FROM run_grup g JOIN periode p ON p.id=g.periode_id " +
                        "WHERE p.tanggal < @p0 AND g.kode_kc=@p1 AND g.aktif=1 AND g.dihapus=0 ORDER BY p.tanggal DESC LIMIT 1", tanggal, (string)r["kodeKC"]))
                    using (var rd = cmd.ExecuteReader())
                        if (rd.Read())
                        {
                            long idLalu = rd.GetInt64(0);
                            string tglLalu = rd.GetString(1);
                            int versiLalu = rd.GetInt32(2);
                            rd.Close();
                            r["lalu"] = new Dictionary<string, object>
                            {
                                { "tanggal", tglLalu }, { "versi", versiLalu }, { "kolektif", BacaJson(con, idLalu, Jenis) },
                                { "pd", BacaJson(con, idLalu, "pd") }, { "individu", StatistikIndividu(con, idLalu) }
                            };
                        }
                    list.Add(r);
                }
            }
            return h;
        }

        private static object BacaJson(SQLiteConnection con, long runId, string jenis)
        {
            object v = Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis=@p1", runId, jenis);
            if (v == null || v is DBNull) return null;
            try { return _json.DeserializeObject(Convert.ToString(v)); } catch { return null; }
        }

        // Ringkasan CKPN Individu satu kiriman + 10 kontrak dengan CKPN terbesar
        private static Dictionary<string, object> StatistikIndividu(SQLiteConnection con, long runId)
        {
            var h = new Dictionary<string, object>();
            using (var cmd = Database.Cmd(con,
                "SELECT COUNT(*), COUNT(DISTINCT cif), SUM(os), SUM(jaminan), SUM(biaya_jual), SUM(penurunan_nilai), " +
                "SUM(CASE WHEN ada_pn='Ya' THEN 1 ELSE 0 END), SUM(CASE WHEN ada_pn='Ya' THEN os ELSE 0 END), SUM(disesuaikan) " +
                "FROM hasil_individu WHERE run_id=@p0", runId))
            using (var rd = cmd.ExecuteReader())
                if (rd.Read())
                {
                    h["kontrak"] = Convert.ToInt32(rd[0]); h["debitur"] = Convert.ToInt32(rd[1]);
                    h["os"] = Database.Dbl(rd[2]); h["jaminan"] = Database.Dbl(rd[3]); h["biaya"] = Database.Dbl(rd[4]);
                    h["ckpn"] = Database.Dbl(rd[5]); h["kontrakPN"] = rd[6] is DBNull ? 0 : Convert.ToInt32(rd[6]);
                    h["osPN"] = Database.Dbl(rd[7]); h["disesuaikan"] = rd[8] is DBNull ? 0 : Convert.ToInt32(rd[8]);
                }
            var top = new List<object>();
            using (var cmd = Database.Cmd(con,
                "SELECT kc, cif, nama, no_kontrak, os, ada_pn, jaminan, biaya_jual, penurunan_nilai, disesuaikan FROM hasil_individu " +
                "WHERE run_id=@p0 ORDER BY penurunan_nilai DESC, os DESC LIMIT 10", runId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    top.Add(new Dictionary<string, object>
                    {
                        { "kc", rd[0] as string }, { "cif", rd[1] as string }, { "nama", rd[2] as string }, { "kontrak", rd[3] as string },
                        { "os", Database.Dbl(rd[4]) }, { "adaPN", rd[5] as string }, { "jaminan", Database.Dbl(rd[6]) },
                        { "biaya", Database.Dbl(rd[7]) }, { "ckpn", Database.Dbl(rd[8]) }, { "disesuaikan", !(rd[9] is DBNull) && Convert.ToInt32(rd[9]) == 1 }
                    });
            h["top"] = top;
            return h;
        }

        private static double Angka(object v)
        {
            if (v is double) return (double)v;
            double d;
            return v != null && double.TryParse(Convert.ToString(v), out d) ? d : 0;
        }

        private static double? AngkaNull(object v)
        {
            if (v is double) return (double)v;
            double d;
            return v != null && double.TryParse(Convert.ToString(v), out d) ? (double?)d : null;
        }
    }
}
