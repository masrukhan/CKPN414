using System;
using System.Collections.Generic;
using System.Data.SQLite;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Riwayat antarperiode (Tahap 4) — pengganti file Dashboard CKPN.
    ///
    /// Per periode di database:
    ///   - total CKPN Net Flow, CKPN Migration, dan PPKA (grup sesuai susunan tahunnya + ABA);
    ///   - CKPN terpakai (metode tahunan; untuk periode Final memakai angka yang dikunci);
    ///   - selisih CKPN − PPKA dan jurnal yang tersimpan;
    ///   - rincian per grup (Net Flow, Migration, PPKA, LGD gabungan).
    ///
    /// Hanya kiriman AKTIF (versi terakhir, tidak dihapus) yang dihitung, dan hanya kiriman
    /// yang kombinasi KC-nya ada di susunan grup tahun tersebut — sama seperti konsolidasi,
    /// agar kiriman percobaan di luar susunan tidak menggandakan angka.
    /// </summary>
    internal static class Riwayat
    {
        public static Dictionary<string, object> Data()
        {
            var periode = new List<object>();
            var namaGrup = new List<string>();
            if (!Database.Ada)
                return new Dictionary<string, object> { { "periode", periode }, { "grup", namaGrup } };

            using (var con = Database.Buka(false))
            {
                var daftar = new List<object[]>();
                using (var cmd = Database.Cmd(con,
                    "SELECT id, tanggal, status, metode, ckpn_total, ppka_total FROM periode ORDER BY tanggal"))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read())
                        daftar.Add(new object[] { rd.GetInt64(0), rd.GetString(1), rd.GetString(2), rd[3], rd[4], rd[5] });

                foreach (var p in daftar)
                {
                    var baris = SatuPeriode(con, (long)p[0], (string)p[1], (string)p[2],
                                            p[3] is DBNull ? null : Convert.ToString(p[3]),
                                            Database.Dbl(p[4]), Database.Dbl(p[5]));
                    if (baris == null) continue;
                    foreach (Dictionary<string, object> g in (List<object>)baris["grup"])
                        if (!namaGrup.Contains((string)g["nama"])) namaGrup.Add((string)g["nama"]);
                    periode.Add(baris);
                }
            }
            return new Dictionary<string, object> { { "periode", periode }, { "grup", namaGrup } };
        }

        private static Dictionary<string, object> SatuPeriode(SQLiteConnection con, long periodeId, string tanggal,
            string status, string metodeFinal, double? ckpnFinal, double? ppkaFinal)
        {
            int tahun = int.Parse(tanggal.Substring(0, 4));
            int sumber;
            var susunan = Periode.Susunan(con, tahun, out sumber) ?? Periode.SusunanBawaan();
            var runs = Periode.DaftarRun(con, tanggal);
            var peng = Periode.Pengaturan(con, tahun);
            bool final = status == "Final";
            string metode = final && !string.IsNullOrEmpty(metodeFinal) ? metodeFinal : (peng == null ? null : peng.Metode);

            var grup = new List<object>();
            double nf = 0, mig = 0, ppka = 0;
            double nfInd = 0, nfKol = 0, migInd = 0, migKol = 0;   // Tahap 4g: pemisahan individu/kolektif
            int tersimpan = 0;
            Dictionary<string, object> ringkasTerbaru = null;
            string waktuTerbaru = "";
            foreach (var g in susunan)
            {
                var aktif = runs.Find(r => (string)r["kodeKC"] == g.KodeKC && (bool)r["aktif"] && !(bool)r["dihapus"]);
                if (aktif == null)
                {
                    grup.Add(new Dictionary<string, object> { { "nama", g.Nama }, { "kodeKC", g.KodeKC }, { "ada", false } });
                    continue;
                }
                tersimpan++;
                var rg = (Dictionary<string, object>)aktif["ringkasan"];
                double gNf = Periode.Nilai(rg, "nf_total"), gMig = Periode.Nilai(rg, "mig_total");
                double gPpka = Convert.ToDouble(aktif["ppkaGrup"]);
                nf += gNf; mig += gMig; ppka += gPpka;
                double gNfInd = Periode.Nilai(rg, "nf_individu"), gNfKol = Periode.Nilai(rg, "nf_kolektif");
                double gMigInd = Periode.Nilai(rg, "mig_individu"), gMigKol = Periode.Nilai(rg, "mig_kolektif");
                nfInd += gNfInd; nfKol += gNfKol; migInd += gMigInd; migKol += gMigKol;
                object lgd;
                rg.TryGetValue("lgd_gabungan", out lgd);
                grup.Add(new Dictionary<string, object>
                {
                    { "nama", g.Nama }, { "kodeKC", g.KodeKC }, { "ada", true }, { "versi", aktif["versi"] },
                    { "nf", gNf }, { "mig", gMig }, { "ppka", gPpka }, { "lgd", lgd },
                    { "nfInd", gNfInd }, { "nfKol", gNfKol }, { "migInd", gMigInd }, { "migKol", gMigKol }
                });
                string w = Convert.ToString(aktif["waktu"]);
                if (string.CompareOrdinal(w, waktuTerbaru) > 0) { waktuTerbaru = w; ringkasTerbaru = rg; }
            }
            if (tersimpan == 0) return null;   // periode tanpa kiriman aktif tidak ditampilkan

            double abaCkpn = Periode.Nilai(ringkasTerbaru, "aba_ckpn");
            double abaPpka = Periode.Nilai(ringkasTerbaru, "ppka_KC0500");
            double totNf = nf + abaCkpn, totMig = mig + abaCkpn, totPpka = ppka + abaPpka;

            double? ckpn = null;
            if (final && ckpnFinal.HasValue) ckpn = ckpnFinal;
            else if (metode == "nf") ckpn = totNf;
            else if (metode == "mig") ckpn = totMig;
            double ppkaTampil = final && ppkaFinal.HasValue ? ppkaFinal.Value : totPpka;

            // Jurnal yang tersimpan saat periode dikunci (Tahap 3c/3d)
            Dictionary<string, object> jurnal = null;
            using (var cmd = Database.Cmd(con,
                "SELECT jenis, SUM(IFNULL(ckpn,0) - IFNULL(ppka,0)), GROUP_CONCAT(IFNULL(akun_debit,''),'|'), GROUP_CONCAT(IFNULL(akun_kredit,''),'|') " +
                "FROM jurnal_ckpn WHERE periode_id=@p0 GROUP BY periode_id", periodeId))
            using (var rd = cmd.ExecuteReader())
                if (rd.Read())
                    jurnal = new Dictionary<string, object>
                    {
                        { "jenis", Convert.ToString(rd[0]) }, { "selisih", Database.Dbl(rd[1]) },
                        { "debit", Convert.ToString(rd[2]).Trim('|') }, { "kredit", Convert.ToString(rd[3]).Trim('|') }
                    };

            return new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "status", status }, { "metode", metode },
                { "lengkap", tersimpan == susunan.Count }, { "grupTersimpan", tersimpan }, { "jumlahGrup", susunan.Count },
                { "nf", totNf }, { "mig", totMig }, { "ppka", ppkaTampil }, { "ckpn", ckpn },
                { "selisih", ckpn.HasValue ? ckpn.Value - ppkaTampil : (double?)null },
                { "aba", new Dictionary<string, object> { { "ckpn", abaCkpn }, { "ppka", abaPpka } } },
                { "jurnal", jurnal }, { "grup", grup },
                { "nfInd", nfInd }, { "nfKol", nfKol }, { "migInd", migInd }, { "migKol", migKol }
            };
        }
    }
}
