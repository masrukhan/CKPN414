using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Lihat & koreksi data tersimpan satu versi grup (Tahap 5b) — tanpa proses ulang di Excel.
    ///
    /// CKPN Individu (tabel A. CKPN - INDV per kontrak):
    ///   CKPN = IF(ada PN = "Ya", MAX(0, OS − (agunan − biaya penjualan)), 0)   — rumus kolom K sheet.
    ///   Agunan/biaya tidak memengaruhi urutan Top-N (berdasarkan OS) → Kolektif tetap.
    ///
    /// LGD CS (tabel B4.LGD-CS MACET):
    ///   LGD CS       = 1 − Σ realisasi (G) / Σ pokok awal (C)
    ///   LGD gabungan = 1 − (Σ G + recovery ER) / (Σ C + hapus buku ER)          — kolom E blok "LGD Weighted"
    ///   Sheet B. CKPN - KOL INDV memakai satu sel LGD (= LGD gabungan) di blok Net Flow maupun Migration,
    ///   sehingga CKPN Kolektif berbanding lurus dengan LGD:  Kolektif baru = Kolektif lama × LGD baru / LGD lama.
    ///
    /// Hasil koreksi disimpan sebagai VERSI BARU grup (versi lama tetap ada). Hanya periode itu yang berubah:
    /// penyesuaian tersimpan untuk perhitungan bulan berikutnya tidak disentuh.
    /// </summary>
    internal static class KoreksiRun
    {
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        // ================================================================
        // Model
        // ================================================================
        private class Run
        {
            public long Id, PeriodeId; public string KodeKC, Tanggal, StatusPeriode, Pengguna, Waktu, Catatan, ParamJson;
            public int Versi; public bool Aktif, Dihapus; public int? TopN;
            public Dictionary<string, double> Ringkasan = new Dictionary<string, double>();
            public List<Indv> Individu = new List<Indv>();
            public List<Cs> LgdCs = new List<Cs>();
        }

        private class Indv
        {
            public long Id; public int Urut; public string Kc, Cif, Nama, Kontrak, AdaPN;
            public double Os, Jaminan, Biaya, Penurunan; public bool Disesuaikan, Diubah;
            public double Hitung() { return Ya(AdaPN) ? Math.Max(0, Os - (Jaminan - Biaya)) : 0; }
        }

        private class Cs
        {
            public long Id; public string Rek, Nama, ThnSerah, ThnEks, Sumber;
            public double Pokok, Agunan, Recovery, Shortfall; public bool Diubah;
        }

        /// <summary>Total LGD yang diperlukan untuk menghitung ulang LGD gabungan.</summary>
        private class InfoLgd
        {
            public double WoCs, RecCs, WoEr, RecEr;
            public double? LgdCs, LgdEr, LgdGab;
            public string SumberEr;      // "tersimpan" | "turunan" | null
            public string Masalah;       // alasan koreksi LGD CS tidak tersedia
            public double? Haircut;
        }

        private static bool Ya(string s) { return string.Equals((s ?? "").Trim(), "Ya", StringComparison.OrdinalIgnoreCase); }
        private static double Lgd(double wo, double rec) { return wo == 0 ? 0 : 1 - rec / wo; }   // IFERROR(1-rec/wo,0)

        // ================================================================
        // Baca
        // ================================================================
        private static Run Muat(SQLiteConnection con, long runId)
        {
            var r = new Run();
            using (var cmd = Database.Cmd(con,
                "SELECT g.id, g.periode_id, g.kode_kc, g.versi, g.aktif, g.dihapus, g.pengguna, g.waktu, g.catatan, g.parameter_json, " +
                "g.top_n, p.tanggal, p.status FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0", runId))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) throw new InvalidOperationException("Versi grup tidak ditemukan.");
                r.Id = rd.GetInt64(0); r.PeriodeId = rd.GetInt64(1); r.KodeKC = rd.GetString(2); r.Versi = rd.GetInt32(3);
                r.Aktif = rd.GetInt32(4) == 1; r.Dihapus = rd.GetInt32(5) == 1;
                r.Pengguna = rd[6] as string; r.Waktu = rd[7] as string; r.Catatan = rd[8] as string; r.ParamJson = rd[9] as string;
                r.TopN = rd[10] is DBNull ? (int?)null : Convert.ToInt32(rd[10]);
                r.Tanggal = rd.GetString(11); r.StatusPeriode = rd.GetString(12);
            }
            using (var cmd = Database.Cmd(con, "SELECT kunci, nilai FROM ringkasan WHERE run_id=@p0", runId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read()) if (!(rd[1] is DBNull)) r.Ringkasan[rd.GetString(0)] = rd.GetDouble(1);

            using (var cmd = Database.Cmd(con,
                "SELECT rowid, urut, kc, cif, nama, no_kontrak, os, ada_pn, jaminan, biaya_jual, penurunan_nilai, disesuaikan " +
                "FROM hasil_individu WHERE run_id=@p0 ORDER BY urut, rowid", runId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    r.Individu.Add(new Indv
                    {
                        Id = rd.GetInt64(0), Urut = rd[1] is DBNull ? 0 : Convert.ToInt32(rd[1]),
                        Kc = Str(rd[2]), Cif = Str(rd[3]), Nama = Str(rd[4]), Kontrak = Str(rd[5]),
                        Os = Num(rd[6]), AdaPN = Str(rd[7]), Jaminan = Num(rd[8]), Biaya = Num(rd[9]), Penurunan = Num(rd[10]),
                        Disesuaikan = !(rd[11] is DBNull) && Convert.ToInt32(rd[11]) == 1
                    });

            using (var cmd = Database.Cmd(con,
                "SELECT rowid, no_rek, nama, pokok_awal, nilai_agunan, thn_diserahkan, thn_eksekusi, recovery, shortfall, sumber " +
                "FROM hasil_lgdcs WHERE run_id=@p0 ORDER BY rowid", runId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read())
                    r.LgdCs.Add(new Cs
                    {
                        Id = rd.GetInt64(0), Rek = Str(rd[1]), Nama = Str(rd[2]), Pokok = Num(rd[3]), Agunan = Num(rd[4]),
                        ThnSerah = Str(rd[5]), ThnEks = Str(rd[6]), Recovery = Num(rd[7]), Shortfall = Num(rd[8]), Sumber = Str(rd[9])
                    });
            return r;
        }

        private static string Str(object o) { return o == null || o is DBNull ? "" : Convert.ToString(o, CultureInfo.InvariantCulture); }
        private static double Num(object o) { return o == null || o is DBNull ? 0 : Convert.ToDouble(o, CultureInfo.InvariantCulture); }
        private static double? Rg(Run r, string k) { double v; return r.Ringkasan.TryGetValue(k, out v) ? v : (double?)null; }

        /// <summary>
        /// Total hapus buku & recovery LGD ER. Versi yang disimpan mulai Tahap 5b menyimpannya langsung;
        /// untuk versi lama diturunkan dari tiga persentase yang tersimpan (LGD CS, ER, gabungan) dan total CS:
        ///   g = 1 − LGD gabungan, a = 1 − LGD ER  →  WO_ER = (g·WO_CS − Rec_CS) / (a − g),  Rec_ER = a·WO_ER
        /// </summary>
        private static InfoLgd HitungInfoLgd(Run r)
        {
            var i = new InfoLgd
            {
                LgdCs = Rg(r, "lgd_cs"), LgdEr = Rg(r, "lgd_er"), LgdGab = Rg(r, "lgd_gabungan"),
                Haircut = Rg(r, "lgd_cs_haircut") ?? HaircutDariParameter(r.ParamJson)
            };
            foreach (var c in r.LgdCs) { i.WoCs += c.Pokok; i.RecCs += c.Recovery; }

            double? woEr = Rg(r, "lgd_er_wo"), recEr = Rg(r, "lgd_er_rec");
            if (woEr.HasValue && recEr.HasValue)
            {
                i.WoEr = woEr.Value; i.RecEr = recEr.Value; i.SumberEr = "tersimpan";
            }
            else if (i.LgdEr.HasValue && i.LgdGab.HasValue)
            {
                double a = 1 - i.LgdEr.Value, g = 1 - i.LgdGab.Value;
                if (Math.Abs(a - g) > 1e-9)
                {
                    double w = (g * i.WoCs - i.RecCs) / (a - g);
                    if (w >= 0 && !double.IsNaN(w) && !double.IsInfinity(w))
                    {
                        i.WoEr = w; i.RecEr = a * w; i.SumberEr = "turunan";
                    }
                }
                else if (i.WoCs == 0)
                    i.Masalah = "Versi ini tidak punya baris LGD CS dan total LGD ER tidak tersimpan. Simpan ulang grup ini sekali dari sheet agar totalnya tercatat.";
            }
            double? acuan = Rg(r, "acuan_versi");
            if (acuan.HasValue)
            {
                i.Masalah = "PD & LGD versi ini memakai acuan Desember (v" + acuan.Value.ToString("0", CultureInfo.InvariantCulture) +
                            "); LGD CS tidak dihitung bulan ini. Koreksi LGD CS dilakukan di versi Desember tersebut.";
                return i;
            }
            if (i.SumberEr == null && i.Masalah == null)
                i.Masalah = "Total LGD ER (hapus buku & recovery) versi ini tidak dapat diketahui. Simpan ulang grup ini sekali dari sheet agar totalnya tercatat.";

            // Pemeriksaan silang: data baris harus menghasilkan LGD yang sama dengan sheet saat disimpan
            if (i.Masalah == null && i.LgdCs.HasValue && Math.Abs(Lgd(i.WoCs, i.RecCs) - i.LgdCs.Value) > 1e-6)
                i.Masalah = "Baris LGD CS tersimpan menghasilkan LGD CS " + (Lgd(i.WoCs, i.RecCs) * 100).ToString("0.00", CultureInfo.InvariantCulture) +
                            "%, berbeda dengan sheet saat disimpan (" + (i.LgdCs.Value * 100).ToString("0.00", CultureInfo.InvariantCulture) + "%).";
            if (i.Masalah == null && i.LgdGab.HasValue &&
                Math.Abs(Lgd(i.WoCs + i.WoEr, i.RecCs + i.RecEr) - i.LgdGab.Value) > 1e-6)
                i.Masalah = "LGD gabungan tidak dapat direkonstruksi dari data tersimpan.";
            if (i.Masalah == null && !i.LgdGab.HasValue)
                i.Masalah = "LGD gabungan tidak tersimpan pada versi ini.";
            return i;
        }

        private static double? HaircutDariParameter(string paramJson)
        {
            if (string.IsNullOrEmpty(paramJson)) return null;
            var m = Regex.Match(paramJson, @"haircut\s+([0-9]+(?:[.,][0-9]+)?)\s*%", RegexOptions.IgnoreCase);
            double v;
            if (m.Success && double.TryParse(m.Groups[1].Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                return v / 100.0;
            return null;
        }

        // ================================================================
        // Untuk panel
        // ================================================================
        public static Dictionary<string, object> Data(long runId)
        {
            Run r;
            using (var con = Database.Buka(false)) r = Muat(con, runId);
            var lgd = HitungInfoLgd(r);
            string info;
            bool bolehMenulis = Database.BolehMenulis(out info);

            var ind = new List<object>();
            foreach (var b in r.Individu)
                ind.Add(new Dictionary<string, object>
                {
                    { "id", b.Id }, { "urut", b.Urut }, { "kc", b.Kc }, { "cif", b.Cif }, { "nama", b.Nama }, { "kontrak", b.Kontrak },
                    { "os", b.Os }, { "pn", b.AdaPN }, { "jaminan", b.Jaminan }, { "biaya", b.Biaya },
                    { "ckpn", b.Penurunan }, { "ckpnHitung", b.Hitung() }, { "disesuaikan", b.Disesuaikan }
                });
            var cs = new List<object>();
            foreach (var c in r.LgdCs)
                cs.Add(new Dictionary<string, object>
                {
                    { "id", c.Id }, { "rek", c.Rek }, { "nama", c.Nama }, { "pokok", c.Pokok }, { "agunan", c.Agunan },
                    { "thnSerah", c.ThnSerah }, { "thnEks", c.ThnEks }, { "recovery", c.Recovery }, { "shortfall", c.Shortfall },
                    { "sumber", c.Sumber }
                });

            var rg = new Dictionary<string, object>();
            foreach (var kv in r.Ringkasan) rg[kv.Key] = kv.Value;

            string tolak = null;
            if (!bolehMenulis) tolak = info;
            else if (r.Dihapus) tolak = "Versi ini sudah dihapus.";
            else if (!r.Aktif) tolak = "Hanya versi aktif yang dapat dikoreksi.";
            else if (r.StatusPeriode == "Final") tolak = "Periode " + r.Tanggal + " sudah dikunci (Final). Buka kunci terlebih dahulu.";
            else if (r.Individu.Count == 0 && r.LgdCs.Count == 0) tolak = "Versi ini tidak menyimpan tabel Individu maupun LGD CS.";

            return new Dictionary<string, object>
            {
                { "runId", r.Id }, { "kodeKC", r.KodeKC }, { "versi", r.Versi }, { "tanggal", r.Tanggal },
                { "statusPeriode", r.StatusPeriode }, { "aktif", r.Aktif }, { "pengguna", r.Pengguna }, { "waktu", r.Waktu },
                { "catatan", r.Catatan }, { "ringkasan", rg }, { "individu", ind }, { "lgdcs", cs },
                { "lgd", new Dictionary<string, object>
                    {
                        { "woCs", lgd.WoCs }, { "recCs", lgd.RecCs }, { "woEr", lgd.WoEr }, { "recEr", lgd.RecEr },
                        { "lgdCs", lgd.LgdCs }, { "lgdEr", lgd.LgdEr }, { "lgdGab", lgd.LgdGab },
                        { "sumberEr", lgd.SumberEr }, { "masalah", lgd.Masalah }, { "haircut", lgd.Haircut }
                    } },
                { "bolehKoreksi", tolak == null }, { "alasanTolak", tolak }
            };
        }

        // ================================================================
        // Simpan koreksi → versi baru
        // ================================================================
        /// <param name="individu">[{id, jaminan, biaya}] — hanya baris yang diubah</param>
        /// <param name="lgdcs">{ubah:[{id, agunan, recovery}], hapus:[id], tambah:[{rek,nama,pokok,agunan,thnSerah,thnEks,recovery}]}</param>
        public static Dictionary<string, object> Simpan(long runId, object individu, object lgdcs, string alasan)
        {
            string info;
            if (!Database.BolehMenulis(out info)) throw new InvalidOperationException(info);
            if (string.IsNullOrWhiteSpace(alasan)) throw new InvalidOperationException("Alasan koreksi wajib diisi.");
            if (Panel.CKPNPipeline.SedangBerjalan || Panel.BatchGrup.Aktif)
                throw new InvalidOperationException("Perhitungan sedang berjalan. Coba lagi setelah selesai.");

            Run r;
            using (var con = Database.Buka(false)) r = Muat(con, runId);
            if (r.Dihapus || !r.Aktif) throw new InvalidOperationException("Hanya versi aktif yang dapat dikoreksi.");
            if (r.StatusPeriode == "Final") throw new InvalidOperationException("Periode " + r.Tanggal + " sudah dikunci (Final). Buka kunci terlebih dahulu.");

            // ---- a. Terapkan perubahan Individu ----
            double indLama = 0, indBaru = 0;
            foreach (var b in r.Individu) indLama += b.Penurunan;
            var indById = new Dictionary<long, Indv>();
            foreach (var b in r.Individu) indById[b.Id] = b;
            int nInd = 0;
            foreach (var o in Daftar(individu))
            {
                var d = o as Dictionary<string, object>;
                if (d == null) continue;
                Indv b;
                if (!indById.TryGetValue(Long(d, "id"), out b)) throw new InvalidOperationException("Baris Individu tidak ditemukan.");
                double jam = Angka(d, "jaminan", b.Jaminan), bia = Angka(d, "biaya", b.Biaya);
                if (jam < 0 || bia < 0) throw new InvalidOperationException("Nilai agunan/biaya kontrak " + b.Kontrak + " tidak boleh negatif.");
                if (Math.Abs(jam - b.Jaminan) < 0.005 && Math.Abs(bia - b.Biaya) < 0.005) continue;
                b.Jaminan = jam; b.Biaya = bia; b.Diubah = true; nInd++;
            }
            foreach (var b in r.Individu)
            {
                if (b.Diubah) b.Penurunan = b.Hitung();
                indBaru += b.Penurunan;
            }

            // ---- b. Terapkan perubahan LGD CS ----
            InfoLgd lgd = HitungInfoLgd(r);
            var csDict = lgdcs as Dictionary<string, object>;
            int nCs = 0;
            var csById = new Dictionary<long, Cs>();
            foreach (var c in r.LgdCs) csById[c.Id] = c;
            if (csDict != null)
            {
                foreach (var o in Daftar(Ambil(csDict, "ubah")))
                {
                    var d = o as Dictionary<string, object>;
                    if (d == null) continue;
                    Cs c;
                    if (!csById.TryGetValue(Long(d, "id"), out c)) throw new InvalidOperationException("Baris LGD CS tidak ditemukan.");
                    double agn = Angka(d, "agunan", c.Agunan), rec = Angka(d, "recovery", c.Recovery);
                    if (agn < 0 || rec < 0) throw new InvalidOperationException("Nilai agunan/realisasi rekening " + c.Rek + " tidak boleh negatif.");
                    if (Math.Abs(agn - c.Agunan) < 0.005 && Math.Abs(rec - c.Recovery) < 0.005) continue;
                    c.Agunan = agn; c.Recovery = rec; c.Diubah = true; nCs++;
                }
                foreach (var o in Daftar(Ambil(csDict, "hapus")))
                {
                    long id = Convert.ToInt64(o, CultureInfo.InvariantCulture);
                    Cs c;
                    if (csById.TryGetValue(id, out c) && r.LgdCs.Remove(c)) nCs++;
                }
                foreach (var o in Daftar(Ambil(csDict, "tambah")))
                {
                    var d = o as Dictionary<string, object>;
                    if (d == null) continue;
                    string rek = Convert.ToString(Ambil(d, "rek") ?? "").Trim();
                    if (rek.Length == 0) throw new InvalidOperationException("No. rekening baris LGD CS baru wajib diisi.");
                    foreach (var c0 in r.LgdCs)
                        if (string.Equals(c0.Rek, rek, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("Rekening " + rek + " sudah ada di tabel LGD CS.");
                    double pokok = Angka(d, "pokok", 0);
                    if (pokok <= 0) throw new InvalidOperationException("Pokok awal rekening " + rek + " wajib diisi.");
                    r.LgdCs.Add(new Cs
                    {
                        Id = 0, Rek = rek, Nama = Convert.ToString(Ambil(d, "nama") ?? "").Trim(), Pokok = pokok,
                        Agunan = Angka(d, "agunan", 0), Recovery = Angka(d, "recovery", 0),
                        ThnSerah = Convert.ToString(Ambil(d, "thnSerah") ?? "").Trim(), ThnEks = Convert.ToString(Ambil(d, "thnEks") ?? "").Trim(),
                        Sumber = "koreksi", Diubah = true
                    });
                    nCs++;
                }
            }
            if (nInd == 0 && nCs == 0) throw new InvalidOperationException("Tidak ada perubahan untuk disimpan.");
            if (nCs > 0 && lgd.Masalah != null) throw new InvalidOperationException("Koreksi LGD CS tidak tersedia: " + lgd.Masalah);

            // ---- c. Hitung ulang ringkasan ----
            var rgBaru = new Dictionary<string, double>(r.Ringkasan);
            double dInd = indBaru - indLama;
            Tambah(rgBaru, "nf_individu", dInd);
            Tambah(rgBaru, "mig_individu", dInd);

            double faktor = 1;
            double woCsBaru = 0, recCsBaru = 0;
            foreach (var c in r.LgdCs)
            {
                c.Shortfall = c.Recovery > c.Pokok ? 0 : c.Pokok - c.Recovery;   // =IF(G>C,0,C-G)
                woCsBaru += c.Pokok; recCsBaru += c.Recovery;
            }
            if (nCs > 0)
            {
                double gabLama = Lgd(lgd.WoCs + lgd.WoEr, lgd.RecCs + lgd.RecEr);
                double gabBaru = Lgd(woCsBaru + lgd.WoEr, recCsBaru + lgd.RecEr);
                if (gabLama <= 0)
                    throw new InvalidOperationException("LGD gabungan versi ini 0%, sehingga CKPN Kolektif tidak dapat disesuaikan secara proporsional. Hitung ulang grup dari sheet.");
                faktor = gabBaru / gabLama;
                rgBaru["lgd_cs"] = Lgd(woCsBaru, recCsBaru);
                rgBaru["lgd_gabungan"] = gabBaru;
                foreach (var k in new[] { "nf_kolektif", "mig_kolektif" })
                {
                    double v;
                    if (rgBaru.TryGetValue(k, out v)) rgBaru[k] = v * faktor;
                }
            }
            rgBaru["lgd_cs_wo"] = woCsBaru;
            rgBaru["lgd_cs_rec"] = recCsBaru;
            if (lgd.SumberEr != null) { rgBaru["lgd_er_wo"] = lgd.WoEr; rgBaru["lgd_er_rec"] = lgd.RecEr; }
            if (lgd.Haircut.HasValue) rgBaru["lgd_cs_haircut"] = lgd.Haircut.Value;
            foreach (var p in new[] { "nf", "mig" })
            {
                double ind, kol;
                if (rgBaru.TryGetValue(p + "_individu", out ind) && rgBaru.TryGetValue(p + "_kolektif", out kol))
                    rgBaru[p + "_total"] = ind + kol;                    // Summary B8 / B15 = SUM(individu, kolektif)
            }
            rgBaru["koreksi_dari_versi"] = r.Versi;

            // ---- d. Tulis versi baru ----
            Database.Cadangkan();
            long runBaru; int versiBaru;
            string ringkas = "Individu " + nInd + " kontrak, LGD CS " + nCs + " baris";
            using (var con = Database.Buka(true))
            using (var tx = con.BeginTransaction())
            {
                Periode.PastikanTerbuka(con, r.Tanggal);
                object aktifKini = Database.Scalar(con, "SELECT aktif FROM run_grup WHERE id=@p0 AND dihapus=0", r.Id);
                if (aktifKini == null || Convert.ToInt32(aktifKini) != 1)
                    throw new InvalidOperationException("Versi ini sudah tidak aktif (ada versi lebih baru). Muat ulang data.");

                string now = Database.Sekarang();
                versiBaru = Convert.ToInt32(Database.Scalar(con,
                    "SELECT COALESCE(MAX(versi),0)+1 FROM run_grup WHERE periode_id=@p0 AND kode_kc=@p1", r.PeriodeId, r.KodeKC));
                Database.Exec(con, "UPDATE run_grup SET aktif=0 WHERE periode_id=@p0 AND kode_kc=@p1", r.PeriodeId, r.KodeKC);
                Database.Exec(con,
                    "INSERT INTO run_grup(periode_id,kode_kc,versi,aktif,status,pengguna,komputer,waktu,versi_addin,parameter_json,catatan,top_n) " +
                    "VALUES(@p0,@p1,@p2,1,'Final',@p3,@p4,@p5,@p6,@p7,@p8,@p9)",
                    r.PeriodeId, r.KodeKC, versiBaru, Environment.UserName, Environment.MachineName, now,
                    typeof(KoreksiRun).Assembly.GetName().Version.ToString(), r.ParamJson,
                    "Koreksi panel dari v" + r.Versi + ": " + alasan.Trim(), r.TopN.HasValue ? (object)r.TopN.Value : DBNull.Value);
                runBaru = con.LastInsertRowId;

                foreach (var kv in rgBaru)
                    Database.Exec(con, "INSERT INTO ringkasan(run_id,kunci,nilai) VALUES(@p0,@p1,@p2)", runBaru, kv.Key, kv.Value);

                foreach (var b in r.Individu)
                    Database.Exec(con,
                        "INSERT INTO hasil_individu(run_id,urut,kc,cif,nama,no_kontrak,os,ada_pn,jaminan,biaya_jual,penurunan_nilai,disesuaikan) " +
                        "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11)",
                        runBaru, b.Urut, b.Kc, b.Cif, b.Nama, b.Kontrak, b.Os, b.AdaPN, b.Jaminan, b.Biaya, b.Penurunan,
                        (b.Disesuaikan || b.Diubah) ? 1 : 0);
                // Tahap 6: bukti objektif penurunan nilai ikut versi lama
                Database.Exec(con,
                    "UPDATE hasil_individu SET " +
                    "kualitas=(SELECT o.kualitas FROM hasil_individu o WHERE o.run_id=@p1 AND o.no_kontrak=hasil_individu.no_kontrak LIMIT 1), " +
                    "hari_tunggakan=(SELECT o.hari_tunggakan FROM hasil_individu o WHERE o.run_id=@p1 AND o.no_kontrak=hasil_individu.no_kontrak LIMIT 1), " +
                    "restruktur=(SELECT o.restruktur FROM hasil_individu o WHERE o.run_id=@p1 AND o.no_kontrak=hasil_individu.no_kontrak LIMIT 1), " +
                    "dasar_pn=(SELECT o.dasar_pn FROM hasil_individu o WHERE o.run_id=@p1 AND o.no_kontrak=hasil_individu.no_kontrak LIMIT 1) " +
                    "WHERE run_id=@p0", runBaru, r.Id);

                foreach (var c in r.LgdCs)
                    Database.Exec(con,
                        "INSERT INTO hasil_lgdcs(run_id,no_rek,nama,pokok_awal,nilai_agunan,thn_diserahkan,thn_eksekusi,recovery,shortfall,sumber) " +
                        "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9)",
                        runBaru, c.Rek, c.Nama, c.Pokok, c.Agunan, c.ThnSerah, c.ThnEks, c.Recovery, c.Shortfall,
                        c.Diubah && c.Sumber != "koreksi" ? c.Sumber + "+koreksi" : c.Sumber);

                // Bahan analisis PD ikut disalin; LGD weighted disesuaikan bila berubah
                using (var cmd = Database.Cmd(con, "SELECT jenis, data FROM analisis_pd WHERE run_id=@p0", r.Id))
                using (var rd = cmd.ExecuteReader())
                {
                    var salin = new List<KeyValuePair<string, string>>();
                    while (rd.Read()) salin.Add(new KeyValuePair<string, string>(rd.GetString(0), rd[1] as string));
                    rd.Close();
                    foreach (var kv in salin)
                    {
                        string data = kv.Value;
                        if (kv.Key == RincianCkpn.Jenis && !string.IsNullOrEmpty(data))
                        {
                            // Tahap 6c: rincian kolektif ikut koreksi (CKPN × faktor, LGD baru, CKPN Individu baru)
                            double indV;
                            data = RincianCkpn.SesuaikanKoreksi(data, faktor, nCs > 0 ? (double?)rgBaru["lgd_gabungan"] : null,
                                nInd > 0 && rgBaru.TryGetValue("nf_individu", out indV) ? (double?)indV : null);
                        }
                        else if (nCs > 0 && !string.IsNullOrEmpty(data))
                        {
                            try
                            {
                                var dj = _json.Deserialize<Dictionary<string, object>>(data);
                                if (dj != null && dj.ContainsKey("lgdWeighted")) { dj["lgdWeighted"] = rgBaru["lgd_gabungan"]; data = _json.Serialize(dj); }
                            }
                            catch { }
                        }
                        Database.Exec(con, "INSERT INTO analisis_pd(run_id,jenis,data) VALUES(@p0,@p1,@p2)", runBaru, kv.Key, data);
                    }
                }

                Database.CatatAktivitas(con, r.Tanggal, "koreksi-versi",
                    r.KodeKC + " v" + r.Versi + " → v" + versiBaru + " | " + ringkas + " | " + alasan.Trim());
                tx.Commit();
            }

            LogProses.CatatPanel(r.Tanggal, "Koreksi data grup", r.KodeKC + " v" + r.Versi + " → v" + versiBaru, LogProses.OK,
                LogProses.R()
                    .Tambah("Kontrak Individu diubah", nInd)
                    .Tambah("Baris LGD CS diubah/dihapus/ditambah", nCs)
                    .Tambah("CKPN Individu", indLama.ToString("#,##0", LogProses.Id) + " → " + indBaru.ToString("#,##0", LogProses.Id))
                    .TambahBila(nCs > 0, "LGD gabungan", nCs > 0 ? FormatPersen(Rg(r, "lgd_gabungan")) + " → " + FormatPersen(rgBaru["lgd_gabungan"]) : null)
                    .TambahBila(nCs > 0, "Faktor Kolektif", faktor.ToString("0.000000", CultureInfo.InvariantCulture))
                    .Tambah("Total Net Flow", Format(Rg(r, "nf_total")) + " → " + Format(Get(rgBaru, "nf_total")))
                    .Tambah("Total Migration", Format(Rg(r, "mig_total")) + " → " + Format(Get(rgBaru, "mig_total")))
                    .Tambah("Alasan", alasan.Trim()));

            var keluar = new Dictionary<string, object>();
            foreach (var kv in rgBaru) keluar[kv.Key] = kv.Value;
            return new Dictionary<string, object>
            {
                { "runId", runBaru }, { "versi", versiBaru }, { "tanggal", r.Tanggal }, { "kodeKC", r.KodeKC },
                { "ringkasan", keluar }, { "faktorKolektif", faktor }, { "jumlahIndividu", nInd }, { "jumlahLgdCs", nCs }
            };
        }

        // ================================================================
        // Util
        // ================================================================
        private static void Tambah(Dictionary<string, double> d, string k, double delta)
        {
            double v;
            if (d.TryGetValue(k, out v)) d[k] = v + delta;
        }

        private static double? Get(Dictionary<string, double> d, string k) { double v; return d.TryGetValue(k, out v) ? v : (double?)null; }
        private static string Format(double? v) { return v.HasValue ? v.Value.ToString("#,##0", LogProses.Id) : "-"; }
        private static string FormatPersen(double? v) { return v.HasValue ? (v.Value * 100).ToString("0.00", LogProses.Id) + "%" : "-"; }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        private static IEnumerable Daftar(object o)
        {
            var e = o as IEnumerable;
            if (e == null || o is string) return new object[0];
            return e;
        }

        private static long Long(Dictionary<string, object> d, string k)
        {
            object v = Ambil(d, k);
            return v == null ? 0 : Convert.ToInt64(v, CultureInfo.InvariantCulture);
        }

        private static double Angka(Dictionary<string, object> d, string k, double bawaan)
        {
            object v = Ambil(d, k);
            if (v == null) return bawaan;
            if (v is string)
            {
                double x;
                return double.TryParse((string)v, NumberStyles.Float, CultureInfo.InvariantCulture, out x) ? x : bawaan;
            }
            return Convert.ToDouble(v, CultureInfo.InvariantCulture);
        }
    }
}
