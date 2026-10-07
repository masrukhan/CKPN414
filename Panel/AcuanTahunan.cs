using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.Web.Script.Serialization;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Mode "PD &amp; LGD dihitung setahun sekali" (Master!D17) — Tahap 5e.
    ///
    /// Untuk laporan Januari–November tahun Y, PD &amp; LGD yang dipakai adalah hasil Desember Y-1.
    /// Daripada menghitung ulang Net Flow (13 file), Migration, LGD ER, dan LGD CS setiap bulan,
    /// panel dapat memakai ACUAN: PD &amp; LGD dari versi aktif grup yang sama di periode Desember
    /// tahun lalu (tersimpan di database sejak Tahap 4: analisis_pd + ringkasan LGD).
    ///
    /// Acuan ditulis ke sheet B. CKPN - KOL INDV, kolom Z:AB (di luar tabel yang sudah ada):
    ///   Z3  = tanggal laporan (sama dengan Master!C4) — acuan hanya berlaku untuk periode ini
    ///   AA3 = label ("Desember 2025 v3"), AB3 = id versi Desember di database
    ///   Z6:Z19  PD Net Flow per bucket      → dipakai V6:V19 (blok tahunan)
    ///   AA6     LGD weighted                 → dipakai W6:W19 dan E38:E42
    ///   Z38:Z42 PD Migration per kualitas    → dipakai D38:D42
    /// Rumus di V/W/D/E diberi pembungkus: bila mode setahun DAN Z3 = Master!C4 DAN acuan terisi
    /// → pakai acuan; selain itu → rumus asli (B1/B2/B4). Mode bulanan tidak terpengaruh.
    ///
    /// Sekalian dua perbaikan rumus mode setahun:
    ///   C29     Kolektif Net Flow = X20 (blok tahunan, EAD bulan berjalan × PD Desember)
    ///           — sebelumnya selalu G20 (EAD posisi file Net Flow, yaitu Desember).
    ///   C38:C42 OS per kualitas AKTUAL bulan berjalan (U24:U28, diisi CKPN Individu)
    ///           — sebelumnya dikelompokkan dari bucket hari tunggakan (U6:U19).
    /// </summary>
    internal static class AcuanTahunan
    {
        public const string SheetKol = "B. CKPN - KOL INDV";
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        internal class Acuan
        {
            public long RunId; public int Versi; public string Tanggal, Status, KodeKC, Waktu;
            public double?[] PdNf = new double?[14];
            public double?[] PdMig = new double?[5];
            public double? Lgd;
            public Dictionary<string, double> RingkasanLgd = new Dictionary<string, double>();
            public string AnalisisJson;

            public string Label
            {
                get
                {
                    DateTime t;
                    string bln = DateTime.TryParseExact(Tanggal, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out t)
                        ? t.ToString("MMMM yyyy", new CultureInfo("id-ID")) : Tanggal;
                    return bln + " v" + Versi;
                }
            }
        }

        // ================================================================
        // Mode & periode acuan
        // ================================================================
        public static bool ModeTahunan(Excel.Workbook wb)
        {
            try
            {
                Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
                if (m == null) return false;
                string s = Convert.ToString(((Excel.Range)m.Range["D17"]).Value2) ?? "";
                return s.IndexOf("setahun", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch { return false; }
        }

        /// <summary>Tahun Desember acuan; null bila laporan Desember (Desember selalu dihitung penuh).</summary>
        public static int? TahunAcuan(DateTime tglLaporan)
        {
            return tglLaporan.Month == 12 ? (int?)null : tglLaporan.Year - 1;
        }

        // ================================================================
        // Cari acuan di database
        // ================================================================
        public static Acuan Cari(int tahunDesember, string kodeKC, out string masalah)
        {
            masalah = null;
            if (!Database.Ada) { masalah = "Database belum ada."; return null; }
            using (var con = Database.Buka(false))
            {
                long runId = 0;
                using (var cmd = Database.Cmd(con,
                    "SELECT g.id FROM run_grup g JOIN periode p ON p.id=g.periode_id " +
                    "WHERE p.tanggal LIKE @p0 AND g.kode_kc=@p1 AND g.aktif=1 AND g.dihapus=0 ORDER BY p.tanggal DESC LIMIT 1",
                    tahunDesember + "-12-%", kodeKC))
                {
                    object o = cmd.ExecuteScalar();
                    if (o != null && !(o is DBNull)) runId = Convert.ToInt64(o);
                }
                if (runId == 0)
                {
                    masalah = "Belum ada versi tersimpan grup " + kodeKC + " untuk Desember " + tahunDesember + ".";
                    return null;
                }
                return Muat(con, runId, out masalah);
            }
        }

        public static Acuan MuatRun(long runId, out string masalah)
        {
            masalah = null;
            using (var con = Database.Buka(false)) return Muat(con, runId, out masalah);
        }

        private static Acuan Muat(SQLiteConnection con, long runId, out string masalah)
        {
            masalah = null;
            var a = new Acuan { RunId = runId };
            using (var cmd = Database.Cmd(con,
                "SELECT g.versi, g.kode_kc, g.waktu, p.tanggal, p.status FROM run_grup g JOIN periode p ON p.id=g.periode_id WHERE g.id=@p0", runId))
            using (var rd = cmd.ExecuteReader())
            {
                if (!rd.Read()) { masalah = "Versi acuan tidak ditemukan."; return null; }
                a.Versi = rd.GetInt32(0); a.KodeKC = rd.GetString(1); a.Waktu = rd[2] as string;
                a.Tanggal = rd.GetString(3); a.Status = rd.GetString(4);
            }
            using (var cmd = Database.Cmd(con, "SELECT kunci, nilai FROM ringkasan WHERE run_id=@p0 AND kunci LIKE 'lgd%'", runId))
            using (var rd = cmd.ExecuteReader())
                while (rd.Read()) if (!(rd[1] is DBNull)) a.RingkasanLgd[rd.GetString(0)] = rd.GetDouble(1);

            object data = Database.Scalar(con, "SELECT data FROM analisis_pd WHERE run_id=@p0 AND jenis='pd'", runId);
            a.AnalisisJson = data as string;
            if (!string.IsNullOrEmpty(a.AnalisisJson))
            {
                try
                {
                    var d = _json.Deserialize<Dictionary<string, object>>(a.AnalisisJson);
                    Isi(a.PdNf, d.ContainsKey("pdNetFlow") ? d["pdNetFlow"] : null);
                    Isi(a.PdMig, d.ContainsKey("pdMigrasi") ? d["pdMigrasi"] : null);
                    if (d.ContainsKey("lgdWeighted") && d["lgdWeighted"] != null) a.Lgd = Convert.ToDouble(d["lgdWeighted"], CultureInfo.InvariantCulture);
                }
                catch (Exception ex) { masalah = "Data PD versi acuan tidak terbaca: " + ex.Message; return null; }
            }
            double v;
            if (!a.Lgd.HasValue && a.RingkasanLgd.TryGetValue("lgd_gabungan", out v)) a.Lgd = v;

            var kurang = new List<string>();
            for (int i = 0; i < 13; i++) if (!a.PdNf[i].HasValue) { kurang.Add("PD Net Flow"); break; }
            for (int i = 0; i < 5; i++) if (!a.PdMig[i].HasValue) { kurang.Add("PD Migration"); break; }
            if (!a.Lgd.HasValue) kurang.Add("LGD weighted");
            if (kurang.Count > 0)
            {
                masalah = "Versi " + a.Label + " tidak menyimpan " + string.Join(", ", kurang.ToArray()) +
                          " (disimpan sebelum Tahap 4). Simpan ulang grup Desember tersebut sekali.";
                return null;
            }
            if (!a.PdNf[13].HasValue) a.PdNf[13] = 1;   // bucket > 360 hari: PD 100% (sel E19/V19 sheet)
            return a;
        }

        private static void Isi(double?[] tujuan, object sumber)
        {
            var arr = sumber as System.Collections.IList;
            if (arr == null) return;
            for (int i = 0; i < tujuan.Length && i < arr.Count; i++)
                tujuan[i] = arr[i] == null ? (double?)null : Convert.ToDouble(arr[i], CultureInfo.InvariantCulture);
        }

        // ================================================================
        // Untuk panel: status mode & ketersediaan acuan per grup
        // ================================================================
        public static Dictionary<string, object> Info(Excel.Application app)
        {
            var hasil = new Dictionary<string, object> { { "tahunan", false } };
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) return hasil;
            DateTime tgl;
            if (!ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return hasil;
            bool tahunan = ModeTahunan(wb);
            int? th = TahunAcuan(tgl);
            hasil["tahunan"] = tahunan;
            hasil["tanggal"] = tgl.ToString("yyyy-MM-dd");
            hasil["desember"] = tgl.Month == 12;
            hasil["tahunAcuan"] = th;
            if (!tahunan || !th.HasValue || !Database.Ada) return hasil;

            var grup = new List<object>();
            List<Periode.Grup> susunan = null;
            using (var con = Database.Buka(false))
            {
                int sumber;
                susunan = Periode.Susunan(con, tgl.Year, out sumber);
            }
            foreach (var g in susunan ?? new List<Periode.Grup>())
            {
                string masalah;
                Acuan a = null;
                try { a = Cari(th.Value, g.KodeKC, out masalah); }
                catch (Exception ex) { masalah = ex.Message; }
                grup.Add(new Dictionary<string, object>
                {
                    { "nama", g.Nama }, { "kodeKC", g.KodeKC },
                    { "acuan", a == null ? null : new Dictionary<string, object>
                        {
                            { "runId", a.RunId }, { "versi", a.Versi }, { "label", a.Label }, { "status", a.Status },
                            { "lgd", a.Lgd }, { "waktu", a.Waktu }
                        } },
                    { "masalah", masalah }
                });
            }
            hasil["grup"] = grup;
            return hasil;
        }

        // ================================================================
        // Tulis ke sheet
        // ================================================================
        /// <summary>Isi acuan untuk periode Master!C4 (dipanggil pipeline setelah CKPN Individu).</summary>
        public static void Tulis(Excel.Workbook wb, Acuan a)
        {
            Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
            object c4 = ((Excel.Range)m.Range["C4"]).Value2;
            if (!(c4 is double)) throw new InvalidOperationException("Master!C4 bukan tanggal.");

            Protection.UbahSheet(wb, SheetKol, ws =>
            {
                PasangRumus(ws);
                Set(ws, "Z3", (double)c4, "dd-mmm-yyyy");
                Set(ws, "AA3", "Acuan " + a.Label, null);
                Set(ws, "AB3", (double)a.RunId, "0");
                for (int i = 0; i < 14; i++) Set(ws, "Z" + (6 + i), a.PdNf[i].Value, "0.00%");
                Set(ws, "AA6", a.Lgd.Value, "0.00%");
                for (int i = 0; i < 5; i++) Set(ws, "Z" + (38 + i), a.PdMig[i].Value, "0.00%");
            });
        }

        /// <summary>Hapus acuan (dipanggil saat PD &amp; LGD dihitung penuh) dan pastikan rumus terpasang.</summary>
        public static void Kosongkan(Excel.Workbook wb)
        {
            Protection.UbahSheet(wb, SheetKol, ws =>
            {
                PasangRumus(ws);
                foreach (var alamat in new[] { "Z3", "AA3", "AB3", "Z6:Z19", "AA6", "Z38:Z42" })
                    ((Excel.Range)ws.Range[alamat]).ClearContents();
            });
        }

        /// <summary>
        /// Acuan yang sedang berlaku di workbook (untuk Simpan grup): mode setahun, Z3 = Master!C4,
        /// dan AB3 berisi id versi Desember. Null bila tidak ada.
        /// </summary>
        public static Acuan AktifDiWorkbook(Excel.Workbook wb, out string masalah)
        {
            masalah = null;
            try
            {
                if (!ModeTahunan(wb)) return null;
                Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetKol);
                Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
                if (ws == null || m == null) return null;
                object z3 = ((Excel.Range)ws.Range["Z3"]).Value2;
                object c4 = ((Excel.Range)m.Range["C4"]).Value2;
                object ab3 = ((Excel.Range)ws.Range["AB3"]).Value2;
                if (!(z3 is double) || !(c4 is double) || !(ab3 is double)) return null;
                if (Math.Abs((double)z3 - (double)c4) > 0.5) return null;
                return MuatRun(Convert.ToInt64((double)ab3), out masalah);
            }
            catch (Exception ex) { masalah = ex.Message; return null; }
        }

        private static void Set(Excel.Worksheet ws, string alamat, object nilai, string format)
        {
            var r = (Excel.Range)ws.Range[alamat];
            if (format != null) r.NumberFormat = format;
            r.Value2 = nilai;
        }

        // ================================================================
        // Rumus (idempoten) — pembungkus acuan di atas rumus asli template
        // ================================================================
        private const string Setahun = "ISNUMBER(SEARCH(\"setahun\",Master!$D$17))";
        private const string AcuanBerlaku = Setahun + ",$Z$3=Master!$C$4";
        private const string LgdB4 = "INDEX('B4.LGD-CS MACET'!$E:$E,MATCH(\"LGD Weighted\",'B4.LGD-CS MACET'!$B:$B,0))";

        public static void PasangRumus(Excel.Worksheet ws)
        {
            // Label area acuan
            ((Excel.Range)ws.Range["Z2"]).Value2 = "ACUAN PD & LGD (diisi Panel CKPN — mode setahun sekali)";
            ((Excel.Range)ws.Range["Z5"]).Value2 = "PD Net Flow acuan";
            ((Excel.Range)ws.Range["AA5"]).Value2 = "LGD acuan";
            ((Excel.Range)ws.Range["Z37"]).Value2 = "PD Migration acuan";

            // Blok tahunan: PD (V) & LGD (W)
            for (int r = 6; r <= 19; r++)
            {
                string asli = r == 19 ? "1"
                    : "INDEX('B1.PD-Net Flow'!$P:$P,MATCH(\"PD\",'B1.PD-Net Flow'!$P:$P,0)+B" + r + ")";
                ((Excel.Range)ws.Range["V" + r]).Formula = "=IF(AND(" + AcuanBerlaku + ",ISNUMBER($Z" + r + ")),$Z" + r + "," + asli + ")";
                ((Excel.Range)ws.Range["W" + r]).Formula = "=IF(AND(" + AcuanBerlaku + ",ISNUMBER($AA$6)),$AA$6," + LgdB4 + ")";
            }

            // Kolektif Net Flow (Summary B7 membaca C29): mode setahun → blok tahunan X20
            ((Excel.Range)ws.Range["C29"]).Formula = "=IF(" + Setahun + ",X20,G20)";

            // OS per kualitas aktual (diisi CKPN Individu pada mode setahun)
            ((Excel.Range)ws.Range["S22"]).Value2 = "OS per kualitas (aktual) — Top-N dikecualikan";
            ((Excel.Range)ws.Range["S23"]).Value2 = "Kualitas";
            ((Excel.Range)ws.Range["U23"]).Value2 = "Baki Debet";
            for (int k = 1; k <= 5; k++) ((Excel.Range)ws.Range["S" + (23 + k)]).Value2 = k;

            string[] bucket = { "U6", "SUM(U7:U9)", "SUM(U10:U12)", "SUM(U13:U15)", "SUM(U16:U19)" };
            string[] bulanan = { "D6", "SUM(D7:D9)", "SUM(D10:D12)", "SUM(D13:D15)", "SUM(D16:D19)" };
            for (int k = 0; k < 5; k++)
            {
                int r = 38 + k;
                // C: EAD per kualitas — mode setahun pakai OS kualitas aktual (fallback bucket bila belum diisi)
                ((Excel.Range)ws.Range["C" + r]).Formula =
                    "=IF(" + Setahun + ",IF(COUNT($U$24:$U$28)>0,$U$" + (24 + k) + "," + bucket[k] + ")," + bulanan[k] + ")";
                // D: PD Migration, E: LGD
                ((Excel.Range)ws.Range["D" + r]).Formula =
                    "=IF(AND(" + AcuanBerlaku + ",ISNUMBER($Z" + r + ")),$Z" + r +
                    ",INDEX('B2.PD-Migration'!$B:$B,MATCH(\"weighted average\",'B2.PD-Migration'!$B:$B,0)+1+" + (k + 1) + "))";
                ((Excel.Range)ws.Range["E" + r]).Formula = "=IF(AND(" + AcuanBerlaku + ",ISNUMBER($AA$6)),$AA$6," + LgdB4 + ")";
            }
        }
    }
}
