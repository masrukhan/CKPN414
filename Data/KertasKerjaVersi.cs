using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using CKPNLibrary.Helpers;
using CKPNLibrary.Panel;
using G = CKPNLibrary.Data.XlsxTulis.Gaya;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// Kertas kerja Excel satu versi grup, dibangun dari DATABASE (bukan dari sheet aplikasi).
    ///
    /// Latar belakang: snapshot .xlsx hanya dibuat saat grup disimpan dari sheet. Versi yang lahir dari koreksi
    /// panel / "Terapkan penyesuaian ke versi tersimpan" dihitung di database sehingga tidak punya snapshot —
    /// user tidak bisa mendapatkan Excel yang angkanya sama dengan versi itu. Kelas ini menyusun workbook dengan
    /// sheet seperti aplikasi (tanpa Master):
    ///   Info Versi · Summary · A. CKPN - INDV · B. CKPN - KOL INDV · B1.PD-Net Flow · B2.PD-Migration · B3.LGD-ER · B4.LGD-CS MACET
    ///
    /// RUMUS HIDUP: rumus aplikasi direkonstruksi (roll rate, flow to loss, PD, matriks migrasi, LGD ER/CS,
    /// EAD × PD × LGD, Summary, ABA, PPKA vs CKPN) sehingga perubahan satu angka ikut mengalir ke total.
    /// Setiap rantai diperiksa dulu terhadap angka tersimpan: bila hasil rumus tidak sama persis dengan yang dipakai
    /// versi (mis. acuan PD/LGD Desember, versi lama tanpa rincian), sel diisi angka tersimpan berlatar kuning dan
    /// alasannya dicatat di sheet Info Versi. Dengan begitu angka file SELALU sama dengan database.
    ///
    /// Pemakaian:
    ///   • BuatSnapshot  — otomatis setelah versi baru lahir dari koreksi (KoreksiRun.Simpan) → library\staging\yyyy-MM\
    ///   • Ekspor        — tombol "Kertas kerja" (Analisis › Versi, Dokumen per grup) → library\dokumen\yyyy-MM\
    ///   • HapusSnapshot — saat versi dihapus (Periode.HapusVersi)
    /// File ditulis langsung (zip + XML) tanpa membuka Excel — ringan, aman dipanggil di luar thread Excel.
    /// </summary>
    internal static class KertasKerjaVersi
    {
        public const string ShInfo = "Info Versi", ShSum = "Summary", ShA = "A. CKPN - INDV", ShB = "B. CKPN - KOL INDV",
                            ShB1 = "B1.PD-Net Flow", ShB2 = "B2.PD-Migration", ShB3 = "B3.LGD-ER", ShB4 = "B4.LGD-CS MACET";
        private const double TolRp = 1.0, TolRate = 1e-7;

        private static readonly string[] LabelBucketB1 =
        {
            "0 hari sd x days", "x days - 30 hari", "31 - 60 hari", "61 - 90 hari", "91 -120 hari", "121 -150 hari", "151 - 180 hari",
            "181-210 hari", "211-240 hari", "241-270 hari", "271-300 hari", "301-330 hari", "331-360 hari"
        };
        private static readonly string[] LabelBucketKol =
        {
            "0 hari sd x days", "x days - 30 hari", "31 - 60 hari", "61 - 90 hari", "91 -120 hari", "121 -150 hari", "151 - 180 hari",
            "181-210 hari", "211-240 hari", "241-270 hari", "271-300 hari", "301 -330 hari", "331-360 hari", ">360 hari"
        };
        private static readonly string[] NamaKualitas = { "Lancar", "Dalam Perhatian Khusus", "Kurang Lancar", "Diragukan", "Macet" };

        // ================================================================
        // Konteks pembangunan: data versi + posisi sel penting untuk rumus antarsheet
        // ================================================================
        private sealed class Ktx
        {
            public Dictionary<string, object> Data, Run, Grup, Rg;
            public string Tanggal, KodeKC, NamaGrup;
            public int Versi;
            public readonly List<string> Catatan = new List<string>();

            // A. Individu
            public int ATotal;
            public double AOs, AJam, ACkpn;

            // B1 Net Flow
            public bool B1Rumus;
            public double?[] B1Ead = new double?[14];   // EAD bulan terakhir per bucket (P4..P16, P18)
            public double?[] B1Pd = new double?[13];    // PD per bucket (P40..P52)

            // B2 Migration
            public bool B2Ada;
            public double?[] B2Pd = new double?[5];      // weighted average (B82..B86)

            // B3/B4 LGD
            public int B4LgdBaris;                       // baris "LGD Weighted" di B4 (kolom C/D/E)
            public double? LgdCs, LgdEr, LgdGab;
            public double ErWo, ErRec;                    // Total WO & Recovery ER di ringkasan B3

            // B. Kolektif
            public double KolNf, KolMig, KolInd;
            public bool KolAda;

            // Summary
            public double SumB6, SumB7, SumB13, SumB14, SumC26, SumH9;
        }

        // ================================================================
        // API
        // ================================================================

        /// <summary>Snapshot versi (dibuat dari database) → library\staging\yyyy-MM\, path dicatat di run_grup.snapshot.</summary>
        public static string BuatSnapshot(long runId)
        {
            var data = PerhitunganGrup.Data(runId);
            var run = (Dictionary<string, object>)data["run"];
            string tanggal = S(run["tanggal"]);
            string folder = Path.Combine(AppPaths.FolderStaging, tanggal.Substring(0, 7));
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "CKPN_" + tanggal.Replace("-", "") + "_" + Suffix(S(run["kodeKC"])) + "_v" + Convert.ToInt32(run["versi"]) + "_db.xlsx");
            Tulis(data, path);
            using (var con = Database.Buka(true))
                Database.Exec(con, "UPDATE run_grup SET snapshot=@p0 WHERE id=@p1", path, runId);
            return path;
        }

        /// <summary>
        /// Ekspor kertas kerja satu atau beberapa versi → library\dokumen\yyyy-MM\ (satu file per grup).
        /// Satu file → langsung dibuka; beberapa → folder dibuka.
        /// </summary>
        public static Dictionary<string, object> Ekspor(IList runIds, bool buka)
        {
            if (runIds == null || runIds.Count == 0) throw new InvalidOperationException("Pilih versi grup yang akan diekspor.");
            var file = new List<object>();
            string folder = null, periode = "";
            var gagal = new List<string>();
            foreach (var o in runIds)
            {
                long runId = Convert.ToInt64(o);
                try
                {
                    var data = PerhitunganGrup.Data(runId);
                    var run = (Dictionary<string, object>)data["run"];
                    var grup = (Dictionary<string, object>)data["grup"];
                    periode = S(run["tanggal"]);
                    folder = Path.Combine(PanelBridge.FolderDokumen, periode.Substring(0, 7));
                    Directory.CreateDirectory(folder);
                    string dasar = periode.Substring(0, 7) + " Kertas kerja " + S(grup["nama"]) + " v" + Convert.ToInt32(run["versi"]);
                    string path = Path.Combine(folder, NamaAman(dasar) + ".xlsx");
                    try { Tulis(data, path); }
                    catch (IOException)
                    {
                        // file lama sedang dibuka di Excel → nama baru
                        path = Path.Combine(folder, NamaAman(dasar) + "_" + DateTime.Now.ToString("HHmmss") + ".xlsx");
                        Tulis(data, path);
                    }
                    file.Add(new Dictionary<string, object> { { "file", path }, { "nama", Path.GetFileName(path) }, { "grup", S(grup["nama"]) } });
                }
                catch (Exception ex)
                {
                    gagal.Add("versi #" + runId + ": " + ex.Message);
                    CatatanLog.Tulis("Kertas kerja versi #" + runId + " gagal: " + ex.Message);
                }
            }
            if (file.Count == 0) throw new InvalidOperationException("Kertas kerja gagal dibuat. " + string.Join("; ", gagal.ToArray()));
            if (buka)
            {
                try
                {
                    if (file.Count == 1) Process.Start(S(((Dictionary<string, object>)file[0])["file"]));
                    else Process.Start(folder);
                }
                catch (Exception ex) { CatatanLog.Tulis("Buka kertas kerja: " + ex.Message); }
            }
            LogProses.CatatPanel(periode, "Dokumen", "Kertas kerja Excel (" + file.Count + " grup)", gagal.Count == 0 ? LogProses.OK : LogProses.Peringatan,
                LogProses.R().Tambah("Folder", folder).TambahBila(gagal.Count > 0, "Gagal", string.Join("; ", gagal.ToArray())));
            return new Dictionary<string, object> { { "file", file }, { "folder", folder }, { "gagal", gagal } };
        }

        /// <summary>
        /// Buka snapshot versi; versi hasil koreksi tanpa snapshot (dibuat sebelum fitur ini) dibuatkan dari database.
        /// User tanpa hak tulis: file dibuat di folder sementara, database tidak diubah.
        /// </summary>
        public static void BukaAtauBuat(long runId)
        {
            string info;
            if (Database.BolehMenulis(out info)) { Process.Start(BuatSnapshot(runId)); return; }
            var data = PerhitunganGrup.Data(runId);
            var run = (Dictionary<string, object>)data["run"];
            string folder = Path.Combine(Path.GetTempPath(), "CKPN414");
            Directory.CreateDirectory(folder);
            string path = Path.Combine(folder, "CKPN_" + S(run["tanggal"]).Replace("-", "") + "_" + Suffix(S(run["kodeKC"])) + "_v" + Convert.ToInt32(run["versi"]) + "_db.xlsx");
            Tulis(data, path);
            Process.Start(path);
        }

        /// <summary>Hapus file snapshot versi (hanya di dalam folder staging). Gagal hapus (file sedang dibuka) → dicatat saja.</summary>
        public static bool HapusSnapshot(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            try
            {
                if (!Path.GetFullPath(path).StartsWith(Path.GetFullPath(AppPaths.FolderStaging), StringComparison.OrdinalIgnoreCase)) return false;
                if (!File.Exists(path)) return false;
                File.Delete(path);
                return true;
            }
            catch (Exception ex)
            {
                CatatanLog.Tulis("Snapshot " + Path.GetFileName(path) + " tidak dapat dihapus: " + ex.Message);
                return false;
            }
        }

        private static string Suffix(string kode)
        {
            var mode = ModeHitung.Aktif;
            if (!mode.Utama) return mode.Id + "_" + kode.Replace(",", "-");
            var parts = kode.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return "Export";
            if (parts.Length >= 6) return "all_segmen";
            var sb = new StringBuilder();
            foreach (var p in parts) sb.Append(p.Trim().Replace(" ", ""));
            return sb.ToString();
        }

        private static string NamaAman(string s)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '-');
            return s.Trim();
        }

        // ================================================================
        // Penyusun workbook
        // ================================================================
        private static void Tulis(Dictionary<string, object> data, string path)
        {
            var k = new Ktx { Data = data };
            k.Run = (Dictionary<string, object>)data["run"];
            k.Grup = (Dictionary<string, object>)data["grup"];
            k.Rg = data["ringkasan"] as Dictionary<string, object> ?? new Dictionary<string, object>();
            k.Tanggal = S(k.Run["tanggal"]);
            k.KodeKC = S(k.Run["kodeKC"]);
            k.NamaGrup = S(k.Grup["nama"]);
            k.Versi = Convert.ToInt32(k.Run["versi"]);
            if (data.ContainsKey("acuanTanggal") && data["acuanTanggal"] != null)
                k.Catatan.Add("PD & LGD versi ini memakai acuan Desember (periode " + S(data["acuanTanggal"]) + ", mode setahun sekali). " +
                              "Sheet B1–B4 berisi rincian acuan; sel kuning di B. CKPN - KOL INDV adalah angka acuan yang dipakai.");

            var x = new XlsxTulis();
            // Urutan sheet seperti aplikasi; Master diganti Info Versi
            var info = x.Tambah(ShInfo); info.WarnaTab("FF2E7D32");
            var sum = x.Tambah(ShSum);
            var a = x.Tambah(ShA);
            var kol = x.Tambah(ShB);
            var b1 = x.Tambah(ShB1);
            var b2 = x.Tambah(ShB2);
            var b3 = x.Tambah(ShB3);
            var b4 = x.Tambah(ShB4);

            // Isi dari hulu ke hilir: rincian dulu (hasil hitungnya dipakai untuk memeriksa rantai rumus)
            SheetB1(b1, k);
            int b3Wo = SheetB3(b3, k);
            SheetB2(b2, k);
            SheetB4(b4, k, b3Wo);
            SheetA(a, k);
            SheetKol(kol, k);
            SheetSummary(sum, k);
            SheetInfo(info, k);
            x.Simpan(path);
        }

        // ================================================================
        // B1. PD Net Flow — saldo bucket 13 bulan, roll rate, flow to loss, PD (rumus sama dengan template)
        // ================================================================
        private static void SheetB1(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Teks(1, 2, "TABEL PD NET FLOW — OUTSTANDING POOL PER BUCKET", G.Judul);
            ws.Lebar(1, 2); ws.Lebar(2, 10); ws.Lebar(3, 18);
            for (int c = 4; c <= 27; c++) ws.Lebar(c, 14);

            var pd = k.Data["pd"] as Dictionary<string, object>;
            var nf = pd == null ? null : D(Ambil(pd, "netflow"));
            var bulan = nf == null ? null : L(Ambil(nf, "bulan"));
            var saldoL = nf == null ? null : L(Ambil(nf, "saldo"));
            var woL = nf == null ? null : L(Ambil(nf, "wo"));
            if (bulan == null || saldoL == null || saldoL.Count < 14)
            {
                ws.Teks(3, 2, "Rincian PD Net Flow tidak tersimpan pada versi ini (versi lama). Gunakan Dokumen per grup › Lengkapi dari sheet bila tersedia.", G.Peringatan);
                k.Catatan.Add("B1: saldo bucket Net Flow tidak tersimpan — EAD & PD Net Flow di B. CKPN - KOL INDV memakai angka tersimpan.");
                return;
            }
            int nb = bulan.Count;
            ws.Teks(2, 2, "Periode Data:", G.Tebal);
            ws.Teks(2, 4, S(bulan[0]) + "  s.d.  " + S(bulan[nb - 1]));
            ws.Teks(2, 9, "Sumber:", G.Tebal);
            ws.Teks(2, 10, "Versi tersimpan v" + k.Versi + " (database)");

            // ---- saldo ----
            ws.Teks(3, 2, "No. Bucket", G.Kepala); ws.Teks(3, 3, "Hari tunggakan", G.Kepala);
            var s = new double?[20, nb];     // [baris sheet, bulan] baris 4..19
            for (int m = 0; m < nb; m++) ws.Teks(3, 4 + m, S(bulan[m]), G.Kepala);
            for (int b = 0; b < 13; b++)
            {
                int r = 4 + b;
                ws.Angka(r, 2, b + 1, G.Biasa); ws.Teks(r, 3, LabelBucketB1[b]);
                var baris = L(saldoL[b]);
                for (int m = 0; m < nb; m++)
                {
                    double? v = baris != null && m < baris.Count ? N(baris[m]) : null;
                    s[r, m] = v;
                    ws.Angka(r, 4 + m, v, G.Input);
                }
            }
            ws.Angka(17, 2, 14, G.Biasa); ws.Teks(17, 3, ">360+WO", G.Tebal);
            ws.Teks(18, 3, "> 360 hari"); ws.Teks(19, 3, "WO");
            var b360 = L(saldoL[13]);
            for (int m = 0; m < nb; m++)
            {
                double? v18 = b360 != null && m < b360.Count ? N(b360[m]) : null;
                double? v19 = woL != null && m < woL.Count ? N(woL[m]) : null;
                s[18, m] = v18; s[19, m] = v19;
                ws.Angka(18, 4 + m, v18, G.Input);
                ws.Angka(19, 4 + m, v19, G.Input);
                s[17, m] = (v18 ?? 0) + (v19 ?? 0);
                string c = XlsxTulis.Kol(4 + m);
                ws.Rumus(17, 4 + m, "=" + c + "18+" + c + "19", s[17, m], G.AngkaTebal);
            }
            // EAD bulan terakhir (dipakai B. CKPN - KOL INDV kolom D)
            for (int b = 0; b < 13; b++) k.B1Ead[b] = s[4 + b, nb - 1];
            k.B1Ead[13] = s[18, nb - 1];
            k.B1Rumus = nb == 13;
            ws.Beku(3, 3);

            if (nb != 13)
            {
                // Bentuk lama (bukan 13 posisi bulan): roll rate/PD tersimpan ditulis sebagai angka
                k.Catatan.Add("B1: data Net Flow berisi " + nb + " posisi bulan (bukan 13) — roll rate & PD ditulis sebagai angka tersimpan.");
                TulisPdTersimpan(ws, k, nf);
                return;
            }

            // ---- roll rate (baris 24..36), kolom D..O = 12 perpindahan, P..AA = perpanjangan rata-rata 6 bulan ----
            ws.Teks(23, 2, "Perpindahan Bucket", G.Kepala); ws.Teks(23, 3, "", G.Kepala);
            for (int m = 0; m < 12; m++)
                ws.Rumus(23, 4 + m, "=" + XlsxTulis.Ref(3, 5 + m), S(bulan[m + 1]), G.Kepala);
            for (int e = 0; e < 12; e++) ws.Teks(23, 16 + e, "rata-rata " + (e + 1), G.Kepala);
            var rr = new double[13, 24];
            for (int b = 0; b < 13; b++)
            {
                int r = 24 + b, r1 = 4 + b, r2 = 5 + b;
                ws.Teks(r, 2, (b + 1) + " to " + (b + 2), G.Tebal);
                for (int m = 0; m < 12; m++)
                {
                    string X = XlsxTulis.Kol(4 + m), Y = XlsxTulis.Kol(5 + m);
                    double? xv = s[r1, m], yv = s[r2, m + 1];
                    double v;
                    if (!xv.HasValue && r1 != 17) v = 1;                       // ISBLANK → 1
                    else
                    {
                        double q = (xv ?? 0) == 0 ? 0 : (yv ?? 0) / xv.Value;    // IFERROR(Y/X,0)
                        v = q > 1 ? 1 : q;
                    }
                    rr[b, m] = v;
                    ws.Rumus(r, 4 + m,
                        "=IF(ISBLANK(" + X + r1 + "),1,IF(IFERROR(" + Y + r2 + "/" + X + r1 + ",0)>1,1,IFERROR(" + Y + r2 + "/" + X + r1 + ",0)))",
                        v, G.Pd);
                }
                for (int e = 12; e < 24; e++)
                {
                    int c = 4 + e;
                    double jml = 0; int n = 0;
                    for (int j = e - 6; j < e; j++) if (rr[b, j] > 0) { jml += rr[b, j]; n++; }
                    rr[b, e] = n == 0 ? 0 : jml / n;
                    ws.Rumus(r, c, "=IFERROR(AVERAGEIF(" + XlsxTulis.Kol(c - 6) + r + ":" + XlsxTulis.Kol(c - 1) + r + ",\">0\"),0)", rr[b, e], G.Pd);
                }
            }

            // ---- compound flow to loss (baris 40..52) + PD (kolom P) ----
            ws.Teks(38, 2, "Compound Flow to Loss", G.Subjudul);
            ws.Teks(39, 2, "Bucket", G.Kepala); ws.Teks(39, 3, "", G.Kepala);
            for (int m = 0; m < 12; m++) ws.Rumus(39, 4 + m, "=" + XlsxTulis.Ref(23, 4 + m), S(bulan[m + 1]), G.Kepala);
            ws.Teks(39, 16, "PD", G.Kepala);
            for (int b = 0; b < 13; b++)
            {
                int r = 40 + b;
                ws.Teks(r, 2, (b + 1) + " to " + (b + 2), G.Tebal);
                double jumlah = 0;
                for (int m = 0; m < 12; m++)
                {
                    var f = new StringBuilder("=IFERROR(");
                    double v = 1;
                    for (int t = 0; t <= 12 - b; t++)
                    {
                        if (t > 0) f.Append("*");
                        f.Append(XlsxTulis.Kol(4 + m + t)).Append(24 + b + t);
                        v *= rr[b + t, m + t];
                    }
                    f.Append(",0)");
                    jumlah += v;
                    ws.Rumus(r, 4 + m, f.ToString(), v, G.Pd);
                }
                double pdB = Math.Min(1, jumlah / 12);
                k.B1Pd[b] = pdB;
                ws.Rumus(r, 16, "=IFERROR(MIN(100%,AVERAGE(D" + r + ":O" + r + ")),0)", pdB, G.Pd);
            }
            ws.Teks(54, 2, "Rumus sama dengan sheet aplikasi: roll rate = saldo bucket berikut bulan berikut ÷ saldo bucket (maks. 100%); " +
                           "flow to loss = perkalian roll rate berantai; PD = rata-rata 12 flow to loss (maks. 100%).", G.Catatan);
        }

        private static void TulisPdTersimpan(XlsxTulis.Lembar ws, Ktx k, Dictionary<string, object> nf)
        {
            var pdF = L(Ambil(nf, "pdFlow"));
            ws.Teks(39, 2, "Bucket", G.Kepala); ws.Teks(39, 16, "PD", G.Kepala);
            for (int b = 0; b < 13; b++)
            {
                ws.Teks(40 + b, 2, (b + 1) + " to " + (b + 2), G.Tebal);
                double? v = pdF != null && b < pdF.Count ? N(pdF[b]) : null;
                ws.Angka(40 + b, 16, v, G.PersenNilai);
                k.B1Pd[b] = v;
            }
        }

        // ================================================================
        // B2. PD Migration — 4 matriks triwulan + rata-rata tertimbang (rumus sama dengan template)
        // ================================================================
        private static void SheetB2(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Lebar(1, 22); ws.Lebar(2, 16);
            for (int c = 3; c <= 10; c++) ws.Lebar(c, 14);
            var pd = k.Data["pd"] as Dictionary<string, object>;
            var mig = pd == null ? null : D(Ambil(pd, "migrasi"));
            var tw = mig == null ? null : L(Ambil(mig, "triwulan"));
            if (tw == null || tw.Count < 4)
            {
                ws.Teks(1, 1, "Rincian PD Migration tidak tersimpan pada versi ini.", G.Peringatan);
                k.Catatan.Add("B2: matriks PD Migration tidak tersimpan — PD Migration di B. CKPN - KOL INDV memakai angka tersimpan.");
                return;
            }
            k.B2Ada = true;
            var saldoAwal = new double[4, 5];
            var J = new double[4, 5];
            for (int t = 0; t < 4; t++)
            {
                var d = D(tw[t]);
                int bs = 4 + 19 * t;
                ws.Teks(bs - 3, 1, "Posisi " + S(Ambil(d, "nama")), G.Subjudul);
                ws.Teks(bs - 2, 1, "Kualitas", G.Kepala); ws.Teks(bs - 2, 2, "Saldo awal", G.Kepala);
                ws.Teks(bs - 2, 3, "Posisi akhir", G.Kepala); ws.GabungSel(bs - 2, 3, bs - 2, 9);
                ws.Teks(bs - 2, 10, "TOTAL", G.Kepala);
                for (int j = 0; j < 5; j++) ws.Angka(bs - 1, 3 + j, j + 1, G.Kepala);
                ws.Teks(bs - 1, 8, "Hapus Buku", G.Kepala); ws.Teks(bs - 1, 9, "Pembayaran", G.Kepala);
                var sa = L(Ambil(d, "saldoAwal"));
                var mx = L(Ambil(d, "matriks"));
                var m = new double[5, 7];
                double totSa = 0, totJ = 0;
                for (int q = 0; q < 5; q++)
                {
                    int r = bs + q;
                    ws.Angka(r, 1, q + 1, G.Biasa);
                    saldoAwal[t, q] = N(sa != null && q < sa.Count ? sa[q] : null) ?? 0;
                    ws.Angka(r, 2, saldoAwal[t, q], G.Input);
                    var baris = mx != null && q < mx.Count ? L(mx[q]) : null;
                    double jml = 0;
                    for (int j = 0; j < 7; j++)
                    {
                        m[q, j] = N(baris != null && j < baris.Count ? baris[j] : null) ?? 0;
                        jml += m[q, j];
                        ws.Angka(r, 3 + j, m[q, j], G.Input);
                    }
                    ws.Rumus(r, 10, "=SUM(C" + r + ":I" + r + ")", jml, G.Angka);
                    totSa += saldoAwal[t, q]; totJ += jml;
                }
                ws.Rumus(bs + 5, 2, "=SUM(B" + bs + ":B" + (bs + 4) + ")", totSa, G.Total);
                ws.Rumus(bs + 5, 10, "=SUM(J" + bs + ":J" + (bs + 4) + ")", totJ, G.Total);

                // ---- rate & PD per kualitas ----
                ws.Teks(bs + 7, 1, "Kualitas", G.Kepala); ws.Teks(bs + 7, 2, "Saldo awal", G.Kepala);
                ws.Teks(bs + 7, 3, "Rate perpindahan", G.Kepala); ws.GabungSel(bs + 7, 3, bs + 7, 9);
                ws.Teks(bs + 7, 10, "PD", G.Kepala);
                for (int j = 0; j < 5; j++) ws.Angka(bs + 8, 3 + j, j + 1, G.Kepala);
                ws.Teks(bs + 8, 8, "Hapus Buku", G.Kepala); ws.Teks(bs + 8, 9, "Pembayaran", G.Kepala);
                var rate = new double[5, 7];
                for (int q = 0; q < 5; q++)
                {
                    int r = bs + 9 + q;
                    ws.Angka(r, 1, q + 1, G.Biasa);
                    ws.Rumus(r, 2, "=B" + (bs + q), saldoAwal[t, q], G.Angka);
                    for (int j = 0; j < 7; j++)
                    {
                        string c = XlsxTulis.Kol(3 + j);
                        rate[q, j] = saldoAwal[t, q] == 0 ? 0 : m[q, j] / saldoAwal[t, q];
                        ws.Rumus(r, 3 + j, "=IFERROR(" + c + (bs + q) + "/$B" + (bs + q) + ",0)", rate[q, j], G.Pd);
                    }
                }
                // J berantai dari bawah: J5 = H+G; Jk = H + Σ (kolom kualitas lebih buruk × J kualitas itu)
                for (int q = 4; q >= 0; q--)
                {
                    int r = bs + 9 + q;
                    string f;
                    double v = rate[q, 5];
                    if (q == 4) { f = "=H" + r + "+G" + r; v += rate[q, 4]; }
                    else
                    {
                        var sb = new StringBuilder("=H" + r);
                        for (int i = 4; i >= q + 1; i--)
                        {
                            sb.Append("+(" + XlsxTulis.Kol(3 + i) + r + "*J" + (bs + 9 + i) + ")");
                            v += rate[q, i] * J[t, i];
                        }
                        f = sb.ToString();
                    }
                    J[t, q] = v;
                    ws.Rumus(r, 10, f, v, G.Pd);
                }
                ws.Rumus(bs + 14, 2, "=SUM(B" + (bs + 9) + ":B" + (bs + 13) + ")", totSa, G.Total);
            }

            // ---- rata-rata tertimbang 4 triwulan (dipakai sebagai PD Migration) ----
            ws.Teks(78, 1, "Rata-rata PD berdasarkan PD Triwulanan", G.Subjudul);
            ws.Teks(80, 1, "Kualitas", G.Kepala); ws.Teks(80, 2, "weighted average", G.Kepala); ws.Teks(80, 3, "Average", G.Kepala);
            for (int q = 0; q < 5; q++)
            {
                int r = 82 + q;
                ws.Angka(r, 1, q + 1, G.Biasa);
                var bag = new string[4]; var atas = new string[4]; var rata = new string[4];
                double num = 0, den = 0, avg = 0;
                for (int t = 0; t < 4; t++)
                {
                    int rr = 13 + 19 * t + q;
                    atas[t] = "(B" + rr + "*J" + rr + ")"; bag[t] = "B" + rr; rata[t] = "J" + rr;
                    num += saldoAwal[t, q] * J[t, q]; den += saldoAwal[t, q]; avg += J[t, q];
                }
                double w = den == 0 ? 0 : num / den;
                k.B2Pd[q] = w;
                ws.Rumus(r, 2, "=IFERROR((" + string.Join("+", atas) + ")/(" + string.Join("+", bag) + "),0)", w, G.Pd);
                ws.Rumus(r, 3, "=(" + string.Join("+", rata) + ")/4", avg / 4, G.Pd);
            }
            ws.Teks(88, 1, "PD Migration yang dipakai = weighted average (kolom B).", G.Catatan);
        }

        // ================================================================
        // B3. LGD Expected Recoveries — blok per kohort tahun hapus buku
        // ================================================================
        /// <returns>baris "Total WO" ringkasan (kolom D); Total Recovery = baris berikutnya</returns>
        private static int SheetB3(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Teks(3, 2, "B3. LGD — Expected Recoveries", G.Judul);
            ws.Lebar(2, 2); ws.Lebar(3, 26); ws.Lebar(4, 14); ws.Lebar(5, 16);
            for (int c = 6; c <= 11; c++) ws.Lebar(c, 14);
            ws.Lebar(14, 16); ws.Lebar(15, 12);

            double? woRg = N(Ambil(k.Rg, "lgd_er_wo")), recRg = N(Ambil(k.Rg, "lgd_er_rec"));
            var er = k.Data["lgder"] as Dictionary<string, object>;
            var kohort = er == null ? null : L(Ambil(er, "kohort"));
            int r = 5;
            var subE = new List<string>(); var subN = new List<string>();
            double totE = 0, totN = 0;
            bool rinci = kohort != null && kohort.Count > 0;
            if (rinci)
            {
                foreach (var ko in kohort)
                {
                    var kd = D(ko);
                    int? th = Bulat(Ambil(kd, "tahun"));
                    ws.Teks(r, 3, "Tahun hapus buku " + (th.HasValue ? th.Value.ToString() : "-"), G.Subjudul);
                    ws.Angka(r, 5, th, G.Biasa);
                    var tk = L(Ambil(kd, "tahunKolom"));
                    for (int j = 0; j < 6; j++) ws.Angka(r + 1, 6 + j, tk != null && j < tk.Count ? N(tk[j]) : null, G.Tebal);
                    int lab = r + 2;
                    string[] kepala = { "Nama Nasabah", "Tahun Hapus Buku", "Balance hapus buku", "year 1", "year 2", "year 3", "year 4", "year 5", "year 6" };
                    for (int j = 0; j < kepala.Length; j++) ws.Teks(lab, 3 + j, kepala[j], G.Kepala);
                    ws.Teks(lab, 14, "Jumlah recovery", G.Kepala); ws.Teks(lab, 15, "Recovery Rate", G.Kepala);
                    int first = lab + 1, baris = first;
                    var rek = L(Ambil(kd, "rekening"));
                    double sE = 0, sN = 0;
                    var sF = new double[6];
                    if (rek != null && rek.Count > 0)
                    {
                        foreach (var ro in rek)
                        {
                            var rd = D(ro);
                            double bal = N(Ambil(rd, "balance")) ?? 0;
                            ws.Teks(baris, 3, S(Ambil(rd, "rek")), G.Teks);
                            ws.Angka(baris, 4, Bulat(Ambil(rd, "tahunHB")), G.Biasa);
                            ws.Angka(baris, 5, bal, G.Input);
                            var sd = L(Ambil(rd, "saldo"));
                            for (int j = 0; j < 6; j++)
                            {
                                double? v = sd != null && j < sd.Count ? N(sd[j]) : null;
                                ws.Angka(baris, 6 + j, v, G.Input);
                                sF[j] += v ?? 0;
                            }
                            ws.Teks(baris + 1, 3, "-"); ws.Teks(baris + 1, 5, "recovery", G.Catatan);
                            var rc = L(Ambil(rd, "recovery"));
                            double jr = 0;
                            for (int j = 0; j < 6; j++)
                            {
                                double? v = rc != null && j < rc.Count ? N(rc[j]) : null;
                                ws.Angka(baris + 1, 6 + j, v, G.Input);
                                jr += v ?? 0;
                            }
                            ws.Rumus(baris + 1, 14, "=SUM(F" + (baris + 1) + ":K" + (baris + 1) + ")", jr, G.Angka);
                            ws.Rumus(baris + 1, 15, "=IFERROR(N" + (baris + 1) + "/E" + baris + ",0)", bal == 0 ? 0 : jr / bal, G.Persen);
                            sE += bal; sN += jr;
                            baris += 2;
                        }
                    }
                    else
                    {
                        ws.Teks(baris, 3, "(tidak ada data)", G.Catatan);
                        ws.Teks(baris + 1, 3, "-"); ws.Teks(baris + 1, 5, "recovery", G.Catatan);
                        baris += 2;
                    }
                    int sub = baris, last = sub - 1;
                    ws.Teks(sub, 3, "combine", G.TotalLabel);
                    ws.Rumus(sub, 5, "=SUM(E" + first + ":E" + last + ")", sE, G.Total);
                    for (int j = 0; j < 6; j++)
                    {
                        string c = XlsxTulis.Kol(6 + j);
                        ws.Rumus(sub, 6 + j, "=SUMPRODUCT((MOD(ROW(" + c + first + ":" + c + last + ")-" + first + ",2)=0)*" + c + first + ":" + c + last + ")",
                                 sF[j], G.Total);
                    }
                    ws.Rumus(sub, 14, "=SUM(N" + first + ":N" + last + ")", sN, G.Total);
                    ws.Rumus(sub, 15, "=IFERROR(N" + sub + "/E" + sub + ",0)", sE == 0 ? 0 : sN / sE, G.TotalPersen);
                    subE.Add("E" + sub); subN.Add("N" + sub);
                    totE += sE; totN += sN;
                    r = sub + 1;
                }
            }
            else
            {
                ws.Teks(r, 3, "Rincian LGD ER per kohort tidak tersimpan pada versi ini — ringkasan memakai angka tersimpan.", G.Peringatan);
                r += 1;
            }

            int tot = r + 1;
            if (rinci)
            {
                ws.Teks(tot, 3, "TOTAL COMBINE", G.TotalLabel);
                ws.Rumus(tot, 5, "=" + string.Join("+", subE.ToArray()), totE, G.Total);
                ws.Rumus(tot, 14, "=" + string.Join("+", subN.ToArray()), totN, G.Total);
                ws.Rumus(tot, 15, "=IFERROR(N" + tot + "/E" + tot + ",0)", totE == 0 ? 0 : totN / totE, G.TotalPersen);
            }
            int ring = tot + 2, wo = ring + 1;
            ws.Teks(ring, 3, "Ringkasan LGD Weighted", G.Subjudul);
            ws.Teks(wo, 3, "Total WO", G.Tebal); ws.Teks(wo + 1, 3, "Total Recovery", G.Tebal); ws.Teks(wo + 2, 3, "LGD Weighted", G.Tebal);

            // Rincian dipakai hanya bila totalnya sama dengan yang tersimpan di versi (ringkasan lgd_er_wo/_rec)
            bool cocok = rinci && (!woRg.HasValue || Math.Abs(totE - woRg.Value) <= TolRp) && (!recRg.HasValue || Math.Abs(totN - recRg.Value) <= TolRp);
            double woV, recV;
            if (cocok)
            {
                ws.Rumus(wo, 4, "=E" + tot, totE, G.AngkaTebal);
                ws.Rumus(wo + 1, 4, "=N" + tot, totN, G.AngkaTebal);
                woV = totE; recV = totN;
            }
            else
            {
                woV = woRg ?? 0; recV = recRg ?? 0;
                ws.Angka(wo, 4, woV, G.Nilai);
                ws.Angka(wo + 1, 4, recV, G.Nilai);
                if (rinci)
                    k.Catatan.Add("B3: total rincian kohort (WO " + Rp(totE) + ", recovery " + Rp(totN) + ") berbeda dengan ringkasan versi (WO " +
                                  Rp(woV) + ", recovery " + Rp(recV) + ") — ringkasan LGD ER memakai angka tersimpan.");
                else if (woRg.HasValue)
                    k.Catatan.Add("B3: rincian LGD ER per kohort tidak tersimpan — Total WO & Recovery ER memakai angka tersimpan.");
            }
            k.ErWo = woV; k.ErRec = recV;
            k.LgdEr = woV == 0 ? 0 : 1 - recV / woV;
            ws.Rumus(wo + 2, 4, "=IFERROR(1-D" + (wo + 1) + "/D" + wo + ",0)", k.LgdEr, G.PersenTebal);
            return wo;
        }

        // ================================================================
        // B4. LGD Collateral Shortfall + LGD Weighted gabungan
        // ================================================================
        private static void SheetB4(XlsxTulis.Lembar ws, Ktx k, int b3Wo)
        {
            ws.Teks(1, 2, "LGD - Collateral Shortfall", G.Judul);
            ws.Teks(2, 2, "Perhitungan Loss Given Default berdasarkan kekurangan nilai eksekusi agunan");
            ws.Teks(3, 2, "Kolom G (nilai hasil eksekusi) = angka versi tersimpan, termasuk koreksi/penyesuaian. Ubah kolom biru untuk simulasi — " +
                          "LGD Weighted dan CKPN Kolektif ikut berubah.", G.Catatan);
            ws.Lebar(1, 2); ws.Lebar(2, 18);
            for (int c = 3; c <= 8; c++) ws.Lebar(c, 16);
            ws.Lebar(9, 28); ws.Lebar(10, 16);
            string[] kepala = { "Nomor Rekening kredit", "Baki debet/ outstanding kredit yg belum lunas", "Nilai agunan yang dijaminkan",
                                "Tahun agunan diserahkan kpd bank/tahun hapus buku", "Tahun eksekusi(Tahun Selesai/Lunas)",
                                "Nilai agunan hasil eksekusi setelah dikurangi biaya-biaya", "shortfall (kerugian)", "Nama Debitur", "Sumber baris" };
            for (int j = 0; j < kepala.Length; j++) ws.Teks(4, 2 + j, kepala[j], G.Kepala);
            ws.Beku(4, 2);

            var cs = L(k.Data["lgdcs"]) ?? new ArrayList();
            int first = 5, r = first;
            double wo = 0, rec = 0, sf = 0;
            foreach (var o in cs)
            {
                var d = D(o);
                double pokok = N(Ambil(d, "pokok")) ?? 0, agn = N(Ambil(d, "agunan")) ?? 0, rcv = N(Ambil(d, "recovery")) ?? 0;
                ws.Teks(r, 2, S(Ambil(d, "rek")), G.Teks);
                ws.Angka(r, 3, pokok, G.Input);
                ws.Angka(r, 4, agn, G.Input);
                ws.Teks(r, 5, S(Ambil(d, "thnSerah")), G.Teks);
                ws.Teks(r, 6, S(Ambil(d, "thnEks")), G.Teks);
                ws.Angka(r, 7, rcv, G.Input);
                double h = rcv > pokok ? 0 : pokok - rcv;
                ws.Rumus(r, 8, "=IF(G" + r + ">C" + r + ",0,C" + r + "-G" + r + ")", h, G.Angka);
                ws.Teks(r, 9, S(Ambil(d, "nama")));
                ws.Teks(r, 10, LabelSumber(S(Ambil(d, "sumber"))));
                wo += pokok; rec += rcv; sf += h;
                r++;
            }
            int last = Math.Max(first, r - 1), total = Math.Max(first, r - 1) + 1;
            if (cs.Count == 0) total = first + 1;
            ws.Teks(total, 2, "Total", G.TotalLabel);
            ws.Rumus(total, 3, "=IFERROR(SUM(C" + first + ":C" + last + "),0)", wo, G.Total);
            ws.Rumus(total, 7, "=IFERROR(SUM(G" + first + ":G" + last + "),0)", rec, G.Total);
            ws.Rumus(total, 8, "=IFERROR(SUM(H" + first + ":H" + last + "),0)", sf, G.Total);
            for (int c = 4; c <= 6; c++) ws.Gayakan(total, c, G.Total);

            int hdr = total + 2, s0 = total + 3;
            ws.Teks(hdr, 3, "Collateral Shortfall", G.Kepala); ws.Teks(hdr, 4, "Expected Recoveries", G.Kepala); ws.Teks(hdr, 5, "Combined", G.Kepala);
            ws.Teks(s0, 2, "Total WO", G.Tebal); ws.Teks(s0 + 1, 2, "Total Recovery", G.Tebal); ws.Teks(s0 + 2, 2, "LGD Weighted", G.Tebal);

            // ER: dari ringkasan B3 (rumus bila B3 punya rincian, angka tersimpan bila tidak)
            double woEr = k.ErWo, recEr = k.ErRec;
            ws.Rumus(s0, 4, "=" + XlsxTulis.RefLuar(ShB3, b3Wo, 4), woEr, G.Angka);
            ws.Rumus(s0 + 1, 4, "=" + XlsxTulis.RefLuar(ShB3, b3Wo + 1, 4), recEr, G.Angka);
            ws.Rumus(s0, 3, "=C" + total, wo, G.Angka);
            ws.Rumus(s0 + 1, 3, "=G" + total, rec, G.Angka);
            ws.Rumus(s0, 5, "=C" + s0 + "+D" + s0, wo + woEr, G.Angka);
            ws.Rumus(s0 + 1, 5, "=C" + (s0 + 1) + "+D" + (s0 + 1), rec + recEr, G.Angka);
            k.LgdCs = wo == 0 ? 0 : 1 - rec / wo;
            k.LgdGab = (wo + woEr) == 0 ? 0 : 1 - (rec + recEr) / (wo + woEr);
            ws.Rumus(s0 + 2, 3, "=IFERROR(1-C" + (s0 + 1) + "/C" + s0 + ",0)", k.LgdCs, G.PersenTebal);
            ws.Rumus(s0 + 2, 4, "=IFERROR(1-D" + (s0 + 1) + "/D" + s0 + ",0)", woEr == 0 ? 0 : 1 - recEr / woEr, G.PersenTebal);
            ws.Rumus(s0 + 2, 5, "=IFERROR(1-E" + (s0 + 1) + "/E" + s0 + ",0)", k.LgdGab, G.PersenTebal);
            k.B4LgdBaris = s0 + 2;

            double? lgdPakai = N(Ambil(k.Rg, "lgd_gabungan"));
            if (lgdPakai.HasValue && Math.Abs(lgdPakai.Value - k.LgdGab.Value) > TolRate)
            {
                ws.Teks(s0 + 4, 2, "LGD gabungan yang dipakai versi ini: " + Persen(lgdPakai.Value) +
                        " (berbeda dengan hasil sheet ini " + Persen(k.LgdGab.Value) + "). Lihat catatan di sheet Info Versi.", G.Peringatan);
                k.Catatan.Add("B4: LGD gabungan hasil rincian (" + Persen(k.LgdGab.Value) + ") berbeda dengan LGD yang dipakai versi (" + Persen(lgdPakai.Value) +
                              ") — kolom LGD di B. CKPN - KOL INDV memakai angka tersimpan.");
            }
        }

        private static string LabelSumber(string s)
        {
            if (string.IsNullOrEmpty(s) || s == "sistem") return "sistem";
            if (s.StartsWith("manual")) return "baris manual" + (s.Contains("koreksi") ? " (dikoreksi)" : "");
            if (s.StartsWith("koreksi")) return "koreksi panel";
            if (s.Contains("koreksi")) return s.Replace("+koreksi", "") + " (dikoreksi)";
            if (s.StartsWith("penyesuaian") || s == "ubah") return "penyesuaian";
            return s;
        }

        // ================================================================
        // A. CKPN Individu
        // ================================================================
        private static void SheetA(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Teks(2, 2, "Penurunan Nilai - Penilaian Individu", G.Judul);
            ws.Teks(3, 2, "Periode", G.Tebal); ws.Teks(3, 3, TanggalPanjang(k.Tanggal));
            ws.Teks(3, 4, "Debitur signifikan", G.Tebal); ws.Teks(3, 5, S(Ambil(k.Run, "labelSignifikan")));
            ws.Teks(4, 2, "Grup: " + k.NamaGrup + " (" + k.KodeKC + ") · versi v" + k.Versi + " dari database. Kolom biru boleh diubah untuk simulasi.", G.Catatan);
            string[] kepala = { "No.", "Sheet KC", "Nomor Rekening", "Nama Debitur", "Nomor Kontrak", "OS Pembiayaan", "Ada Penurunan Nilai?",
                                "Nilai Jaminan", "Biaya Penjualan", "Penurunan Nilai", "Keterangan" };
            for (int j = 0; j < kepala.Length; j++) ws.Teks(5, 2 + j, kepala[j], G.Kepala);
            ws.Lebar(1, 2); ws.Lebar(2, 6); ws.Lebar(3, 9); ws.Lebar(4, 14); ws.Lebar(5, 30); ws.Lebar(6, 15);
            for (int c = 7; c <= 11; c++) ws.Lebar(c, 15);
            ws.Lebar(8, 11); ws.Lebar(12, 22);
            ws.Beku(5, 3);

            var ind = L(k.Data["individu"]) ?? new ArrayList();
            int r = 6, no = 0, beda = 0;
            foreach (var o in ind)
            {
                var d = D(o);
                no++;
                double os = N(Ambil(d, "os")) ?? 0, jam = N(Ambil(d, "jaminan")) ?? 0, bj = N(Ambil(d, "biaya")) ?? 0, ck = N(Ambil(d, "ckpn")) ?? 0;
                string pn = S(Ambil(d, "adaPN"));
                int? urut = Bulat(Ambil(d, "urut"));
                ws.Angka(r, 2, urut ?? no, G.Biasa);
                ws.Teks(r, 3, S(Ambil(d, "kc")));
                ws.Teks(r, 4, S(Ambil(d, "cif")), G.Teks);
                ws.Teks(r, 5, S(Ambil(d, "nama")));
                ws.Teks(r, 6, S(Ambil(d, "kontrak")), G.Teks);
                ws.Angka(r, 7, os, G.Input);
                ws.Teks(r, 8, pn);
                ws.Angka(r, 9, jam, G.Input);
                ws.Angka(r, 10, bj, G.Input);
                double rumus = pn == "Ya" ? Math.Max(0, os - (jam - bj)) : 0;
                if (Math.Abs(rumus - ck) <= TolRp)
                    ws.Rumus(r, 11, "=IF(H" + r + "=\"Ya\",MAX(0,G" + r + "-(I" + r + "-J" + r + ")),0)", rumus, G.Angka);
                else { ws.Angka(r, 11, ck, G.Nilai); beda++; }
                bool sesuai = Ambil(d, "disesuaikan") is bool && (bool)Ambil(d, "disesuaikan");
                ws.Teks(r, 12, sesuai ? "agunan/biaya disesuaikan" : "");
                k.AOs += os; k.AJam += jam; k.ACkpn += ck;
                r++;
            }
            int total = ind.Count == 0 ? 7 : r, last = total - 1;
            ws.Teks(total, 2, "TOTAL", G.TotalLabel);
            for (int c = 3; c <= 6; c++) ws.Gayakan(total, c, G.Total);
            ws.Rumus(total, 7, "=SUM(G6:G" + last + ")", k.AOs, G.Total);
            ws.Gayakan(total, 8, G.Total);
            ws.Rumus(total, 9, "=SUM(I6:I" + last + ")", k.AJam, G.Total);
            double tb = 0; foreach (var o in ind) tb += N(Ambil(D(o), "biaya")) ?? 0;
            ws.Rumus(total, 10, "=SUM(J6:J" + last + ")", tb, G.Total);
            ws.Rumus(total, 11, "=SUM(K6:K" + last + ")", k.ACkpn, G.Total);
            ws.Gayakan(total, 12, G.Total);
            k.ATotal = total;
            if (beda > 0)
                k.Catatan.Add("A: " + beda + " kontrak memakai penurunan nilai tersimpan (sel kuning) karena tidak sama dengan rumus " +
                              "MAX(0, OS − (jaminan − biaya)) — mis. penilaian individu dengan arus kas.");
        }

        // ================================================================
        // B. CKPN Kolektif (Net Flow per bucket, Migration per kualitas)
        // ================================================================
        private static void SheetKol(XlsxTulis.Lembar ws, Ktx k)
        {
            var kol = k.Data["kolektif"] as Dictionary<string, object>;
            var nf = kol == null ? null : L(Ambil(kol, "netflow"));
            var mg = kol == null ? null : L(Ambil(kol, "migrasi"));
            bool setahun = kol != null && S(Ambil(kol, "mode")) == "setahun";
            ws.Lebar(1, 2); ws.Lebar(2, 12); ws.Lebar(3, 18);
            for (int c = 4; c <= 7; c++) ws.Lebar(c, 17);
            ws.Lebar(8, 40);

            ws.Teks(2, 2, "1. PERHITUNGAN PENURUNAN NILAI BERDASARKAN HARI TUNGGAKAN MENGGUNAKAN METODE NETFLOW DENGAN WEIGHTED LGD " +
                          "(LGD EXPECTED RECOVERIES DAN LGD COLLATERAL SHORTFALL)", G.Subjudul);
            ws.Teks(3, 2, "Grup " + k.NamaGrup + " · " + TanggalPanjang(k.Tanggal) + " · v" + k.Versi + (setahun ? " · PD & LGD dihitung setahun sekali" : ""), G.Catatan);
            ws.Teks(4, 2, "Kolektif", G.Tebal);
            string[] kp = { "Bucket", "Hari Tunggakan", "Baki Debet (EAD)", "PD Net Flow", "LGD Weighted", "Penurunan Nilai", "Sumber angka" };
            for (int j = 0; j < kp.Length; j++) ws.Teks(5, 2 + j, kp[j], G.Kepala);

            string refLgd = XlsxTulis.RefLuar(ShB4, k.B4LgdBaris, 5);
            var eadNf = new double[14];
            double jml = 0;
            int nLink = 0, nNilai = 0;
            for (int b = 0; b < 14; b++)
            {
                int r = 6 + b;
                ws.Angka(r, 2, b + 1, G.Biasa); ws.Teks(r, 3, LabelBucketKol[b]);
                var d = nf != null && b < nf.Count ? D(nf[b]) : null;
                double? ead = d == null ? null : N(Ambil(d, "ead")), pd = d == null ? null : N(Ambil(d, "pd")), lgd = d == null ? null : N(Ambil(d, "lgd"));
                var sumber = new List<string>();
                // EAD
                if (ead.HasValue && k.B1Ead[b].HasValue && Math.Abs(ead.Value - k.B1Ead[b].Value) <= TolRp && !setahun)
                    ws.Rumus(r, 4, "=" + XlsxTulis.RefLuar(ShB1, b < 13 ? 4 + b : 18, 16), k.B1Ead[b].Value, G.Angka);
                else { ws.Angka(r, 4, ead, setahun ? G.Input : G.Nilai); if (!setahun && ead.HasValue) sumber.Add("EAD tersimpan"); }
                eadNf[b] = ead ?? 0;
                // PD
                if (b == 13 && pd.HasValue && Math.Abs(pd.Value - 1) <= TolRate) ws.Angka(r, 5, 1, G.PdInput);
                else if (b < 13 && pd.HasValue && k.B1Pd[b].HasValue && Math.Abs(pd.Value - k.B1Pd[b].Value) <= TolRate && k.B1Rumus)
                    ws.Rumus(r, 5, "=" + XlsxTulis.RefLuar(ShB1, 40 + b, 16), k.B1Pd[b].Value, G.Pd);
                else { ws.Angka(r, 5, pd, G.PersenNilai); if (pd.HasValue) sumber.Add("PD tersimpan"); }
                // LGD
                if (lgd.HasValue && k.LgdGab.HasValue && Math.Abs(lgd.Value - k.LgdGab.Value) <= TolRate)
                    ws.Rumus(r, 6, "=" + refLgd, k.LgdGab.Value, G.Pd);
                else { ws.Angka(r, 6, lgd, G.PersenNilai); if (lgd.HasValue) sumber.Add("LGD tersimpan"); }
                double ck = (ead ?? 0) * (pd ?? 0) * (lgd ?? 0);
                jml += ck;
                ws.Rumus(r, 7, "=D" + r + "*E" + r + "*F" + r, ck, G.Angka);
                if (sumber.Count > 0) { ws.Teks(r, 8, string.Join(", ", sumber.ToArray()), G.Catatan); nNilai++; } else nLink++;
            }
            ws.Teks(20, 2, "Jumlah cadangan yang harus dibentuk/jumlah penurunan nilai =", G.TotalLabel);
            for (int c = 3; c <= 6; c++) ws.Gayakan(20, c, G.Total);
            ws.Rumus(20, 7, "=SUM(G6:G19)", jml, G.Total);
            k.KolNf = jml;
            if (nf == null)
            {
                double? v = N(Ambil(k.Rg, "nf_kolektif"));
                ws.Angka(20, 7, v, G.Nilai);
                k.KolNf = v ?? 0;
                k.Catatan.Add("B. KOL: rincian kolektif per bucket tidak tersimpan pada versi ini — total Kolektif Net Flow memakai angka tersimpan.");
            }
            if (setahun)
                k.Catatan.Add("B. KOL: mode setahun sekali — EAD per bucket adalah baki debet posisi " + TanggalPanjang(k.Tanggal) +
                              " (bukan bulan terakhir sheet B1).");

            // ---- Individual ----
            string refA = "'" + ShA + "'!";
            ws.Teks(22, 2, "Individual", G.Subjudul);
            ws.Teks(23, 2, "Individual", G.Kepala); ws.Teks(23, 3, "Nilai Tercatat", G.Kepala);
            ws.Teks(23, 4, "Estimasi arus kas diterima", G.Kepala); ws.Teks(23, 5, "Penurunan Nilai", G.Kepala);
            ws.Teks(24, 2, "Individu 1");
            ws.Rumus(24, 3, "=" + refA + "$G$" + k.ATotal, k.AOs, G.Angka);
            ws.Rumus(24, 4, "=" + refA + "$I$" + k.ATotal, k.AJam, G.Angka);
            ws.Rumus(24, 5, "=" + refA + "$K$" + k.ATotal, k.ACkpn, G.Angka);
            k.KolInd = k.ACkpn;

            ws.Teks(26, 2, "Total CKPN", G.Subjudul);
            ws.Teks(27, 2, "Kategori", G.Kepala); ws.Teks(27, 3, "Penurunan Nilai", G.Kepala);
            ws.Teks(28, 2, "Individual"); ws.Rumus(28, 3, "=E24", k.ACkpn, G.Angka);
            ws.Teks(29, 2, "Kolektif"); ws.Rumus(29, 3, "=G20", k.KolNf, G.Angka);
            ws.Teks(30, 2, "Total CKPN", G.TotalLabel); ws.Rumus(30, 3, "=SUM(C28:C29)", k.ACkpn + k.KolNf, G.Total);

            // ---- Migration ----
            ws.Teks(34, 2, "3. PERHITUNGAN PENURUNAN NILAI BERDASARKAN KUALITAS MENGGUNAKAN METODE MIGRATION DAN WEIGHTED LGD " +
                           "(LGD EXPECTED RECOVERIES DAN LGD COLLATERAL SHORTFALL)", G.Subjudul);
            string[] km = { "Kualitas", "Baki Debet (EAD)", "PD Migration", "LGD Weighted", "Penurunan Nilai", "Sumber angka" };
            for (int j = 0; j < km.Length; j++) ws.Teks(36, 2 + j, km[j], G.Kepala);
            // EAD per kualitas = jumlah bucket (bulanan): Kol1 = 0 hari..x, Kol2 = s.d. 90 hari, dst. (sama dengan template)
            int[,] rentang = { { 6, 6 }, { 7, 9 }, { 10, 12 }, { 13, 15 }, { 16, 19 } };
            double jmlM = 0;
            for (int q = 0; q < 5; q++)
            {
                int r = 38 + q;
                ws.Angka(r, 2, q + 1, G.Biasa);
                var d = mg != null && q < mg.Count ? D(mg[q]) : null;
                double? ead = d == null ? null : N(Ambil(d, "ead")), pd = d == null ? null : N(Ambil(d, "pd")), lgd = d == null ? null : N(Ambil(d, "lgd"));
                var sumber = new List<string>();
                int r1 = rentang[q, 0], r2 = rentang[q, 1];
                double dariNf = 0; for (int x = r1; x <= r2; x++) dariNf += eadNf[x - 6];
                if (ead.HasValue && nf != null && !setahun && Math.Abs(ead.Value - dariNf) <= TolRp)
                    ws.Rumus(r, 3, r1 == r2 ? "=D" + r1 : "=SUM(D" + r1 + ":D" + r2 + ")", dariNf, G.Angka);
                else { ws.Angka(r, 3, ead, setahun ? G.Input : G.Nilai); if (!setahun && ead.HasValue) sumber.Add("EAD tersimpan"); }
                if (pd.HasValue && k.B2Ada && k.B2Pd[q].HasValue && Math.Abs(pd.Value - k.B2Pd[q].Value) <= TolRate)
                    ws.Rumus(r, 4, "=" + XlsxTulis.RefLuar(ShB2, 82 + q, 2), k.B2Pd[q].Value, G.Pd);
                else { ws.Angka(r, 4, pd, G.PersenNilai); if (pd.HasValue) sumber.Add("PD tersimpan"); }
                if (lgd.HasValue && k.LgdGab.HasValue && Math.Abs(lgd.Value - k.LgdGab.Value) <= TolRate)
                    ws.Rumus(r, 5, "=" + refLgd, k.LgdGab.Value, G.Pd);
                else { ws.Angka(r, 5, lgd, G.PersenNilai); if (lgd.HasValue) sumber.Add("LGD tersimpan"); }
                double ck = (ead ?? 0) * (pd ?? 0) * (lgd ?? 0);
                jmlM += ck;
                ws.Rumus(r, 6, "=C" + r + "*D" + r + "*E" + r, ck, G.Angka);
                if (sumber.Count > 0) { ws.Teks(r, 7, string.Join(", ", sumber.ToArray()), G.Catatan); nNilai++; } else nLink++;
            }
            ws.Teks(43, 2, "Jumlah cadangan yang harus dibentuk/jumlah penurunan nilai =", G.TotalLabel);
            for (int c = 3; c <= 5; c++) ws.Gayakan(43, c, G.Total);
            ws.Rumus(43, 6, "=SUM(F38:F42)", jmlM, G.Total);
            k.KolMig = jmlM;
            if (mg == null)
            {
                double? v = N(Ambil(k.Rg, "mig_kolektif"));
                ws.Angka(43, 6, v, G.Nilai);
                k.KolMig = v ?? 0;
                k.Catatan.Add("B. KOL: rincian kolektif per kualitas tidak tersimpan pada versi ini — total Kolektif Migration memakai angka tersimpan.");
            }

            ws.Teks(45, 2, "Individual", G.Subjudul);
            ws.Teks(46, 2, "Individual", G.Kepala); ws.Teks(46, 3, "Nilai Tercatat", G.Kepala);
            ws.Teks(46, 4, "PV of CF", G.Kepala); ws.Teks(46, 5, "Penurunan Nilai", G.Kepala);
            ws.Teks(47, 2, "Individu 1");
            ws.Rumus(47, 3, "=C24", k.AOs, G.Angka); ws.Rumus(47, 4, "=D24", k.AJam, G.Angka); ws.Rumus(47, 5, "=E24", k.ACkpn, G.Angka);
            ws.Teks(49, 2, "Total CKPN", G.Subjudul);
            ws.Teks(50, 2, "Kategori", G.Kepala); ws.Teks(50, 3, "Penurunan Nilai", G.Kepala);
            ws.Teks(51, 2, "Individual"); ws.Rumus(51, 3, "=E47", k.ACkpn, G.Angka);
            ws.Teks(52, 2, "Kolektif"); ws.Rumus(52, 3, "=F43", k.KolMig, G.Angka);
            ws.Teks(53, 2, "Total CKPN", G.TotalLabel); ws.Rumus(53, 3, "=SUM(C51:C52)", k.ACkpn + k.KolMig, G.Total);
            k.KolAda = nf != null || mg != null;
        }

        // ================================================================
        // Summary — CKPN Net Flow / Migration, ABA, PPKA vs CKPN (rumus sama dengan template)
        // ================================================================
        private static void SheetSummary(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Lebar(1, 46); ws.Lebar(2, 20); ws.Lebar(3, 20); ws.Lebar(6, 3); ws.Lebar(7, 26); ws.Lebar(8, 18);
            string refB = "'" + ShB + "'!";
            ws.Teks(1, 1, "SUMMARY PERHITUNGAN PSAK 414", G.Judul);
            ws.Teks(2, 1, "Grup " + k.NamaGrup + " (" + k.KodeKC + ") · " + TanggalPanjang(k.Tanggal) + " · versi v" + k.Versi + " — dibangun dari database", G.Catatan);

            ws.Teks(3, 1, "A. Individual + Collective PD Net Flow Rate dan Weighted LGD (LGD Expected Recoveries and LGD Collateral Shortfall)", G.Subjudul);
            ws.Teks(5, 1, "Kategori", G.Kepala); ws.Teks(5, 2, "CKPN Penurunan Nilai", G.Kepala); ws.Teks(5, 3, "Lapbul", G.Kepala);
            ws.Teks(6, 1, "Individual"); ws.Rumus(6, 2, "=" + refB + "$C$28", k.ACkpn, G.Angka); ws.Angka(6, 3, N(Ambil(k.Rg, "tpl_individu")), G.Input);
            ws.Teks(7, 1, "Kolektif"); ws.Rumus(7, 2, "=" + refB + "$C$29", k.KolNf, G.Angka); ws.Angka(7, 3, N(Ambil(k.Rg, "tpl_kolektif")), G.Input);
            ws.Teks(8, 1, "Total CKPN", G.TotalLabel);
            ws.Rumus(8, 2, "=SUM(B6:B7)", k.ACkpn + k.KolNf, G.Total);
            ws.Rumus(8, 3, "=SUM(C6:C7)", (N(Ambil(k.Rg, "tpl_individu")) ?? 0) + (N(Ambil(k.Rg, "tpl_kolektif")) ?? 0), G.Total);
            k.SumB6 = k.ACkpn; k.SumB7 = k.KolNf;

            ws.Teks(10, 1, "C. Individual + Collective PD Migration dan Weighted LGD (LGD Expected Recoveries and LGD Collateral Shortfall)", G.Subjudul);
            ws.Teks(12, 1, "Kategori", G.Kepala); ws.Teks(12, 2, "CKPN Penurunan Nilai", G.Kepala); ws.Teks(12, 3, "Lapbul", G.Kepala);
            ws.Teks(13, 1, "Individual"); ws.Rumus(13, 2, "=" + refB + "$C$51", k.ACkpn, G.Angka); ws.Rumus(13, 3, "=C6", N(Ambil(k.Rg, "tpl_individu")) ?? 0, G.Angka);
            ws.Teks(14, 1, "Kolektif"); ws.Rumus(14, 2, "=" + refB + "$C$52", k.KolMig, G.Angka); ws.Rumus(14, 3, "=C7", N(Ambil(k.Rg, "tpl_kolektif")) ?? 0, G.Angka);
            ws.Teks(15, 1, "Total CKPN", G.TotalLabel);
            ws.Rumus(15, 2, "=SUM(B13:B14)", k.ACkpn + k.KolMig, G.Total);
            ws.Rumus(15, 3, "=SUM(C13:C14)", (N(Ambil(k.Rg, "tpl_individu")) ?? 0) + (N(Ambil(k.Rg, "tpl_kolektif")) ?? 0), G.Total);
            k.SumB13 = k.ACkpn; k.SumB14 = k.KolMig;

            // ---- D. ABA ----
            ws.Teks(17, 1, "D. Perhitungan CKPN Aset Antar Bank (ABA - KC0500)", G.Subjudul);
            double? d1 = N(Ambil(k.Rg, "aba_dijamin")), d2 = N(Ambil(k.Rg, "aba_di_atas_plafon")), pdA = N(Ambil(k.Rg, "aba_pd")),
                    l1 = N(Ambil(k.Rg, "aba_lgd_dijamin")), l2 = N(Ambil(k.Rg, "aba_lgd_atas")), abaCk = N(Ambil(k.Rg, "aba_ckpn"));
            ws.Teks(18, 1, "EAD porsi dijamin LPS"); ws.Angka(18, 3, d1, G.Input);
            ws.Teks(19, 1, "EAD porsi di atas plafon LPS"); ws.Angka(19, 3, d2, G.Input);
            ws.Teks(20, 1, "Total EAD ABA", G.Tebal); ws.Rumus(20, 3, "=C18+C19", (d1 ?? 0) + (d2 ?? 0), G.AngkaTebal);
            ws.Teks(21, 1, "PD Industri (asumsi LPS)"); ws.Angka(21, 3, pdA, G.PersenInput);
            ws.Teks(22, 1, "LGD porsi dijamin"); ws.Angka(22, 3, l1, G.PersenInput);
            ws.Teks(23, 1, "LGD porsi di atas plafon"); ws.Angka(23, 3, l2, G.PersenInput);
            double c24 = (d1 ?? 0) * (pdA ?? 0) * (l1 ?? 0), c25 = (d2 ?? 0) * (pdA ?? 0) * (l2 ?? 0);
            ws.Teks(24, 1, "CKPN porsi dijamin"); ws.Rumus(24, 3, "=C18*C21*C22", c24, G.Angka);
            ws.Teks(25, 1, "CKPN porsi di atas plafon"); ws.Rumus(25, 3, "=C19*C21*C23", c25, G.Angka);
            ws.Teks(26, 1, "Total CKPN Antar Bank (ABA)", G.TotalLabel); ws.Gayakan(26, 2, G.Total);
            if (!abaCk.HasValue || Math.Abs(abaCk.Value - (c24 + c25)) <= TolRp)
            {
                ws.Rumus(26, 3, "=C24+C25", c24 + c25, G.Total);
                k.SumC26 = c24 + c25;
            }
            else
            {
                ws.Angka(26, 3, abaCk, G.Nilai);
                k.SumC26 = abaCk.Value;
                k.Catatan.Add("Summary: CKPN ABA tersimpan (" + Rp(abaCk.Value) + ") tidak sama dengan EAD × PD × LGD ABA — dipakai angka tersimpan.");
            }

            // ---- PPKA per KC ----
            ws.Teks(1, 7, "PPKA per KC (KC grup ini + ABA)", G.Subjudul);
            string[] kc = { "KC0500", "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };
            string[] lab = { "PPKA ABA KC500", "PPKA KC600", "PPKA KC700", "PPKA KC800", "PPKA KC900", "PPKA KC1000", "PPKA KC1100" };
            double tp = 0;
            for (int i = 0; i < kc.Length; i++)
            {
                double? v = N(Ambil(k.Rg, "ppka_" + kc[i]));
                ws.Teks(2 + i, 7, lab[i]); ws.Angka(2 + i, 8, v ?? 0, G.Input);
                tp += v ?? 0;
            }
            ws.Teks(9, 7, "TOTAL PPKA", G.TotalLabel); ws.Rumus(9, 8, "=SUM(H2:H8)", tp, G.Total);
            k.SumH9 = tp;

            // ---- E. PPKA vs CKPN ----
            string metode = S(k.Data.ContainsKey("metode") ? k.Data["metode"] : null);
            string m = metode == "mig" ? "C" : "A";
            ws.Teks(28, 1, "E. Perbandingan PPKA vs CKPN — HANYA GRUP INI (bukan dasar jurnal)", G.Subjudul);
            ws.Teks(29, 1, "Metode CKPN dipakai (A = Net Flow, C = Migration)"); ws.Teks(29, 3, m, G.Input);
            double ckpn = (m == "A" ? k.ACkpn + k.KolNf : k.ACkpn + k.KolMig) + k.SumC26;
            ws.Teks(30, 1, "Total PPKA terbentuk"); ws.Rumus(30, 3, "=H9", tp, G.Angka);
            ws.Teks(31, 1, "Total CKPN (metode terpilih + ABA)"); ws.Rumus(31, 3, "=IF($C$29=\"A\",$B$8,$B$15)+C26", ckpn, G.Angka);
            double sel = tp - ckpn;
            ws.Teks(32, 1, "Selisih (PPKA - CKPN)", G.Tebal); ws.Rumus(32, 3, "=C30-C31", sel, G.AngkaTebal);
            ws.Teks(33, 1, "Kondisi");
            ws.Rumus(33, 3, "=IF(C32>0,\"PPKA lebih besar dari CKPN\",IF(C32<0,\"PPKA lebih kecil dari CKPN\",\"PPKA = CKPN\"))",
                     sel > 0 ? "PPKA lebih besar dari CKPN" : sel < 0 ? "PPKA lebih kecil dari CKPN" : "PPKA = CKPN", G.Biasa);
            ws.Teks(34, 1, "Nilai penyesuaian"); ws.Rumus(34, 3, "=ABS(C32)", Math.Abs(sel), G.Angka);
            ws.Teks(35, 1, "Jurnal transisi (1x - awal 2029)");
            ws.Rumus(35, 3, "=IF(C32<0,\"Db. Laba - Kr. CKPN (bentuk tambahan CKPN)\",IF(C32>0,\"Db. CKPN - Kr. Saldo Laba (pemulihan kelebihan)\",\"Tidak ada jurnal\"))",
                     sel < 0 ? "Db. Laba - Kr. CKPN (bentuk tambahan CKPN)" : sel > 0 ? "Db. CKPN - Kr. Saldo Laba (pemulihan kelebihan)" : "Tidak ada jurnal", G.Biasa);
            ws.Teks(36, 1, "Jurnal reguler (periode berikutnya)");
            ws.Rumus(36, 3, "=IF(C32<0,\"Db. Biaya - Kr. CKPN (CKPN kurang)\",IF(C32>0,\"Db. CKPN - Kr. Biaya/Pendapatan (CKPN lebih)\",\"Tidak ada jurnal\"))",
                     sel < 0 ? "Db. Biaya - Kr. CKPN (CKPN kurang)" : sel > 0 ? "Db. CKPN - Kr. Biaya/Pendapatan (CKPN lebih)" : "Tidak ada jurnal", G.Biasa);
            ws.Teks(38, 1, "Catatan: bagian E hanya menghitung grup ini. CKPN ABA (C26) ikut di Summary setiap grup, jadi jangan dijumlahkan antargrup.", G.Catatan);
        }

        // ================================================================
        // Info Versi — identitas, legenda warna, kontrol angka vs database, penyesuaian dalam versi
        // ================================================================
        private static void SheetInfo(XlsxTulis.Lembar ws, Ktx k)
        {
            ws.Lebar(1, 34); ws.Lebar(2, 22); ws.Lebar(3, 22); ws.Lebar(4, 18); ws.Lebar(5, 14); ws.Lebar(6, 30);
            ws.Teks(1, 1, "Kertas kerja CKPN — " + k.NamaGrup + " · " + TanggalPanjang(k.Tanggal) + " · v" + k.Versi, G.Judul);
            ws.Teks(2, 1, "Dibangun dari database aplikasi CKPN pada " + DateTime.Now.ToString("d MMM yyyy HH:mm", new CultureInfo("id-ID")) +
                          ". Angka sama dengan versi tersimpan; rumus hidup dihitung ulang Excel saat file dibuka.", G.Catatan);
            var mode = k.Data["mode"] as Dictionary<string, object>;
            bool aktif = Ambil(k.Run, "aktif") is bool && (bool)Ambil(k.Run, "aktif");
            bool dihapus = Ambil(k.Run, "dihapus") is bool && (bool)Ambil(k.Run, "dihapus");
            string metode = S(k.Data.ContainsKey("metode") ? k.Data["metode"] : null);
            var kv = new List<string[]>
            {
                new[] { "Periode", TanggalPanjang(k.Tanggal) + (S(Ambil(k.Run, "statusPeriode")) == "Final" ? " (Final)" : "") },
                new[] { "Grup", k.NamaGrup },
                new[] { "Kode grup / KC", S(Ambil(k.Grup, "labelKode")).Length > 0 ? S(Ambil(k.Grup, "labelKode")) : k.KodeKC },
                new[] { "Mode perhitungan", mode == null ? "" : S(Ambil(mode, "nama")) },
                new[] { "Versi", "v" + k.Versi + (dihapus ? " (dihapus)" : aktif ? " (aktif)" : " (tidak aktif)") },
                new[] { "Disimpan oleh", S(Ambil(k.Run, "pengguna")) + " · " + S(Ambil(k.Run, "waktu")) +
                                         (S(Ambil(k.Run, "komputer")).Length > 0 ? " · " + S(Ambil(k.Run, "komputer")) : "") },
                new[] { "Versi add-in", S(Ambil(k.Run, "versiAddin")) },
                new[] { "Debitur signifikan", S(Ambil(k.Run, "labelSignifikan")) },
                new[] { "Metode konsolidasi tahun ini", metode == "mig" ? "Migration (C)" : metode == "nf" ? "Net Flow (A)" : "—" },
                new[] { "Catatan versi", S(Ambil(k.Run, "catatan")) },
                new[] { "Asal angka", k.Rg.ContainsKey("koreksi_dari_versi")
                        ? "Koreksi / penyesuaian panel dari v" + Convert.ToInt32(N(k.Rg["koreksi_dari_versi"]) ?? 0) + " (dihitung di database, sheet aplikasi tidak berubah)"
                        : "Disimpan dari sheet aplikasi" }
            };
            int r = 4;
            foreach (var p in kv) { ws.Teks(r, 1, p[0], G.Tebal); ws.Teks(r, 2, p[1]); r++; }

            r++;
            ws.Teks(r++, 1, "Arti warna sel", G.Subjudul);
            ws.Angka(r, 1, 1234567, G.Angka); ws.Teks(r++, 2, "Rumus hidup — dihitung ulang Excel (sama dengan rumus sheet aplikasi)");
            ws.Angka(r, 1, 1234567, G.Input); ws.Teks(r++, 2, "Data versi tersimpan (biru) — boleh diubah untuk simulasi, total ikut berubah");
            ws.Angka(r, 1, 1234567, G.Nilai); ws.Teks(r++, 2, "Angka tersimpan dipakai apa adanya (kuning) — rumus tidak dapat direkonstruksi persis, lihat catatan");

            // ---- Kontrol angka ----
            r++;
            ws.Teks(r++, 1, "Kontrol: kertas kerja vs database", G.Subjudul);
            ws.Teks(r, 1, "Pos", G.Kepala); ws.Teks(r, 2, "Database (versi)", G.Kepala); ws.Teks(r, 3, "Kertas kerja (rumus)", G.Kepala);
            ws.Teks(r, 4, "Selisih", G.Kepala); ws.Teks(r, 5, "Status", G.Kepala);
            r++;
            string sSum = "'" + ShSum + "'!", sB4 = "'" + ShB4 + "'!";
            var kontrol = new List<object[]>
            {
                new object[] { "CKPN Individu", "nf_individu", "=" + sSum + "$B$6", k.SumB6, false },
                new object[] { "CKPN Kolektif — Net Flow", "nf_kolektif", "=" + sSum + "$B$7", k.SumB7, false },
                new object[] { "Total CKPN — Net Flow", "nf_total", "=" + sSum + "$B$8", k.SumB6 + k.SumB7, false },
                new object[] { "CKPN Kolektif — Migration", "mig_kolektif", "=" + sSum + "$B$14", k.SumB14, false },
                new object[] { "Total CKPN — Migration", "mig_total", "=" + sSum + "$B$15", k.SumB13 + k.SumB14, false },
                new object[] { "LGD Collateral Shortfall", "lgd_cs", "=" + sB4 + "$C$" + k.B4LgdBaris, k.LgdCs ?? 0, true },
                new object[] { "LGD Expected Recoveries", "lgd_er", "=" + sB4 + "$D$" + k.B4LgdBaris, k.LgdEr ?? 0, true },
                new object[] { "LGD gabungan (weighted)", "lgd_gabungan", "=" + sB4 + "$E$" + k.B4LgdBaris, k.LgdGab ?? 0, true },
                new object[] { "CKPN ABA", "aba_ckpn", "=" + sSum + "$C$26", k.SumC26, false },
                new object[] { "Total PPKA", "ppka_total", "=" + sSum + "$H$9", k.SumH9, false }
            };
            int beda = 0;
            foreach (var kr in kontrol)
            {
                double? db = N(Ambil(k.Rg, (string)kr[1]));
                if (!db.HasValue) continue;
                bool rate = (bool)kr[4];
                double kk = (double)kr[3];
                double sel = kk - db.Value;
                bool sama = Math.Abs(sel) <= (rate ? 1e-6 : TolRp);
                if (!sama) beda++;
                ws.Teks(r, 1, (string)kr[0]);
                ws.Angka(r, 2, db.Value, rate ? G.PdInput : G.Input);
                ws.Rumus(r, 3, (string)kr[2], kk, rate ? G.Pd : G.Angka);
                ws.Rumus(r, 4, "=C" + r + "-B" + r, sel, rate ? G.Pd : G.Angka);
                ws.Rumus(r, 5, "=IF(ABS(D" + r + ")<" + (rate ? "0.000001" : "1") + ",\"sama\",\"berbeda\")", sama ? "sama" : "berbeda", sama ? G.Biasa : G.Peringatan);
                r++;
            }
            ws.Teks(r++, 1, "Status \"berbeda\" muncul bila angka di kertas kerja diubah (simulasi), atau bila versi menyimpan angka yang tidak berasal dari rincian " +
                            "(lihat catatan).", G.Catatan);
            if (beda > 0)
                CatatanLog.Tulis("Kertas kerja " + k.KodeKC + " v" + k.Versi + ": " + beda + " pos kontrol berbeda dengan database");

            // ---- Catatan rekonstruksi ----
            if (k.Catatan.Count > 0)
            {
                r++;
                ws.Teks(r++, 1, "Catatan", G.Subjudul);
                foreach (var c in k.Catatan) { ws.Teks(r, 1, "• " + c, G.Bungkus); ws.GabungSel(r, 1, r, 6); r++; }
            }

            // ---- Penyesuaian dalam versi ----
            var daftar = new List<string[]>();
            foreach (var o in L(k.Data["individu"]) ?? new ArrayList())
            {
                var d = D(o);
                if (Ambil(d, "disesuaikan") is bool && (bool)Ambil(d, "disesuaikan"))
                    daftar.Add(new[] { "CKPN Individu", S(Ambil(d, "kontrak")), S(Ambil(d, "nama")), "agunan/biaya disesuaikan" });
            }
            foreach (var o in L(k.Data["lgdcs"]) ?? new ArrayList())
            {
                var d = D(o);
                string sb = S(Ambil(d, "sumber"));
                if (sb.Length > 0 && sb != "sistem")
                    daftar.Add(new[] { "LGD CS", S(Ambil(d, "rek")), S(Ambil(d, "nama")), LabelSumber(sb) });
            }
            r++;
            ws.Teks(r++, 1, "Penyesuaian & koreksi dalam versi ini (" + daftar.Count + ")", G.Subjudul);
            if (daftar.Count == 0) ws.Teks(r++, 1, "Tidak ada — semua angka hasil perhitungan sistem.", G.Catatan);
            else
            {
                ws.Teks(r, 1, "Modul", G.Kepala); ws.Teks(r, 2, "Kontrak / rekening", G.Kepala); ws.Teks(r, 3, "Nama", G.Kepala); ws.Teks(r, 4, "Keterangan", G.Kepala);
                ws.GabungSel(r, 4, r, 6);
                r++;
                foreach (var p in daftar)
                {
                    ws.Teks(r, 1, p[0]); ws.Teks(r, 2, p[1], G.Teks); ws.Teks(r, 3, p[2]); ws.Teks(r, 4, p[3]);
                    r++;
                }
            }
        }

        // ================================================================
        // Util JSON & format
        // ================================================================
        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }
        private static Dictionary<string, object> D(object o) { return o as Dictionary<string, object>; }
        private static IList L(object o) { return o as IList; }
        private static string S(object o) { return o == null || o is DBNull ? "" : Convert.ToString(o, CultureInfo.InvariantCulture); }

        private static double? N(object o)
        {
            if (o == null || o is DBNull) return null;
            if (o is double) return (double)o;
            if (o is int) return (int)o;
            if (o is long) return (long)o;
            if (o is decimal) return (double)(decimal)o;
            if (o is float) return (float)o;
            double d;
            return double.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? (double?)d : null;
        }
        private static int? Bulat(object o)
        {
            double? v = N(o);
            return v.HasValue ? (int?)Convert.ToInt32(Math.Round(v.Value)) : null;
        }

        private static readonly CultureInfo Id = new CultureInfo("id-ID");
        private static string Rp(double v) { return "Rp" + v.ToString("#,##0", Id); }
        private static string Persen(double v) { return (v * 100).ToString("0.0000", Id) + "%"; }

        private static string TanggalPanjang(string iso)
        {
            DateTime t;
            return DateTime.TryParseExact(iso, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out t)
                ? t.ToString("d MMMM yyyy", Id) : iso;
        }
    }
}
