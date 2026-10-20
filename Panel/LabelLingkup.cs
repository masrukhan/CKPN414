using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Label lingkup kertas kerja (Tahap 7h).
    ///
    /// Sheet Summary, "A. CKPN - INDV", dan "B. CKPN - KOL INDV" hanya berisi SATU grup — grup yang terakhir
    /// dihitung — sedangkan panel menampilkan seluruh grup. Modul ini menulis keterangan di kertas kerja:
    ///   Summary A2 (menggantikan subjudul) · INDV B4 · KOL INDV B1
    /// berisi nama grup, periode, dan status terhadap database (belum disimpan / sama dengan vN /
    /// database sudah vM). Judul Summary bagian E dan judul tabel PPKA juga diberi keterangan "grup ini".
    ///
    /// Status diambil dari konteks lokal (Periode.BacaKonteks: grup terakhir dihitung/disimpan di PC ini)
    /// dibandingkan versi aktif di database. Label ditulis ulang hanya bila teksnya berubah.
    /// </summary>
    internal static class LabelLingkup
    {
        private const string SheetSum  = "Summary";
        private const string SheetIndv = "A. CKPN - INDV";
        private const string SheetKol  = "B. CKPN - KOL INDV";

        // Warna latar (Excel: R + G*256 + B*65536)
        private static int Rgb(int r, int g, int b) { return r + g * 256 + b * 65536; }
        private static readonly int WarnaKuning = Rgb(255, 242, 204);
        private static readonly int WarnaHijau  = Rgb(226, 239, 218);
        private static readonly int WarnaAbu    = Rgb(242, 242, 242);
        private static readonly int WarnaTeks   = Rgb(64, 64, 64);

        // Status singkat → dipakai panel juga
        public const string SBelum     = "belum";        // grup Master belum pernah dihitung di PC ini
        public const string SDihitung  = "dihitung";     // dihitung, belum disimpan
        public const string SManual    = "manual";       // dihitung dari tombol Master, belum disimpan
        public const string SSama      = "sama";         // sama dengan versi aktif database
        public const string SBeda      = "beda";         // database sudah versi lain (koreksi / PC lain)

        // =================================================================
        // Info: grup di workbook + status terhadap database
        // =================================================================
        public static Dictionary<string, object> Info(string tanggalMaster, string kodeKCMaster)
        {
            var h = new Dictionary<string, object>
            {
                { "tanggal", tanggalMaster }, { "kodeKC", kodeKCMaster }, { "nama", null },
                { "status", SBelum }, { "versiSheet", null }, { "versiDb", null }, { "koreksiDb", false },
                { "waktuHitung", null }, { "label", "" }
            };
            if (string.IsNullOrEmpty(tanggalMaster) || string.IsNullOrEmpty(kodeKCMaster)) return h;

            DateTime tgl;
            if (!DateTime.TryParseExact(tanggalMaster, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out tgl))
                return h;

            string nama;
            Periode.TopNGrup(tgl.Year, kodeKCMaster, out nama);
            h["nama"] = nama;

            // ---- versi aktif di database ----
            int? versiDb = null; bool koreksiDb = false;
            try
            {
                if (Database.Ada)
                    using (var con = Database.Buka(false))
                    {
                        var aktif = Periode.DaftarRun(con, tanggalMaster).Find(r =>
                            string.Equals((string)r["kodeKC"], kodeKCMaster, StringComparison.OrdinalIgnoreCase) &&
                            (bool)r["aktif"] && !(bool)r["dihapus"]);
                        if (aktif != null)
                        {
                            versiDb = (int)aktif["versi"];
                            var rg = aktif["ringkasan"] as Dictionary<string, object>;
                            koreksiDb = rg != null && rg.ContainsKey("koreksi_dari_versi");
                        }
                    }
            }
            catch (Exception ex) { CatatanLog.Tulis("Label lingkup: baca versi gagal: " + ex.Message); }
            h["versiDb"] = versiDb;
            h["koreksiDb"] = koreksiDb;

            // ---- konteks lokal: apa yang terakhir terjadi pada sheet di PC ini ----
            var k = Periode.BacaKonteks();
            bool cocok = k != null &&
                         Convert.ToString(Ambil(k, "tanggal")) == tanggalMaster &&
                         string.Equals(Convert.ToString(Ambil(k, "kodeKC")), kodeKCMaster, StringComparison.OrdinalIgnoreCase);
            if (cocok)
            {
                object vs = Ambil(k, "tersimpanVersi");
                int? versiSheet = vs == null ? (int?)null : Convert.ToInt32(vs);
                h["versiSheet"] = versiSheet;
                h["waktuHitung"] = Ambil(k, "waktuHitung");
                bool manual = Convert.ToString(Ambil(k, "sumber")) == "manual";

                if (versiSheet == null)
                    h["status"] = Ambil(k, "waktuHitung") == null ? SBelum : (manual ? SManual : SDihitung);
                else
                    h["status"] = versiDb.HasValue && versiDb.Value == versiSheet.Value ? SSama : SBeda;
            }
            else if (versiDb.HasValue)
            {
                // Tidak tercatat di PC ini (mis. file disalin dari komputer lain)
                h["status"] = SBelum;
            }

            h["lite"] = EdisiLite();
            h["menyeluruh"] = SeluruhSegmen(kodeKCMaster);
            h["label"] = TeksStatus(h);
            return h;
        }

        /// <summary>Kalimat status (dipakai Excel dan panel).</summary>
        public static string TeksStatus(Dictionary<string, object> h)
        {
            object vd = Ambil(h, "versiDb"), vs = Ambil(h, "versiSheet");
            string s0 = Convert.ToString(Ambil(h, "status"));
            if (Ambil(h, "lite") is bool && (bool)Ambil(h, "lite") && (s0 == SDihitung || s0 == SManual))
                return "Edisi Lite: hasil perhitungan terakhir, tidak disimpan ke database.";
            switch (s0)
            {
                case SDihitung: return "Belum disimpan ke database" + (vd != null ? " (database masih v" + vd + ")" : "") + ".";
                case SManual:   return "Dihitung dari tombol Master (manual), belum disimpan ke database.";
                case SSama:     return "Sama dengan versi tersimpan v" + vs + ".";
                case SBeda:
                    return "Database sudah v" + vd + ((bool)Ambil(h, "koreksiDb") ? " (hasil koreksi dari database)" : "") +
                           "; angka sheet ini v" + vs + ". Angka yang berlaku: versi database.";
                default:
                    return vd != null
                        ? "Isi sheet tidak tercatat di komputer ini. Database: v" + vd + "."
                        : "Grup ini belum dihitung untuk periode ini.";
            }
        }

        // =================================================================
        // Tulis label ke workbook (dipanggil di konteks makro Excel)
        // =================================================================
        public static void Perbarui(Excel.Workbook wb)
        {
            try
            {
                if (wb == null || CKPNPipeline.SedangBerjalan) return;
                DateTime t;
                if (!ParameterMaster.BacaTanggalLaporan(wb, out t)) return;
                Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
                if (m == null) return;
                string kode = string.Join(",", ParameterMaster.BacaKCDicentang(m).ToArray());
                Tulis(wb, Info(t.ToString("yyyy-MM-dd"), kode));
            }
            catch (Exception ex) { CatatanLog.Tulis("Label lingkup: " + ex.Message); }
        }

        public static void Tulis(Excel.Workbook wb, Dictionary<string, object> info)
        {
            if (wb == null || info == null || string.IsNullOrEmpty(Convert.ToString(Ambil(info, "kodeKC")))) return;

            string status = Convert.ToString(Ambil(info, "status"));
            int warna = status == SSama ? WarnaHijau : (status == SBelum ? WarnaAbu : WarnaKuning);
            string grup = JudulGrup(info);
            string periode = Bulan(Convert.ToString(Ambil(info, "tanggal")));
            string st = Convert.ToString(Ambil(info, "label"));

            bool lite = Ambil(info, "lite") is bool && (bool)Ambil(info, "lite");
            bool menyeluruh = Ambil(info, "menyeluruh") is bool && (bool)Ambil(info, "menyeluruh");
            if (menyeluruh && string.IsNullOrEmpty(Convert.ToString(Ambil(info, "nama"))))
                grup = "Seluruh segmen (KC0600–KC1100)";

            // Baris ketiga Summary & catatan bagian E — tergantung cakupan dan edisi
            string rujukan, judulPpka, judulE, teksCatatan;
            if (menyeluruh)
            {
                rujukan = lite
                    ? "Edisi Lite: Summary ini berisi hasil seluruh segmen sekaligus."
                    : "Summary berisi seluruh segmen; angka resmi tetap versi tersimpan di Panel › Tetapkan › Konsolidasi.";
                judulPpka = "PPKA per KC (seluruh segmen + ABA)";
                judulE = "E. Perbandingan PPKA vs CKPN — seluruh segmen";
                teksCatatan = lite
                    ? "Catatan: edisi Lite menghitung seluruh segmen KC0600–KC1100 sekaligus. Penyimpanan versi, konsolidasi " +
                      "di panel, penetapan periode, dan dokumen tersedia di edisi Pro."
                    : "Catatan: bagian E mencakup seluruh segmen KC0600–KC1100. Usulan jurnal resmi ada di " +
                      "Panel › Tetapkan › Konsolidasi vs PPKA OJK.";
            }
            else
            {
                rujukan = lite
                    ? "Edisi Lite: ini hasil lama satu grup. Hitung ulang untuk mendapatkan hasil seluruh segmen."
                    : "Total seluruh grup & usulan jurnal: Panel › Tetapkan › Konsolidasi vs PPKA OJK.";
                judulPpka = "PPKA per KC (KC grup ini + ABA)";
                judulE = "E. Perbandingan PPKA vs CKPN — HANYA GRUP INI (bukan dasar jurnal)";
                teksCatatan =
                    "Catatan: bagian E hanya menghitung grup di atas. CKPN ABA (C26) ikut di Summary setiap grup, " +
                    "jadi jangan dijumlahkan antargrup. Total seluruh grup dan usulan jurnal ada di " +
                    "Panel › Tetapkan › Konsolidasi vs PPKA OJK.";
            }

            string teksSum = (menyeluruh ? "" : "HANYA GRUP: ") + grup + " · " + periode + "\n" + st + "\n" + rujukan;
            string teksBaris = (menyeluruh ? "" : "Grup: ") + grup + " · " + periode + " — " + st;

            Pasang(wb, SheetSum,
                ws => Teks(ws, "A2") == teksSum && Warna(ws, "A2") == warna && Teks(ws, "A28") == judulE &&
                      Teks(ws, "G1") == judulPpka && Teks(ws, "A38") == teksCatatan,
                ws =>
                {
                    Excel.Range a2 = (Excel.Range)ws.Range["A2"];
                    a2.Value2 = teksSum;
                    Excel.Range blok = (Excel.Range)ws.Range["A2", "C2"];
                    blok.Interior.Color = warna;
                    blok.WrapText = true;
                    blok.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    blok.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    a2.Font.Name = "Calibri"; a2.Font.Size = 9; a2.Font.Italic = false; a2.Font.Color = WarnaTeks;
                    ((Excel.Range)ws.Rows[2]).RowHeight = 50;

                    // Judul tabel PPKA & bagian E: tegaskan "grup ini"
                    ((Excel.Range)ws.Range["G1"]).Value2 = judulPpka;
                    ((Excel.Range)ws.Range["A28"]).Value2 = judulE;
                    Excel.Range catatan = (Excel.Range)ws.Range["A38"];
                    catatan.Value2 = teksCatatan;
                    catatan.Font.Italic = true; catatan.Font.Size = 9; catatan.Font.Color = WarnaTeks;
                });

            Pasang(wb, SheetIndv, ws => Teks(ws, "B4") == teksBaris && Warna(ws, "B4") == warna,
                                  ws => TulisBaris(ws, "B4", "K4", teksBaris, warna));
            Pasang(wb, SheetKol,  ws => Teks(ws, "B1") == teksBaris && Warna(ws, "B1") == warna,
                                  ws => TulisBaris(ws, "B1", "G1", teksBaris, warna));
        }

        private static void TulisBaris(Excel.Worksheet ws, string sel, string selAkhir, string teks, int warna)
        {
            Excel.Range c = (Excel.Range)ws.Range[sel];
            c.Value2 = teks;
            Excel.Range blok = (Excel.Range)ws.Range[sel, selAkhir];
            blok.Interior.Color = warna;
            c.WrapText = false;                 // teks meluber ke sel kosong di kanan
            c.Font.Name = "Calibri"; c.Font.Size = 9; c.Font.Bold = true; c.Font.Color = WarnaTeks;
            c.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }

        /// <summary>Proteksi sheet hanya dibuka bila isi label memang berubah.</summary>
        private static void Pasang(Excel.Workbook wb, string sheet,
                                   Func<Excel.Worksheet, bool> sudahSama, Action<Excel.Worksheet> tulis)
        {
            try
            {
                Excel.Worksheet ws = ParameterMaster.CariSheet(wb, sheet);
                if (ws == null || sudahSama(ws)) return;
                Protection.UbahSheet(wb, sheet, tulis);
            }
            catch (Exception ex) { CatatanLog.Tulis("Label lingkup " + sheet + ": " + ex.Message); }
        }

        private static string Teks(Excel.Worksheet ws, string sel)
        {
            try { return Convert.ToString(((Excel.Range)ws.Range[sel]).Value2) ?? ""; } catch { return ""; }
        }

        private static int Warna(Excel.Worksheet ws, string sel)
        {
            try { return Convert.ToInt32(((Excel.Range)ws.Range[sel]).Interior.Color); } catch { return -1; }
        }

        // =================================================================
        // Tombol VBA di Master (hitung manual) — dipanggil dari CKPNFunctions
        // =================================================================
        /// <summary>
        /// Perhitungan dari tombol Master mengubah isi sheet tanpa lewat panel. Catat konteks "manual"
        /// (status grup → dihitung, belum disimpan) dan tulis label. Tidak berlaku saat panel yang menjalankan.
        /// </summary>
        public static void TandaiManual(Excel.Application app, string sheetKCList)
        {
            try
            {
                if (CKPNPipeline.SedangBerjalan || BatchGrup.Aktif) return;
                Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app) ?? app.ActiveWorkbook;
                DateTime t;
                if (wb == null || !ParameterMaster.BacaTanggalLaporan(wb, out t)) return;
                string kode = string.IsNullOrEmpty(sheetKCList) ? "" : string.Join(",", sheetKCList.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries));
                if (kode.Length == 0) return;
                Periode.CatatHitung(t.ToString("yyyy-MM-dd"), kode, new[] { "manual" }, "manual");
                Tulis(wb, Info(t.ToString("yyyy-MM-dd"), kode));
            }
            catch (Exception ex) { CatatanLog.Tulis("Label lingkup (manual): " + ex.Message); }
        }

        // =================================================================
        // Pencocokan Summary Excel vs versi aktif database
        // =================================================================
        private static readonly string[,] SelCocok =
        {
            { "B6",  "nf_individu",  "Net Flow · Individu" },
            { "B7",  "nf_kolektif",  "Net Flow · Kolektif" },
            { "B8",  "nf_total",     "Net Flow · Total" },
            { "B13", "mig_individu", "Migration · Individu" },
            { "B14", "mig_kolektif", "Migration · Kolektif" },
            { "B15", "mig_total",    "Migration · Total" },
            { "C26", "aba_ckpn",     "CKPN ABA" },
            { "H9",  "ppka_total",   "Total PPKA (Summary H9)" }
        };

        public static Dictionary<string, object> Cocokkan(Excel.Workbook wb)
        {
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
            DateTime t;
            if (!ParameterMaster.BacaTanggalLaporan(wb, out t)) throw new InvalidOperationException("Tanggal laporan (Master!C4) belum diisi.");
            string tanggal = t.ToString("yyyy-MM-dd");
            Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
            string kode = m == null ? "" : string.Join(",", ParameterMaster.BacaKCDicentang(m).ToArray());
            var info = Info(tanggal, kode);

            Excel.Worksheet ws = ParameterMaster.CariSheet(wb, SheetSum);
            if (ws == null) throw new InvalidOperationException("Sheet Summary tidak ditemukan.");
            try { ws.Calculate(); } catch { }

            Dictionary<string, object> rg = null;
            if (Database.Ada)
                using (var con = Database.Buka(false))
                {
                    var aktif = Periode.DaftarRun(con, tanggal).Find(r =>
                        string.Equals((string)r["kodeKC"], kode, StringComparison.OrdinalIgnoreCase) &&
                        (bool)r["aktif"] && !(bool)r["dihapus"]);
                    if (aktif != null) rg = aktif["ringkasan"] as Dictionary<string, object>;
                }

            var baris = new List<object>();
            int beda = 0;
            for (int i = 0; i < SelCocok.GetLength(0); i++)
            {
                double? ex = Angka(((Excel.Range)ws.Range[SelCocok[i, 0]]).Value2);
                double? db = null;
                object v;
                if (rg != null && rg.TryGetValue(SelCocok[i, 1], out v) && v != null) db = Convert.ToDouble(v);
                bool sama = ex.HasValue && db.HasValue && Math.Abs(ex.Value - db.Value) < 1;
                if (rg != null && !sama) beda++;
                baris.Add(new Dictionary<string, object>
                {
                    { "label", SelCocok[i, 2] }, { "sel", SelCocok[i, 0] },
                    { "excel", ex }, { "database", db }, { "sama", sama }
                });
            }
            return new Dictionary<string, object>
            {
                { "info", info }, { "adaVersi", rg != null }, { "baris", baris }, { "jumlahBeda", beda }
            };
        }

        // =================================================================
        // util
        // =================================================================
        private static readonly string[] SemuaKC = { "KC0600", "KC0700", "KC0800", "KC0900", "KC1000", "KC1100" };

        private static bool SeluruhSegmen(string kodeKC)
        {
            var ada = new HashSet<string>((kodeKC ?? "").Split(','), StringComparer.OrdinalIgnoreCase);
            foreach (var k in SemuaKC) if (!ada.Contains(k)) return false;
            return true;
        }

        private static bool EdisiLite()
        {
            try { return Lisensi.Status().Edisi == Lisensi.Lite; } catch { return false; }
        }

        private static string JudulGrup(Dictionary<string, object> info)
        {
            string nama = Convert.ToString(Ambil(info, "nama"));
            string kode = Convert.ToString(Ambil(info, "kodeKC"));
            return string.IsNullOrEmpty(nama) ? kode : nama + " (" + kode + ")";
        }

        private static string Bulan(string tanggal)
        {
            DateTime t;
            return DateTime.TryParseExact(tanggal, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out t)
                ? t.ToString("MMMM yyyy", LogProses.Id) : tanggal;
        }

        private static double? Angka(object v)
        {
            if (v is double) return (double)v;
            double d;
            return v != null && double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.InvariantCulture, out d) ? d : (double?)null;
        }

        private static object Ambil(Dictionary<string, object> d, string k)
        {
            object v; return d != null && d.TryGetValue(k, out v) ? v : null;
        }
    }
}
