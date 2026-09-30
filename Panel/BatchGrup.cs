using System;
using System.Collections;
using System.Collections.Generic;
using CKPNLibrary.Data;
using CKPNLibrary.Helpers;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Hitung semua grup sekaligus (Tahap 4d) — dari tab Per grup.
    ///
    /// Grup dijalankan berurutan memakai jalur yang SAMA dengan tombol "Hitung grup ini":
    ///   CKPNPipeline.MulaiGrup (centang KC + Top-N grup ditulis ke Master, semua langkah,
    ///   penyesuaian tersimpan diterapkan). Setelah satu grup selesai, hasilnya diperiksa:
    ///
    ///   TEMUAN (perlu dicek manusia sebelum disimpan):
    ///     - kontrak CKPN Individu baru (belum pernah ada di staging)
    ///     - nilai agunan sistem berubah pada kontrak yang punya penyesuaian
    ///     - debitur LGD CS baru
    ///     - riwayat staging belum ada (kontrak/debitur baru tidak bisa dideteksi)
    ///     - masalah membaca penyesuaian, peringatan Simpan grup (mis. Top-N berbeda),
    ///       atau perubahan penyesuaian yang terdeteksi
    ///
    ///   Tanpa temuan  → grup langsung disimpan (versi baru, snapshot .xlsx seperti biasa).
    ///   Ada temuan    → grup TIDAK disimpan. Sesuai pilihan user:
    ///                   "berhenti"  : proses berhenti di grup itu; sheet masih berisi grup
    ///                                 tersebut sehingga bisa langsung direview & disimpan;
    ///                   "lanjut"    : grup ditandai perlu review, grup berikutnya tetap dihitung
    ///                                 (sheet akan tertimpa; grup bertanda dibuka lagi satu per satu).
    ///
    /// Hitung per grup satu per satu tidak berubah sama sekali.
    /// </summary>
    internal static class BatchGrup
    {
        public static bool Aktif { get; private set; }

        private class Item { public string Nama, KodeKC; public int? TopN; }

        private static List<Item> _antrian;
        private static int _indeks;
        private static bool _berhentiBilaTemuan, _mintaBatal;
        private static string _tanggal;
        private static List<Dictionary<string, object>> _hasil;
        private static System.Windows.Forms.Timer _timer;

        /// <param name="lewatiTersimpan">true = hanya grup yang belum punya versi tersimpan di periode ini</param>
        /// <param name="berhentiBilaTemuan">true = berhenti di grup pertama yang punya temuan</param>
        /// <returns>null bila dimulai; selain itu pesan penolakan</returns>
        public static string Mulai(Excel.Application app, bool lewatiTersimpan, bool berhentiBilaTemuan)
        {
            if (Aktif || CKPNPipeline.SedangBerjalan) return "Perhitungan lain sedang berjalan.";
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) return "Workbook aplikasi CKPN tidak sedang terbuka.";

            string info;
            if (!Database.BolehMenulis(out info)) return info;

            DateTime tgl;
            if (!ParameterMaster.BacaTanggalLaporan(wb, out tgl)) return "Tanggal laporan (Master!C4) belum diisi.";
            string tanggal = tgl.ToString("yyyy-MM-dd");
            Excel.Worksheet m = ParameterMaster.CariSheet(wb, "Master");
            string kodeMaster = m == null ? "" : string.Join(",", ParameterMaster.BacaKCDicentang(m).ToArray());

            var status = Periode.Status(tanggal, kodeMaster, tanggal);
            if (!(bool)status["susunanDitetapkan"]) return "Susunan grup tahun " + tgl.Year + " belum ditetapkan.";
            if (Convert.ToString(status["status"]) == "Final") return "Periode " + tanggal + " sudah dikunci (Final).";

            var antrian = new List<Item>();
            foreach (Dictionary<string, object> g in (IEnumerable)status["grup"])
            {
                string st = Convert.ToString(g["status"]);
                if (lewatiTersimpan && (st == "tersimpan" || st == "dihitung-ulang")) continue;
                antrian.Add(new Item
                {
                    Nama = Convert.ToString(g["nama"]), KodeKC = Convert.ToString(g["kodeKC"]),
                    TopN = g["topN"] == null ? (int?)null : Convert.ToInt32(g["topN"])
                });
            }
            if (antrian.Count == 0) return "Semua grup periode " + tanggal + " sudah tersimpan. Matikan pilihan \"lewati grup tersimpan\" untuk menghitung ulang semuanya.";

            _antrian = antrian;
            _indeks = 0;
            _tanggal = tanggal;
            _berhentiBilaTemuan = berhentiBilaTemuan;
            _mintaBatal = false;
            _hasil = new List<Dictionary<string, object>>();
            Aktif = true;

            var daftar = new List<object>();
            foreach (var it in antrian)
                daftar.Add(new Dictionary<string, object> { { "nama", it.Nama }, { "kodeKC", it.KodeKC }, { "topN", it.TopN } });
            CatatanLog.Tulis("=== HITUNG SEMUA GRUP · periode " + tanggal + " · " + antrian.Count + " grup · " +
                             (berhentiBilaTemuan ? "berhenti bila ada temuan" : "lanjut bila ada temuan"));
            PanelBridge.Siarkan("batchMulai", new Dictionary<string, object>
            {
                { "tanggal", tanggal }, { "grup", daftar }, { "berhentiBilaTemuan", berhentiBilaTemuan }
            });

            MulaiGrupBerikut();
            return null;
        }

        /// <summary>Tombol batal: grup yang sedang berjalan dihentikan, grup berikutnya tidak dimulai.</summary>
        public static void MintaBatal()
        {
            if (Aktif) _mintaBatal = true;
        }

        // ------------------------------------------------------------
        private static void MulaiGrupBerikut()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;
            if (_mintaBatal) { Selesai("dibatalkan", null); return; }
            if (_indeks >= _antrian.Count) { Selesai("selesai", null); return; }

            var it = _antrian[_indeks];
            Kabar(it, "menghitung", null, null);
            CKPNPipeline.SetelahSelesai = GrupSelesai;
            string tolak = CKPNPipeline.MulaiGrup(app, it.KodeKC, true);
            if (tolak != null)
            {
                CKPNPipeline.SetelahSelesai = null;
                Kabar(it, "gagal", new List<string> { tolak }, null);
                Selesai("gagal", it.Nama + ": " + tolak);
            }
        }

        /// <summary>Dipanggil CKPNPipeline di akhir run satu grup (konteks makro).</summary>
        private static void GrupSelesai(Excel.Application app, string statusRun)
        {
            var it = _antrian[_indeks];
            if (statusRun != "selesai")
            {
                Kabar(it, statusRun == "dibatalkan" ? "dibatalkan" : "gagal", null, null);
                Selesai(statusRun == "dibatalkan" ? "dibatalkan" : "gagal", statusRun == "dibatalkan" ? null : it.Nama + ": perhitungan gagal");
                return;
            }

            List<string> temuan;
            try { temuan = PeriksaTemuan(app); }
            catch (Exception ex) { temuan = new List<string> { "Pemeriksaan hasil gagal: " + ex.Message }; }

            if (temuan.Count > 0)
            {
                Kabar(it, "perlu-review", temuan, null);
                if (_berhentiBilaTemuan) { Selesai("perlu-review", null); return; }
            }
            else
            {
                try
                {
                    // Summary baru saja dihitung ulang oleh pipeline → tidak perlu Refresh Summary lagi
                    var r = StagingGrup.Simpan(app, null, "Hitung semua grup (otomatis, tanpa temuan)", false);
                    Kabar(it, "tersimpan", null, r);
                }
                catch (Exception ex)
                {
                    Kabar(it, "gagal", new List<string> { "Simpan gagal: " + ex.Message }, null);
                    Selesai("gagal", it.Nama + ": simpan gagal — " + ex.Message);
                    return;
                }
            }

            _indeks++;
            JadwalkanBerikut();
        }

        /// <summary>Daftar temuan yang membuat grup tidak boleh disimpan otomatis (kosong = aman).</summary>
        private static List<string> PeriksaTemuan(Excel.Application app)
        {
            var t = new List<string>();
            var rv = Penyesuaian.InfoReview();
            var ind = rv["individu"] as Dictionary<string, object>;
            var cs = rv["lgdcs"] as Dictionary<string, object>;

            if (ind == null) t.Add("Hasil review CKPN Individu tidak tersedia.");
            else
            {
                int baru = Jumlah(ind, "baru"), berubah = Jumlah(ind, "berubah");
                if (baru > 0) t.Add(baru + " kontrak CKPN Individu baru — cek nilai agunan (I) dan biaya penjualan (J).");
                if (berubah > 0) t.Add(berubah + " kontrak berpenyesuaian dengan nilai agunan sistem berubah.");
                if (!Bool(ind, "dikenalTersedia")) t.Add("Riwayat staging Individu belum ada — kontrak baru tidak dapat dideteksi.");
                if (ind.ContainsKey("masalah") && ind["masalah"] != null) t.Add("Individu: " + ind["masalah"]);
            }
            if (cs == null) t.Add("Hasil review LGD CS tidak tersedia.");
            else
            {
                int baru = Jumlah(cs, "baru");
                if (baru > 0) t.Add(baru + " debitur LGD CS baru — cek ke remedial.");
                if (!Bool(cs, "dikenalTersedia")) t.Add("Riwayat staging LGD CS belum ada — debitur baru tidak dapat dideteksi.");
                if (cs.ContainsKey("masalah") && cs["masalah"] != null) t.Add("LGD CS: " + cs["masalah"]);
            }

            // Pemeriksaan yang sama dengan kartu Simpan grup
            var p = StagingGrup.Periksa(app);
            foreach (var s in (IEnumerable)p["peringatan"]) t.Add(Convert.ToString(s));
            int nPerubahan = 0;
            foreach (var o in (IEnumerable)p["perubahan"]) nPerubahan++;
            if (nPerubahan > 0) t.Add(nPerubahan + " perubahan penyesuaian terdeteksi di sheet.");
            if (!(bool)p["bolehMenulis"] && t.Count == 0) t.Add("Grup ini tidak dapat disimpan (" + p["infoPengirim"] + ").");
            return t;
        }

        private static void JadwalkanBerikut()
        {
            // Jeda singkat seperti antarlangkah pipeline, agar pesan "batal" dari panel sempat diproses
            if (_timer == null)
            {
                _timer = new System.Windows.Forms.Timer { Interval = 400 };
                _timer.Tick += (s, e) => { _timer.Stop(); ExcelAsyncUtil.QueueAsMacro(MulaiGrupBerikut); };
            }
            _timer.Stop();
            _timer.Start();
        }

        private static void Kabar(Item it, string status, List<string> temuan, Dictionary<string, object> simpan)
        {
            var d = new Dictionary<string, object>
            {
                { "ke", _indeks + 1 }, { "dari", _antrian.Count }, { "nama", it.Nama }, { "kodeKC", it.KodeKC },
                { "status", status }, { "temuan", temuan ?? new List<string>() }
            };
            if (simpan != null)
            {
                d["versi"] = simpan["versi"];
                d["nfTotal"] = simpan["nfTotal"];
                d["migTotal"] = simpan["migTotal"];
            }
            if (status != "menghitung") _hasil.Add(d);
            CatatanLog.Tulis("  [semua grup] " + it.Nama + " · " + status + (temuan != null && temuan.Count > 0 ? " · " + string.Join(" | ", temuan.ToArray()) : ""));
            PanelBridge.Siarkan("batchGrup", d);
        }

        private static void Selesai(string status, string error)
        {
            Aktif = false;
            CKPNPipeline.SetelahSelesai = null;
            int sisa = _antrian == null ? 0 : Math.Max(0, _antrian.Count - _indeks - (status == "selesai" ? 0 : 1));
            CatatanLog.Tulis("=== HITUNG SEMUA GRUP " + status.ToUpperInvariant() + (error == null ? "" : " · " + error));
            PanelBridge.Siarkan("batchSelesai", new Dictionary<string, object>
            {
                { "status", status }, { "error", error }, { "hasil", _hasil }, { "sisa", sisa }, { "tanggal", _tanggal }
            });
        }

        private static int Jumlah(Dictionary<string, object> d, string k)
        {
            object v;
            if (!d.TryGetValue(k, out v) || v == null) return 0;
            var c = v as ICollection;
            return c == null ? 0 : c.Count;
        }

        private static bool Bool(Dictionary<string, object> d, string k)
        {
            object v;
            return d.TryGetValue(k, out v) && v is bool && (bool)v;
        }
    }
}
