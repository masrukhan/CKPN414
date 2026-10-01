using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using CKPNLibrary.Helpers;
using CKPNLibrary.Modules;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Panel
{
    /// <summary>
    /// Menjalankan rangkaian perhitungan CKPN dari Panel dengan progres per langkah.
    ///
    /// CARA KERJA
    ///   Setiap langkah dijadwalkan terpisah. Di antara dua langkah ada JEDA
    ///   singkat (timer ~250 ms) sebelum langkah berikutnya diantrekan lewat
    ///   ExcelAsyncUtil.QueueAsMacro. Selama jeda itu Excel benar-benar idle:
    ///   antrean pesan Windows dikosongkan, termasuk klik "Batalkan" dari
    ///   panel yang tertahan selama modul berjalan. Tanpa jeda ini, antrean
    ///   makro berikutnya bisa dieksekusi LEBIH DULU daripada pesan batal,
    ///   sehingga pembatalan saat PD Net Flow tidak pernah terbaca.
    ///   Modul perhitungan sendiri tetap berjalan sinkron seperti dari VBA.
    ///
    /// TAMPILAN EXCEL SAAT MODUL BERJALAN
    ///   Modul panjang (PD Net Flow membuka 12–13 file) memblok thread utama
    ///   lebih dari 5 detik. Windows lalu "ghosting": jendela Excel diputihkan
    ///   dan berjudul "(Not Responding)", walaupun proses berjalan normal.
    ///   DisableProcessWindowsGhosting mematikan efek itu untuk proses Excel,
    ///   sehingga tampilan tetap seperti saat CKPN Individu berjalan.
    ///
    /// LOGIKA HITUNG TIDAK DIDUPLIKASI
    ///   Pipeline memanggil kelas modul yang SAMA dengan CKPNFunctions
    ///   (CKPNIndividu, CKPNPDNetFlow, PDMigration, LGD..., RefreshSummary),
    ///   dengan parameter yang dibaca ParameterMaster persis seperti makro VBA.
    ///
    /// WORKBOOK AKTIF
    ///   Modul lama memakai _app.ActiveWorkbook. Sebelum setiap langkah, workbook
    ///   aplikasi diaktifkan ulang, sehingga modul selalu menulis ke workbook
    ///   yang benar walaupun user sempat mengklik workbook lain.
    /// </summary>
    internal static class CKPNPipeline
    {
        // ------------------------------------------------------------
        // Definisi langkah (urutan = urutan eksekusi)
        // ------------------------------------------------------------
        internal class Langkah
        {
            public string Id, Nama, SheetDitulis;
            public Action<Excel.Application, ParameterCKPN> Jalankan;
        }

        public static readonly Langkah[] SemuaLangkah =
        {
            new Langkah { Id = "individu",  Nama = "CKPN Individu",            SheetDitulis = "A. CKPN - INDV",            Jalankan = JalankanIndividu  },
            new Langkah { Id = "netflow",   Nama = "PD Net Flow",              SheetDitulis = "B1.PD-Net Flow",            Jalankan = JalankanNetFlow   },
            new Langkah { Id = "migration", Nama = "PD Migration",             SheetDitulis = "B2.PD-Migration, state PD Migration", Jalankan = JalankanMigration },
            new Langkah { Id = "lgder",     Nama = "LGD Expected Recoveries",  SheetDitulis = "B3.LGD-ER",                 Jalankan = JalankanLgdEr     },
            new Langkah { Id = "lgdcs",     Nama = "LGD Collateral Shortfall", SheetDitulis = "B4.LGD-CS MACET",           Jalankan = JalankanLgdCs     },
            new Langkah { Id = "summary",   Nama = "Refresh Summary & PPKA",   SheetDitulis = "Summary",                   Jalankan = JalankanSummary   },
        };

        /// <summary>Id langkah hitung (dipakai ParameterMaster untuk error umum).</summary>
        public static IEnumerable<string> IdLangkahHitung
        {
            get { foreach (var l in SemuaLangkah) yield return l.Id; }
        }

        // ------------------------------------------------------------
        // State run (hanya diakses di thread utama Excel)
        // ------------------------------------------------------------
        public static bool SedangBerjalan { get; private set; }

        /// <summary>
        /// Callback sekali pakai setelah run berakhir (dipakai BatchGrup untuk lanjut ke grup
        /// berikutnya). Null untuk run biasa, sehingga hitung manual/per grup tidak berubah.
        /// </summary>
        internal static Action<Excel.Application, string> SetelahSelesai;

        /// <summary>True hanya saat sebuah langkah sedang dieksekusi oleh pipeline.</summary>
        public static bool DiDalamLangkah { get; private set; }

        private static bool           _mintaBatal;
        private static List<Langkah>  _antrian;
        private static int            _indeks;
        private static ParameterCKPN  _param;
        private static Excel.Workbook _wbApp;
        private static Stopwatch      _swTotal;
        private static List<Dictionary<string, object>> _hasilLangkah;
        private static string         _langkahAktif = "";

        /// <summary>
        /// Dipanggil Protection.CekProteksi: tolak perintah dari tombol VBA saat
        /// pipeline panel sedang berjalan (di sela dua langkah), supaya dua
        /// perhitungan tidak menulis ke sheet yang sama secara bersamaan.
        /// </summary>
        public static bool TolakPanggilanLuar()
        {
            return SedangBerjalan && !DiDalamLangkah;
        }

        public static object Status()
        {
            return new Dictionary<string, object>
            {
                { "berjalan", SedangBerjalan },
                { "langkahAktif", _langkahAktif },
                { "ke", _indeks + 1 },
                { "dari", _antrian == null ? 0 : _antrian.Count }
            };
        }

        public static void MintaBatal()
        {
            if (SedangBerjalan) _mintaBatal = true;
        }

        // ------------------------------------------------------------
        // Mulai — dipanggil di konteks makro (QueueAsMacro dari PanelBridge)
        // ------------------------------------------------------------
        /// <returns>null bila berhasil dimulai; selain itu pesan penolakan.</returns>
        /// <param name="sumberLog">label sumber di log proses, mis. "Panel · Grup Murabahah" (bawaan "Panel · Manual")</param>
        public static string Mulai(Excel.Application app, IList<string> idDipilih, bool terapkanPenyesuaian = true,
                                   string sumberLog = null)
        {
            if (SedangBerjalan) return "Perhitungan lain sedang berjalan.";

            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) return "Workbook aplikasi CKPN tidak sedang terbuka.";

            // Baca ulang Master tepat sebelum jalan — nilai yang ditampilkan panel
            // bisa sudah basi bila user mengubah Master setelah menekan "Periksa".
            ParameterCKPN p = ParameterMaster.Baca(wb);

            var pilih = new HashSet<string>(idDipilih ?? new string[0], StringComparer.OrdinalIgnoreCase);
            var antrian = new List<Langkah>();
            var masalah = new List<string>();
            foreach (var l in SemuaLangkah)
            {
                if (!pilih.Contains(l.Id)) continue;
                if (!p.Siap(l.Id)) masalah.Add(l.Nama + ": " + string.Join("; ", p.Error[l.Id].ToArray()));
                antrian.Add(l);
            }
            if (antrian.Count == 0) return "Belum ada langkah yang dipilih.";
            if (masalah.Count > 0) return "Parameter belum lengkap:\n- " + string.Join("\n- ", masalah.ToArray());

            _wbApp        = wb;
            _param        = p;
            _antrian      = antrian;
            _indeks       = -1;          // -1 = langkah pemeriksaan proteksi
            _mintaBatal   = false;
            _swTotal      = Stopwatch.StartNew();
            _hasilLangkah = new List<Dictionary<string, object>>();
            SedangBerjalan = true;
            Pemberitahu.MulaiModePanel();
            // Penyesuaian manual tersimpan (Individu I/J, LGD CS) diterapkan di
            // dalam modul. Bisa dimatikan per run untuk melihat hasil sistem murni.
            Data.Penyesuaian.Aktif = terapkanPenyesuaian;
            MatikanGhostingSekali();
            try { app.Cursor = Excel.XlMousePointer.xlWait; } catch { }

            var daftar = new List<object>();
            foreach (var l in antrian)
                daftar.Add(new Dictionary<string, object> { { "id", l.Id }, { "nama", l.Nama } });

            CatatanLog.Tulis("=== RUN MULAI oleh " + Environment.UserName + " · periode " + p.BulanLaporan +
                             " · KC " + p.KCList + " · langkah: " + string.Join(",", Ids(antrian)));

            // Log proses (Tahap 5): semua catatan modul selama run ini memakai id jalan yang sama
            DateTime tglLog;
            string periodeLog = ParameterMaster.BacaTanggalLaporan(wb, out tglLog) ? tglLog.ToString("yyyy-MM-dd") : "";
            LogProses.MulaiJalan(string.IsNullOrEmpty(sumberLog) ? "Panel · Manual" : sumberLog, periodeLog);
            var namaLangkah = new List<string>();
            foreach (var l in antrian) namaLangkah.Add(l.Nama);
            LogProses.Catat("Panel", "Jalan mulai", LogProses.Info, LogProses.R()
                .Tambah("KC", p.KCList)
                .Tambah("Bulan laporan", p.BulanLaporan)
                .Tambah("Langkah", string.Join(", ", namaLangkah.ToArray()))
                .Tambah("Penyesuaian tersimpan", terapkanPenyesuaian ? "diterapkan" : "tidak diterapkan (hasil sistem murni)")
                .Tambah("Top-N Master!C10", ParameterMaster.BacaTopN(wb)));
            PanelBridge.Siarkan("runMulai", new Dictionary<string, object>
            {
                { "langkah", daftar }, { "periode", p.BulanLaporan }, { "kc", p.KCList },
                { "penyesuaian", terapkanPenyesuaian }
            });

            JadwalkanBerikut();
            return null;
        }

        // ------------------------------------------------------------
        // Eksekusi satu langkah, lalu jadwalkan langkah berikutnya
        // ------------------------------------------------------------
        private static void LangkahBerikut()
        {
            var app = (Excel.Application)ExcelDnaUtil.Application;

            if (_mintaBatal) { Akhiri(app, "dibatalkan", null); return; }

            // ---- Langkah 0: verifikasi proteksi file (sekali per run) ----
            if (_indeks == -1)
            {
                _langkahAktif = "proteksi";
                bool ok;
                DiDalamLangkah = true;
                try
                {
                    AktifkanWorkbookAplikasi();
                    ok = Protection.CekProteksi(app);   // menampilkan dialog sendiri bila gagal
                }
                catch (Exception ex) { Akhiri(app, "gagal", "Pemeriksaan proteksi: " + ex.Message); return; }
                finally { DiDalamLangkah = false; }

                if (!ok) { Akhiri(app, "gagal", "Pemeriksaan proteksi file gagal (lihat dialog yang muncul)."); return; }
                _indeks = 0;
                JadwalkanBerikut();
                return;
            }

            if (_indeks >= _antrian.Count) { Akhiri(app, "selesai", null); return; }

            Langkah l = _antrian[_indeks];
            _langkahAktif = l.Id;
            PanelBridge.Siarkan("langkahMulai", new Dictionary<string, object>
            {
                { "id", l.Id }, { "nama", l.Nama }, { "ke", _indeks + 1 }, { "dari", _antrian.Count }
            });
            try { app.StatusBar = "Panel CKPN: " + l.Nama + " (" + (_indeks + 1) + "/" + _antrian.Count + ")…"; } catch { }

            var sw = Stopwatch.StartNew();
            string error = null;

            // State aplikasi sama seperti makro VBA: layar & event dimatikan
            // selama modul berjalan, lalu dikembalikan ke kondisi semula.
            bool su = true, ev = true, da = true;
            Excel.XlCalculation calc = Excel.XlCalculation.xlCalculationAutomatic;
            try { su = app.ScreenUpdating; ev = app.EnableEvents; da = app.DisplayAlerts; calc = app.Calculation; } catch { }

            DiDalamLangkah = true;
            try
            {
                AktifkanWorkbookAplikasi();
                try
                {
                    app.ScreenUpdating = false;
                    app.Calculation    = Excel.XlCalculation.xlCalculationManual;
                    app.DisplayAlerts  = false;
                    app.EnableEvents   = false;
                }
                catch { }

                l.Jalankan(app, _param);
            }
            catch (Exception ex)
            {
                error = PesanError(ex);
            }
            finally
            {
                DiDalamLangkah = false;
                try { app.EnableEvents = ev; }   catch { }
                try { app.DisplayAlerts = da; }  catch { }
                try { app.Calculation = calc; }  catch { }
                try { app.ScreenUpdating = su; } catch { }
            }
            sw.Stop();

            var hasil = new Dictionary<string, object>
            {
                { "id", l.Id }, { "nama", l.Nama }, { "durasiMs", sw.ElapsedMilliseconds },
                { "ok", error == null }, { "error", error }
            };
            _hasilLangkah.Add(hasil);
            PanelBridge.Siarkan("langkahSelesai", hasil);
            CatatanLog.Tulis("  " + l.Nama + " · " + (error == null ? "OK" : "GAGAL: " + error) +
                             " · " + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + " dtk");
            LogProses.Catat("Panel", "Langkah " + l.Nama, error == null ? LogProses.OK : LogProses.Gagal, LogProses.R()
                .Tambah("Durasi", (sw.ElapsedMilliseconds / 1000.0).ToString("0.0", LogProses.Id) + " dtk")
                .TambahBila(error != null, "Error", error));

            if (error != null) { Akhiri(app, "gagal", l.Nama + ": " + error); return; }

            _indeks++;
            JadwalkanBerikut();
        }

        /// <summary>
        /// Hitung satu grup dari tab Periode: centang KC di Master diubah sesuai
        /// grup, lalu seluruh langkah dijalankan (penyesuaian tersimpan diterapkan).
        /// </summary>
        public static string MulaiGrup(Excel.Application app, string kodeKC, bool terapkanPenyesuaian)
        {
            if (SedangBerjalan) return "Perhitungan lain sedang berjalan.";
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) return "Workbook aplikasi CKPN tidak sedang terbuka.";
            try { ParameterMaster.SetKCDicentang(wb, kodeKC.Split(',')); }
            catch (Exception ex) { return "Tidak dapat mengubah centang KC di Master: " + PesanError(ex); }

            // Top-N debitur Individu mengikuti ketetapan grup di susunan tahunan (SOP) — ditulis ke Master!C10
            DateTime tgl;
            if (ParameterMaster.BacaTanggalLaporan(wb, out tgl))
            {
                string namaGrup;
                int? topN = Data.Periode.TopNGrup(tgl.Year, kodeKC, out namaGrup);
                if (topN.HasValue)
                {
                    try { ParameterMaster.SetTopN(wb, topN.Value); }
                    catch (Exception ex) { return "Tidak dapat menulis Top-N ke Master!C10: " + PesanError(ex); }
                }
            }

            string grupLog = kodeKC;
            DateTime tglG;
            if (ParameterMaster.BacaTanggalLaporan(wb, out tglG))
            {
                string nm;
                Data.Periode.TopNGrup(tglG.Year, kodeKC, out nm);
                if (!string.IsNullOrEmpty(nm)) grupLog = nm;
            }
            var semua = new List<string>(IdLangkahHitung);
            return Mulai(app, semua, terapkanPenyesuaian,
                         (BatchGrup.Aktif ? "Panel · Semua grup · " : "Panel · Grup ") + grupLog);
        }

        private static void Akhiri(Excel.Application app, string status, string error)
        {
            // Catat grup yang kini ada di workbook (dipakai tab Periode untuk status grup)
            if (status == "selesai" && _wbApp != null && _param != null)
            {
                try
                {
                    DateTime tgl;
                    if (ParameterMaster.BacaTanggalLaporan(_wbApp, out tgl))
                        Data.Periode.CatatHitung(tgl.ToString("yyyy-MM-dd"), _param.KCList, Ids(_antrian));
                }
                catch (Exception ex) { CatatanLog.Tulis("Catat konteks hitung gagal: " + ex.Message); }
            }

            var pesanModul = Pemberitahu.SelesaiModePanel();
            long durasi = _swTotal == null ? 0 : _swTotal.ElapsedMilliseconds;

            SedangBerjalan = false;
            DiDalamLangkah = false;
            Data.Penyesuaian.Aktif = true;   // tombol VBA selalu menerapkan penyesuaian
            _langkahAktif  = "";
            try { app.StatusBar = false; } catch { }
            try { app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            if (_timer != null) { try { _timer.Stop(); } catch { } }

            CatatanLog.Tulis("=== RUN " + status.ToUpperInvariant() + " · " + (durasi / 1000.0).ToString("0.0") + " dtk" +
                             (error == null ? "" : " · " + error));
            LogProses.Catat("Panel", "Jalan " + status,
                status == "selesai" ? LogProses.OK : status == "dibatalkan" ? LogProses.Dilewati : LogProses.Gagal,
                LogProses.R()
                    .Tambah("Durasi", (durasi / 1000.0).ToString("0.0", LogProses.Id) + " dtk")
                    .Tambah("Langkah selesai", _hasilLangkah == null ? 0 : _hasilLangkah.Count)
                    .TambahBila(error != null, "Error", error));
            string jalanIni = LogProses.IdJalan;

            PanelBridge.Siarkan("runSelesai", new Dictionary<string, object>
            {
                { "status", status }, { "error", error }, { "durasiMs", durasi },
                { "langkah", _hasilLangkah }, { "pesan", pesanModul }
            });

            _antrian = null; _param = null; _wbApp = null;

            var lanjut = SetelahSelesai;
            SetelahSelesai = null;
            if (lanjut != null)
            {
                // Lanjutan (mis. simpan otomatis Hitung semua grup) masih tercatat dengan id jalan ini
                try { lanjut(app, status); }
                catch (Exception ex) { CatatanLog.Tulis("Lanjutan setelah run gagal: " + ex.Message); }
            }
            // Tutup konteks — kecuali lanjutan sudah memulai run baru (grup berikutnya) dengan id sendiri
            if (LogProses.IdJalan == jalanIni) LogProses.SelesaiJalan();
        }

        // ------------------------------------------------------------
        // Penjadwalan langkah berikutnya dengan jeda
        // ------------------------------------------------------------
        private static System.Windows.Forms.Timer _timer;
        private const int JedaAntarLangkahMs = 250;

        /// <summary>
        /// Antrekan langkah berikutnya SETELAH jeda singkat. Timer WinForms
        /// berjalan di thread utama Excel (dibuat di sini, di konteks makro),
        /// jadi Tick selalu datang lewat loop pesan Excel — tepat setelah
        /// pesan-pesan yang tertunda (termasuk "batal") selesai diproses.
        /// </summary>
        private static void JadwalkanBerikut()
        {
            if (_timer == null)
            {
                _timer = new System.Windows.Forms.Timer { Interval = JedaAntarLangkahMs };
                _timer.Tick += (s, e) =>
                {
                    _timer.Stop();
                    ExcelAsyncUtil.QueueAsMacro(LangkahBerikut);
                };
            }
            _timer.Stop();
            _timer.Start();
        }

        [DllImport("user32.dll")]
        private static extern void DisableProcessWindowsGhosting();
        private static bool _ghostingMati;

        internal static void MatikanGhostingSekali()
        {
            if (_ghostingMati) return;
            try { DisableProcessWindowsGhosting(); } catch { }
            _ghostingMati = true;
        }

        private static void AktifkanWorkbookAplikasi()
        {
            try { if (_wbApp != null) _wbApp.Activate(); } catch { }
        }

        // ============================================================
        // Pemanggilan modul — cermin dari CKPNFunctions.cs
        // ============================================================
        private static void JalankanIndividu(Excel.Application app, ParameterCKPN p)
        {
            var x = p.Individu;
            new CKPNIndividu(app).Hitung(x.FilePath, x.TopN, x.NPF, x.Kol2, x.Restru, p.KC);
        }

        private static void JalankanNetFlow(Excel.Application app, ParameterCKPN p)
        {
            var x = p.NetFlow;
            // Pemecahan string identik dengan CKPNFunctions.IsiBucketPDNetflow
            // (RemoveEmptyEntries) supaya hasil sama persis dengan tombol VBA.
            var opt = StringSplitOptions.RemoveEmptyEntries;
            new CKPNPDNetFlow(app).IsiBucket(
                x.FilePaths.Split(new[] { '|' }, opt),
                x.BulanLabels.Split(new[] { '|' }, opt),
                x.RefDates.Split(new[] { '|' }, opt),
                x.TopN, p.KC);
        }

        private static void JalankanMigration(Excel.Application app, ParameterCKPN p)
        {
            var x = p.Migration;
            if (x.RunAll)
                new PDMigration(app).HitungSemua(x.AwalGab, x.AkhirGab, x.TglAwalGab, x.TglAkhirGab, x.TopN, p.KCList);
            else
                new PDMigration(app).Hitung(x.Triwulan, x.PathAwal, x.PathAkhir, x.TglAwal, x.TglAkhir, x.TopN, p.KCList);
        }

        private static void JalankanLgdEr(Excel.Application app, ParameterCKPN p)
        {
            new LGDExpectedRecoveries(app).Hitung(p.LgdEr.CurrYear, p.LgdEr.FilePathsStr, p.KCList);
        }

        private static void JalankanLgdCs(Excel.Application app, ParameterCKPN p)
        {
            new LGDCollateralShortfall(app).Hitung(p.LgdCs.FileConfigStr, p.LgdCs.Haircut, p.KCList);
        }

        private static void JalankanSummary(Excel.Application app, ParameterCKPN p)
        {
            new RefreshSummary(app).Refresh(p.Summary.FilePath, p.KCList);
        }

        // ------------------------------------------------------------
        private static string PesanError(Exception ex)
        {
            // Exception COM sering membungkus penyebab asli di InnerException.
            Exception e = ex;
            while (e is System.Reflection.TargetInvocationException && e.InnerException != null) e = e.InnerException;
            return e.Message;
        }

        private static IEnumerable<string> Ids(List<Langkah> l)
        {
            foreach (var x in l) yield return x.Id;
        }
    }
}
