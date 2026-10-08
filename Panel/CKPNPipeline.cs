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

        /// <summary>
        /// Langkah khusus mode setahun sekali (Tahap 5e): menulis PD &amp; LGD Desember tahun lalu
        /// (acuan) ke sheet kolektif. Tidak tampil di tab Manual; disisipkan setelah CKPN Individu
        /// bila run memakai acuan, menggantikan Net Flow, Migration, LGD ER, dan LGD CS.
        /// </summary>
        private static readonly Langkah LangkahAcuan = new Langkah
        {
            Id = "acuan", Nama = "PD & LGD acuan Desember", SheetDitulis = AcuanTahunan.SheetKol, Jalankan = JalankanAcuan
        };
        private static AcuanTahunan.Acuan _acuan;

        /// <summary>
        /// Langkah otomatis Tahap 5g: ambil Overview Data (OS, EAD, NPF, PPKA per KC dari template
        /// Master!D14) dan simpan sebagai snapshot periode di database. Disisipkan di awal run hanya
        /// bila snapshot periode Master!C4 belum ada atau file template berubah, dan hanya untuk user
        /// pengirim. Kegagalannya tidak menggagalkan perhitungan CKPN (dicatat sebagai peringatan).
        /// </summary>
        private static readonly Langkah LangkahOverview = new Langkah
        {
            Id = "overview", Nama = "Data overview (OS, EAD, NPF)", SheetDitulis = "— (database, sheet tidak diubah)", Jalankan = JalankanOverview
        };
        private static string _alasanOverview;

        /// <summary>Keputusan acuan pada MulaiGrup terakhir (untuk BatchGrup/panel); null = tidak relevan.</summary>
        public static string CatatanAcuan { get; private set; }

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
                                   string sumberLog = null, AcuanTahunan.Acuan acuan = null)
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
                if (l.Id == "individu" && acuan != null) antrian.Add(LangkahAcuan);   // acuan ditulis setelah Individu
            }
            if (acuan != null && !pilih.Contains("individu")) return "Acuan PD & LGD hanya dapat dipakai bersama CKPN Individu.";
            if (antrian.Count == 0) return "Belum ada langkah yang dipilih.";
            if (masalah.Count > 0) return "Parameter belum lengkap:\n- " + string.Join("\n- ", masalah.ToArray());

            // Tahap 5g: snapshot Overview periode ini diambil dulu bila belum tersimpan / template berubah
            _alasanOverview = null;
            try
            {
                DateTime tglOv;
                string alasanOv;
                if (ParameterMaster.BacaTanggalLaporan(wb, out tglOv) &&
                    Data.PenjelasanCKPN.OverviewPerluDiambil(tglOv.ToString("yyyy-MM-dd"), ParameterMaster.BacaPathTemplate(wb), out alasanOv))
                {
                    _alasanOverview = alasanOv;
                    antrian.Insert(0, LangkahOverview);
                }
            }
            catch (Exception ex) { CatatanLog.Tulis("Cek snapshot overview gagal: " + PesanError(ex)); }

            _wbApp        = wb;
            _param        = p;
            _acuan        = acuan;
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
                .Tambah("Top-N Master!C10", ParameterMaster.BacaTopN(wb))
                .TambahBila(acuan != null, "PD & LGD", acuan == null ? null : "acuan " + acuan.Label + " (Net Flow, Migration, LGD ER, LGD CS dilewati)"));
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

                // PD & LGD dihitung penuh pada run ini → acuan lama (bila ada) tidak boleh terpakai lagi.
                // Sekalian memastikan rumus mode setahun terpasang di sheet kolektif (Tahap 5e).
                if (_acuan == null && _antrian.Exists(x => x.Id == "netflow" || x.Id == "migration" || x.Id == "lgder" || x.Id == "lgdcs"))
                {
                    try { AcuanTahunan.Kosongkan(_wbApp); }
                    catch (Exception ex) { CatatanLog.Tulis("Kosongkan acuan PD & LGD gagal: " + PesanError(ex)); }
                }
                // Tahap 5h: PD & LGD ABA tahunan (⚙ Pengaturan) → Summary C21:C23
                try { AktifkanWorkbookAplikasi(); ParameterAba.Terapkan(_wbApp); }
                catch (Exception ex)
                {
                    CatatanLog.Tulis("Parameter ABA gagal diterapkan: " + PesanError(ex));
                    LogProses.Catat("Parameter ABA", "Summary C21:C23", LogProses.Peringatan, LogProses.R().Tambah("Error", PesanError(ex)));
                }
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
        /// <param name="pakaiAcuan">mode setahun sekali, laporan Jan–Nov: pakai PD &amp; LGD Desember tahun lalu
        /// dari database bila tersedia (bila tidak, grup dihitung penuh — lihat <see cref="CatatanAcuan"/>)</param>
        public static string MulaiGrup(Excel.Application app, string kodeKC, bool terapkanPenyesuaian, bool pakaiAcuan = false)
        {
            CatatanAcuan = null;
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
            AcuanTahunan.Acuan acuan = null;
            int? thAcuan = null;
            if (pakaiAcuan && ParameterMaster.BacaTanggalLaporan(wb, out tglG) && AcuanTahunan.ModeTahunan(wb))
                thAcuan = AcuanTahunan.TahunAcuan(tglG);
            if (thAcuan.HasValue)
            {
                string masalahAcuan;
                try { acuan = AcuanTahunan.Cari(thAcuan.Value, kodeKC, out masalahAcuan); }
                catch (Exception ex) { masalahAcuan = PesanError(ex); }
                if (acuan != null)
                {
                    semua = new List<string> { "individu", "summary" };
                    CatatanAcuan = "PD & LGD acuan " + acuan.Label;
                }
                else CatatanAcuan = "Acuan Desember " + thAcuan.Value + " tidak tersedia (" + masalahAcuan + ") — dihitung penuh.";
            }
            return Mulai(app, semua, terapkanPenyesuaian,
                         (BatchGrup.Aktif ? "Panel · Semua grup · " : "Panel · Grup ") + grupLog, acuan);
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

            _antrian = null; _param = null; _wbApp = null; _acuan = null;

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

        private static void JalankanAcuan(Excel.Application app, ParameterCKPN p)
        {
            AcuanTahunan.Tulis(_wbApp ?? PanelBridge.CariWorkbookAplikasi(app), _acuan);
            Data.Penyesuaian.LewatiLgdCs("LGD memakai acuan " + _acuan.Label);
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Mode setahun sekali: PD & LGD memakai versi " + _acuan.Label + " (status periode " + _acuan.Status + ").");
            sb.AppendLine("Net Flow, Migration, LGD ER, dan LGD CS tidak dihitung ulang.");
            sb.Append("LGD weighted " + (_acuan.Lgd.Value * 100).ToString("0.00", LogProses.Id) + "% · PD Migration Kol 1–5: ");
            for (int i = 0; i < 5; i++) sb.Append((i > 0 ? " / " : "") + (_acuan.PdMig[i].Value * 100).ToString("0.00", LogProses.Id) + "%");
            Pemberitahu.Info("PD & LGD acuan Desember", sb.ToString());
            LogProses.Catat("PD & LGD acuan", _acuan.Label, LogProses.OK, LogProses.R()
                .Tambah("Versi acuan", _acuan.RunId)
                .Tambah("Status periode acuan", _acuan.Status)
                .Tambah("LGD weighted", (_acuan.Lgd.Value * 100).ToString("0.00", LogProses.Id) + "%"));
        }

        // Tahap 5g: tidak melempar error — perhitungan CKPN tetap lanjut bila template overview gagal dibaca
        private static void JalankanOverview(Excel.Application app, ParameterCKPN p)
        {
            Excel.Workbook wb = _wbApp ?? PanelBridge.CariWorkbookAplikasi(app);
            try
            {
                var ov = new DataOverviewBuilder(app).UntukPanel(wb);
                string periode = Convert.ToString(ov["periode"]);
                bool ok = Data.PenjelasanCKPN.SimpanOverview(periode, ov, "Otomatis saat hitung · " + Convert.ToString(ov["fileTemplate"]));
                var total = ov["total"] as Dictionary<string, object>;
                LogProses.Catat("Overview data", periode, ok ? LogProses.OK : LogProses.Peringatan, LogProses.R()
                    .Tambah("Alasan", _alasanOverview)
                    .Tambah("File", ov["fileTemplate"])
                    .TambahBila(total != null, "Total OS", total == null ? null : total["totalOS"])
                    .TambahBila(total != null, "NPF", total == null || total["npf"] == null ? null :
                        (Convert.ToDouble(total["npf"]) * 100).ToString("0.00", LogProses.Id) + "%")
                    .TambahBila(!ok, "Keterangan", "snapshot tidak tersimpan ke database"));
            }
            catch (Exception ex)
            {
                string msg = PesanError(ex);
                CatatanLog.Tulis("  Data overview gagal (perhitungan tetap lanjut): " + msg);
                LogProses.Catat("Overview data", "Master!D14", LogProses.Peringatan, LogProses.R()
                    .Tambah("Alasan", _alasanOverview).Tambah("Error", msg));
                Pemberitahu.Info("Data overview belum tersimpan",
                    "Overview Data (OS, EAD, NPF per KC) gagal dibaca dari template Master!D14: " + msg + "\n" +
                    "Perhitungan CKPN tetap dilanjutkan. Muat ulang dari Ringkasan setelah template diperbaiki.");
            }
        }

        private static void JalankanSummary(Excel.Application app, ParameterCKPN p)
        {
            new RefreshSummary(app).Refresh(p.Summary.FilePath, p.KCList);
            // Tahap 5h: bila modul Refresh Summary menulis ulang C21:C23, parameter tahunan tetap berlaku
            ParameterAba.Terapkan(_wbApp ?? PanelBridge.CariWorkbookAplikasi(app));
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
