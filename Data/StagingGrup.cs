using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Globalization;
using System.IO;
using System.Web.Script.Serialization;
using CKPNLibrary.Helpers;
using CKPNLibrary.Modules;
using CKPNLibrary.Panel;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Data
{
    /// <summary>
    /// "Simpan grup ke staging":
    ///   1. Baca isi akhir sheet A. CKPN - INDV dan B4.LGD-CS MACET (setelah edit user).
    ///   2. Bandingkan dengan data dasar hasil sistem (disimpan Penyesuaian saat
    ///      modul berjalan) → deteksi edit manual → jadikan penyesuaian tersimpan.
    ///   3. Simpan hasil grup + nilai Summary ke database sebagai versi baru.
    ///   4. Ekspor snapshot .xlsx ke library\staging\yyyy-MM\.
    ///
    /// Semua method dipanggil di konteks makro (thread utama Excel).
    /// </summary>
    internal static class StagingGrup
    {
        private const double Toleransi = 0.5;   // selisih < Rp0,5 dianggap sama
        private static readonly JavaScriptSerializer _json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        private const string SheetIndv = "A. CKPN - INDV";
        private const string SheetCs   = "B4.LGD-CS MACET";
        private const string SheetSum  = "Summary";

        // Sel Summary yang disimpan (lihat RefreshSummary.cs)
        //   B = hasil perhitungan aplikasi (Net Flow: B6-B8, Migration: B13-B15)
        //   C = CKPN per jenis dari file template (C6-C8)
        //   H = PPKA OJK per KC dari template, C18/C19 = CKPN antar bank
        private static readonly string[,] SelSummary =
        {
            { "B6",  "nf_individu"   }, { "B7",  "nf_kolektif"   }, { "B8",  "nf_total"  },
            { "B13", "mig_individu"  }, { "B14", "mig_kolektif"  }, { "B15", "mig_total" },
            { "C6",  "tpl_individu"  }, { "C7",  "tpl_kolektif"  }, { "C8",  "tpl_total" },
            { "C18", "aba_dijamin"   }, { "C19", "aba_di_atas_plafon" }, { "C26", "aba_ckpn" },
            { "C21", "aba_pd" }, { "C22", "aba_lgd_dijamin" }, { "C23", "aba_lgd_atas" },   // Tahap 5h
            { "H2",  "ppka_KC0500" }, { "H3", "ppka_KC0600" }, { "H4", "ppka_KC0700" }, { "H5", "ppka_KC0800" },
            { "H6",  "ppka_KC0900" }, { "H7", "ppka_KC1000" }, { "H8", "ppka_KC1100" }, { "H9", "ppka_total" }
        };

        // ================================================================
        // Model
        // ================================================================
        internal class BarisIndv
        {
            public int Baris, Urut;
            public string Kc, Cif, Nama, Kontrak, AdaPN;
            public double Os, Jaminan, Biaya, Penurunan;
        }

        internal class BarisCs
        {
            public int Baris;
            public string Rek, Nama, ThnSerah, ThnEks;
            public double Pokok, Agunan, Recovery, Shortfall;
            public bool GRumus;
        }

        internal class Perubahan
        {
            public string Modul;      // individu | lgdcs
            public string Kunci;
            public string Aksi;       // baru | ubah | kembali | hapus-baris | hapus-manual | sistem
            public bool   Hapus;      // true = hapus penyesuaian dari database
            public object Data;       // PenyesuaianIndividu / PenyesuaianLgdCs
            public string Sebelum, Sesudah, Keterangan;
            public bool   PerluAlasan;   // Tahap 3c: artinya "boleh diberi alasan" (opsional, tidak memblokir simpan)
        }

        internal class Rencana
        {
            public Excel.Workbook Wb;
            public DateTime Tanggal;
            public string TanggalStr, KodeKC, BulanLaporan;
            public List<string> KC = new List<string>();
            public List<BarisIndv> Indv = new List<BarisIndv>();
            public List<BarisCs>   Cs   = new List<BarisCs>();
            public List<Perubahan> Perubahan = new List<Perubahan>();
            public List<string>    Peringatan = new List<string>();
            public int VersiBerikut = 1;
            public HashSet<string> KontrakBerPenyesuaian = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public AcuanTahunan.Acuan Acuan;   // Tahap 5e: PD & LGD memakai acuan Desember (LGD CS tidak dihitung)
        }

        // ================================================================
        // 1. Periksa (pratinjau) — tidak menulis apa pun
        // ================================================================
        public static Dictionary<string, object> Periksa(Excel.Application app)
        {
            Rencana r = Susun(app);
            string alasanTulis;
            bool boleh = Database.BolehMenulis(out alasanTulis);

            var daftar = new List<object>();
            foreach (var p in r.Perubahan)
            {
                if (p.Aksi == "sistem") continue;   // pembaruan diam-diam, tidak perlu ditampilkan
                daftar.Add(new Dictionary<string, object>
                {
                    { "modul", p.Modul }, { "kunci", p.Kunci }, { "aksi", p.Aksi },
                    { "sebelum", p.Sebelum }, { "sesudah", p.Sesudah },
                    { "keterangan", p.Keterangan }, { "bisaAlasan", p.PerluAlasan }
                });
            }

            // Grup mana (menurut susunan tahunan) yang sedang disimpan, dan status periodenya
            string namaGrup = null, statusPeriode = "Terbuka";
            if (Database.Ada)
            {
                using (var con = Database.Buka(false))
                {
                    int sumber;
                    var susunan = Periode.Susunan(con, r.Tanggal.Year, out sumber);
                    if (susunan != null)
                    {
                        var g = susunan.Find(x => x.KodeKC == r.KodeKC);
                        if (g != null) namaGrup = g.Nama;
                        else r.Peringatan.Add("Kombinasi KC " + r.KodeKC + " tidak sesuai susunan grup tahun " + sumber +
                                              ". Hasil tetap bisa disimpan, tetapi tidak ikut konsolidasi.");
                    }
                    object st = Database.Scalar(con, "SELECT status FROM periode WHERE tanggal=@p0", r.TanggalStr);
                    if (st != null) statusPeriode = Convert.ToString(st);
                }
            }
            if (statusPeriode == "Final")
                r.Peringatan.Add("Periode ini sudah dikunci (Final). Buka kunci di tab Periode sebelum menyimpan.");

            // Top-N debitur Individu harus sama dengan ketetapan tahunan (SOP) — Tahap 4b
            string tolakTopN = CekTopN(r);
            if (tolakTopN != null) r.Peringatan.Add(tolakTopN);

            // Tahap 5h: PD/LGD ABA di Summary vs parameter tahunan (diterapkan otomatis saat Simpan)
            string bedaAba = ParameterAba.CekBeda(r.Wb);
            if (bedaAba != null) r.Peringatan.Add(bedaAba);

            return new Dictionary<string, object>
            {
                { "namaGrup", namaGrup }, { "statusPeriode", statusPeriode },
                { "periode", r.BulanLaporan }, { "tanggal", r.TanggalStr },
                { "kodeKC", r.KodeKC }, { "versi", r.VersiBerikut },
                { "jumlahIndividu", r.Indv.Count }, { "jumlahLgdCs", r.Cs.Count },
                { "perubahan", daftar }, { "peringatan", r.Peringatan },
                { "bolehMenulis", boleh && statusPeriode != "Final" && tolakTopN == null }, { "infoPengirim", alasanTulis },
                { "topN", ParameterMaster.BacaTopN(r.Wb) },
                { "acuan", r.Acuan == null ? null : "PD & LGD memakai acuan " + r.Acuan.Label + " (LGD CS tidak dihitung bulan ini)" },
                { "berjalan", CKPNPipeline.SedangBerjalan }
            };
        }

        // ================================================================
        // 2. Simpan
        // ================================================================
        /// <param name="alasan">no. rekening LGD CS yang dihapus → alasan (opsional sejak Tahap 3c;
        /// yang menyaring perhitungan bulan berikutnya adalah catatan pengecualiannya, bukan alasannya)</param>
        public static Dictionary<string, object> Simpan(Excel.Application app, Dictionary<string, string> alasan,
                                                        string catatan, bool refreshSummary)
        {
            if (CKPNPipeline.SedangBerjalan)
                throw new InvalidOperationException("Perhitungan sedang berjalan. Tunggu sampai selesai.");

            string alasanTulis;
            if (!Database.BolehMenulis(out alasanTulis)) throw new InvalidOperationException(alasanTulis);

            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");

            var pesanModul = new List<Dictionary<string, object>>();
            CKPNPipeline.MatikanGhostingSekali();   // Refresh Summary + ekspor bisa > 5 detik
            bool su = true;
            try { su = app.ScreenUpdating; } catch { }

            try
            {
                try { app.ScreenUpdating = false; app.Cursor = Excel.XlMousePointer.xlWait; } catch { }

                // ---- a. Refresh Summary agar total mencerminkan edit manual ----
                if (refreshSummary)
                {
                    ParameterCKPN p0 = ParameterMaster.Baca(wb);
                    if (!p0.Siap("summary"))
                        throw new InvalidOperationException("Refresh Summary: " + string.Join("; ", p0.Error["summary"].ToArray()));
                    wb.Activate();
                    Pemberitahu.MulaiModePanel();
                    try { new RefreshSummary(app).Refresh(p0.Summary.FilePath, p0.KCList); }
                    finally { pesanModul.AddRange(Pemberitahu.SelesaiModePanel()); }
                }
                // Tahap 5h: PD & LGD ABA mengikuti parameter tahunan sebelum angka Summary dibaca
                ParameterAba.Terapkan(wb);
                try { app.Calculate(); } catch { }

                // ---- b. Susun ulang rencana dari kondisi sheet terkini ----
                Rencana r = Susun(app);
                foreach (var p in r.Perubahan)
                {
                    if (!p.PerluAlasan) continue;
                    string a;
                    if (alasan == null || !alasan.TryGetValue(p.Kunci, out a) || string.IsNullOrWhiteSpace(a))
                        continue;   // tanpa alasan tetap disimpan sebagai pengecualian (alasan diisi catatan simpan)
                    ((PenyesuaianLgdCs)p.Data).Alasan = a.Trim();
                    p.Keterangan = a.Trim();
                }
                string tolakTopN = CekTopN(r);
                if (tolakTopN != null) throw new InvalidOperationException(tolakTopN);
                var ringkasan = BacaRingkasan(wb);
                string analisisJson = AnalisisPD.BacaWorkbookJson(wb);   // Tahap 4: bahan sankey
                string kolektifJson = RincianCkpn.BacaWorkbookJson(wb);  // Tahap 6c: EAD/PD/LGD/CKPN per bucket & kualitas
                ParameterCKPN param = ParameterMaster.Baca(wb);
                if (param.LgdCs != null) ringkasan["lgd_cs_haircut"] = param.LgdCs.Haircut;   // Tahap 5b: prefill realisasi saat koreksi

                // Tahap 5e: mode acuan — LGD & bahan analisis PD milik versi Desember, bukan isi sheet B1/B2/B4
                if (r.Acuan != null)
                {
                    foreach (var k in new List<string>(ringkasan.Keys))
                        if (k.StartsWith("lgd", StringComparison.OrdinalIgnoreCase)) ringkasan.Remove(k);
                    foreach (var kv in r.Acuan.RingkasanLgd) ringkasan[kv.Key] = kv.Value;
                    ringkasan["lgd_gabungan"] = r.Acuan.Lgd.Value;
                    ringkasan["acuan_run_id"] = r.Acuan.RunId;
                    ringkasan["acuan_versi"] = r.Acuan.Versi;
                    if (!string.IsNullOrEmpty(r.Acuan.AnalisisJson)) analisisJson = r.Acuan.AnalisisJson;
                    catatan = (string.IsNullOrEmpty(catatan) ? "" : catatan + " · ") + "PD & LGD acuan " + r.Acuan.Label;
                }

                // ---- c. Tulis database dalam satu transaksi ----
                Database.Cadangkan();
                long runId;
                int versi;
                using (var con = Database.Buka(true))
                using (var tx = con.BeginTransaction())
                {
                    string now = Database.Sekarang();
                    Database.Exec(con, "INSERT OR IGNORE INTO periode(tanggal, status, dibuat) VALUES(@p0,'Terbuka',@p1)", r.TanggalStr, now);
                    long periodeId = Convert.ToInt64(Database.Scalar(con, "SELECT id FROM periode WHERE tanggal=@p0", r.TanggalStr));
                    Periode.PastikanTerbuka(con, r.TanggalStr);

                    versi = Convert.ToInt32(Database.Scalar(con,
                        "SELECT COALESCE(MAX(versi),0)+1 FROM run_grup WHERE periode_id=@p0 AND kode_kc=@p1", periodeId, r.KodeKC));
                    Database.Exec(con, "UPDATE run_grup SET aktif=0 WHERE periode_id=@p0 AND kode_kc=@p1", periodeId, r.KodeKC);

                    Database.Exec(con,
                        "INSERT INTO run_grup(periode_id,kode_kc,versi,aktif,status,pengguna,komputer,waktu,versi_addin,parameter_json,catatan) " +
                        "VALUES(@p0,@p1,@p2,1,'Final',@p3,@p4,@p5,@p6,@p7,@p8)",
                        periodeId, r.KodeKC, versi, Environment.UserName, Environment.MachineName, now,
                        typeof(StagingGrup).Assembly.GetName().Version.ToString(),
                        _json.Serialize(new Dictionary<string, object> { { "periode", param.BulanLaporan }, { "ringkas", param.Ringkas } }),
                        catatan ?? "");
                    runId = con.LastInsertRowId;
                    Database.Exec(con, "UPDATE run_grup SET top_n=@p0 WHERE id=@p1", ParameterMaster.BacaTopN(wb), runId);

                    foreach (var kv in ringkasan)
                        Database.Exec(con, "INSERT INTO ringkasan(run_id,kunci,nilai) VALUES(@p0,@p1,@p2)", runId, kv.Key, kv.Value);
                    AnalisisPD.SimpanRun(con, runId, analisisJson);
                    RincianCkpn.SimpanRun(con, runId, kolektifJson);

                    var disesuaikan = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var p in r.Perubahan) if (p.Modul == "individu" && !p.Hapus) disesuaikan.Add(p.Kunci);
                    foreach (var b in r.Indv)
                        Database.Exec(con,
                            "INSERT INTO hasil_individu(run_id,urut,kc,cif,nama,no_kontrak,os,ada_pn,jaminan,biaya_jual,penurunan_nilai,disesuaikan) " +
                            "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11)",
                            runId, b.Urut, b.Kc, b.Cif, b.Nama, b.Kontrak, b.Os, b.AdaPN, b.Jaminan, b.Biaya, b.Penurunan,
                            (disesuaikan.Contains(b.Kontrak) || r.KontrakBerPenyesuaian.Contains(b.Kontrak)) ? 1 : 0);
                    SimpanBuktiIndividu(con, runId);   // Tahap 6

                    var dasarCs = DasarCs();
                    foreach (var b in r.Cs)
                        Database.Exec(con,
                            "INSERT INTO hasil_lgdcs(run_id,no_rek,nama,pokok_awal,nilai_agunan,thn_diserahkan,thn_eksekusi,recovery,shortfall,sumber) " +
                            "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9)",
                            runId, b.Rek, b.Nama, b.Pokok, b.Agunan, b.ThnSerah, b.ThnEks, b.Recovery, b.Shortfall,
                            dasarCs != null && dasarCs.ContainsKey(b.Rek) ? "sistem" : "manual");

                    TerapkanPerubahan(con, r, catatan);
                    tx.Commit();
                }
                Periode.CatatSimpan(r.TanggalStr, r.KodeKC, versi);

                // ---- d. Snapshot .xlsx ----
                string snapshot = null, pesanSnapshot = null;
                try
                {
                    string folder = Path.Combine(AppPaths.FolderStaging, r.Tanggal.ToString("yyyy-MM"));
                    Directory.CreateDirectory(folder);
                    string suffix = new ExportToExcel(app).BuatSuffixKC(r.KodeKC);
                    snapshot = Path.Combine(folder, "CKPN_" + r.Tanggal.ToString("yyyyMMdd") + "_" + suffix + "_v" + versi + ".xlsx");
                    wb.Activate();
                    Pemberitahu.MulaiModePanel();
                    try { new ExportToExcel(app).Export(snapshot, r.KodeKC); }
                    finally { Pemberitahu.SelesaiModePanel(); }   // pesan "Export berhasil" tidak perlu ditampilkan

                    using (var con = Database.Buka(true))
                        Database.Exec(con, "UPDATE run_grup SET snapshot=@p0 WHERE id=@p1", snapshot, runId);
                }
                catch (Exception ex)
                {
                    pesanSnapshot = "Data tersimpan, tetapi snapshot .xlsx gagal dibuat: " + ex.Message;
                    snapshot = null;
                    CatatanLog.Tulis(pesanSnapshot);
                }

                // ---- e. Tahap 6b: profil rekening Individu & LGD CS dari template APOLLO (untuk dokumen CKPN) ----
                string pesanProfil = null;
                try
                {
                    string pathTpl = ParameterMaster.BacaPathTemplate(wb);
                    if (!string.IsNullOrEmpty(pathTpl) && File.Exists(pathTpl))
                    {
                        var rekProfil = new List<string>();
                        foreach (var b in r.Indv) rekProfil.Add(b.Kontrak);
                        foreach (var c in r.Cs) rekProfil.Add(c.Rek);
                        var hp = ProfilTemplate.AmbilDanSimpan(app, r.TanggalStr, pathTpl, rekProfil, true);
                        var tidak = hp["tidakDitemukan"] as List<string>;
                        if (tidak != null && tidak.Count > 0)
                            pesanProfil = "Profil template tidak ditemukan untuk " + tidak.Count + " rekening (mis. baris manual LGD CS): " +
                                          string.Join(", ", tidak.GetRange(0, Math.Min(5, tidak.Count)).ToArray());
                    }
                    wb.Activate();
                }
                catch (Exception ex)
                {
                    pesanProfil = "Data tersimpan, tetapi profil debitur dari template gagal dibaca: " + ex.Message;
                    CatatanLog.Tulis(pesanProfil);
                }

                CatatanLog.Tulis("SIMPAN GRUP " + r.KodeKC + " periode " + r.TanggalStr + " v" + versi +
                                 " · " + r.Perubahan.Count + " perubahan penyesuaian");

                double v;
                LogProses.CatatPanel(r.TanggalStr, "Simpan grup", r.KodeKC + " v" + versi, LogProses.OK, LogProses.R()
                    .Tambah("CKPN Net Flow", ringkasan.TryGetValue("nf_total", out v) ? (object)v : null)
                    .Tambah("CKPN Migration", ringkasan.TryGetValue("mig_total", out v) ? (object)v : null)
                    .Tambah("PPKA", ringkasan.TryGetValue("ppka_total", out v) ? (object)v : null)
                    .Tambah("Perubahan penyesuaian", r.Perubahan.Count)
                    .Tambah("Catatan", catatan)
                    .Tambah("Snapshot", snapshot == null ? (pesanSnapshot ?? "-") : Path.GetFileName(snapshot)));
                return new Dictionary<string, object>
                {
                    { "runId", runId }, { "versi", versi }, { "kodeKC", r.KodeKC }, { "periode", r.BulanLaporan },
                    { "nfTotal",   ringkasan.TryGetValue("nf_total", out v)   ? (object)v : null },
                    { "migTotal",  ringkasan.TryGetValue("mig_total", out v)  ? (object)v : null },
                    { "ppkaTotal", ringkasan.TryGetValue("ppka_total", out v) ? (object)v : null },
                    { "jumlahPerubahan", r.Perubahan.Count },
                    { "snapshot", snapshot }, { "pesanSnapshot", pesanSnapshot }, { "pesanProfil", pesanProfil },
                    { "pesan", pesanModul }
                };
            }
            finally
            {
                try { app.ScreenUpdating = su; app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            }
        }

        // Tahap 6: bukti objektif penurunan nilai (kualitas, hari tunggakan, restrukturisasi, dasar PN)
        // dari data dasar perhitungan Individu terakhir di PC ini. Kontrak tanpa data dibiarkan NULL.
        private static void SimpanBuktiIndividu(SQLiteConnection con, long runId)
        {
            try
            {
                var dasar = Penyesuaian.BacaDasar("individu.json");
                object baris;
                if (dasar == null || !dasar.TryGetValue("baris", out baris) || baris == null) return;
                foreach (var x in (IEnumerable)baris)
                {
                    var b = x as Dictionary<string, object>;
                    if (b == null) continue;
                    object kual, hari, restru, dasarPn;
                    b.TryGetValue("Kualitas", out kual);
                    b.TryGetValue("HariTunggakan", out hari);
                    b.TryGetValue("Restruktur", out restru);
                    b.TryGetValue("DasarPN", out dasarPn);
                    if (kual == null && dasarPn == null) continue;
                    Database.Exec(con,
                        "UPDATE hasil_individu SET kualitas=@p0, hari_tunggakan=@p1, restruktur=@p2, dasar_pn=@p3 WHERE run_id=@p4 AND no_kontrak=@p5",
                        kual == null ? (object)DBNull.Value : Convert.ToInt32(kual),
                        hari == null ? (object)DBNull.Value : Convert.ToDouble(hari),
                        restru == null ? (object)DBNull.Value : (Convert.ToBoolean(restru) ? 1 : 0),
                        dasarPn == null ? (object)DBNull.Value : Convert.ToString(dasarPn),
                        runId, Convert.ToString(b["Kontrak"]));
                }
            }
            catch (Exception ex) { CatatanLog.Tulis("Simpan bukti penurunan nilai gagal: " + ex.Message); }
        }

        // ================================================================
        // 3. Hitung ulang Summary saja (setelah edit Individu / LGD CS),
        //    untuk melihat total sebelum menyimpan.
        // ================================================================
        public static Dictionary<string, object> RefreshSaja(Excel.Application app)
        {
            if (CKPNPipeline.SedangBerjalan)
                throw new InvalidOperationException("Perhitungan sedang berjalan. Tunggu sampai selesai.");
            Excel.Workbook wb = PanelBridge.CariWorkbookAplikasi(app);
            if (wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");
            ParameterCKPN p = ParameterMaster.Baca(wb);
            if (!p.Siap("summary"))
                throw new InvalidOperationException(string.Join("; ", p.Error["summary"].ToArray()));

            CKPNPipeline.MatikanGhostingSekali();
            List<Dictionary<string, object>> pesan;
            bool su = true;
            try { su = app.ScreenUpdating; } catch { }
            try
            {
                try { app.ScreenUpdating = false; app.Cursor = Excel.XlMousePointer.xlWait; } catch { }
                wb.Activate();
                Pemberitahu.MulaiModePanel();
                try { new RefreshSummary(app).Refresh(p.Summary.FilePath, p.KCList); }
                finally { pesan = Pemberitahu.SelesaiModePanel(); }
                ParameterAba.Terapkan(wb);   // Tahap 5h
                try { app.Calculate(); } catch { }
            }
            finally
            {
                try { app.ScreenUpdating = su; app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            }

            var rg = BacaRingkasan(wb);
            double ppkaGrup = 0, v;
            foreach (var kc in p.KC) if (rg.TryGetValue("ppka_" + kc, out v)) ppkaGrup += v;
            return new Dictionary<string, object>
            {
                { "kodeKC", p.KCList },
                { "nfTotal",  rg.TryGetValue("nf_total", out v)  ? (object)v : null },
                { "migTotal", rg.TryGetValue("mig_total", out v) ? (object)v : null },
                { "ppkaGrup", ppkaGrup },
                { "pesan", pesan }
            };
        }

        // ================================================================
        // Susun rencana: baca sheet, bandingkan dengan data dasar & penyesuaian
        // ================================================================
        private static Rencana Susun(Excel.Application app)
        {
            var r = new Rencana();
            r.Wb = PanelBridge.CariWorkbookAplikasi(app);
            if (r.Wb == null) throw new InvalidOperationException("Workbook aplikasi CKPN tidak sedang terbuka.");

            if (!ParameterMaster.BacaTanggalLaporan(r.Wb, out r.Tanggal))
                throw new InvalidOperationException("Tanggal laporan (Master!C4) belum diisi / tidak valid.");
            r.TanggalStr   = r.Tanggal.ToString("yyyy-MM-dd");
            r.BulanLaporan = r.Tanggal.ToString("dd MMMM yyyy", new CultureInfo("id-ID"));

            Excel.Worksheet master = ParameterMaster.CariSheet(r.Wb, "Master");
            r.KC = ParameterMaster.BacaKCDicentang(master);
            if (r.KC.Count == 0) throw new InvalidOperationException("Tidak ada KC yang dicentang di Master.");
            r.KodeKC = string.Join(",", r.KC.ToArray());   // urutan tetap KC0600..KC1100

            Excel.Worksheet wsI = ParameterMaster.CariSheet(r.Wb, SheetIndv);
            Excel.Worksheet wsC = ParameterMaster.CariSheet(r.Wb, SheetCs);
            if (wsI != null) r.Indv = BacaIndividu(wsI);

            // Tahap 5e: mode setahun sekali dengan acuan Desember → sheet B4 bukan milik run ini
            string masalahAcuan;
            r.Acuan = AcuanTahunan.AktifDiWorkbook(r.Wb, out masalahAcuan);
            if (masalahAcuan != null)
                throw new InvalidOperationException("Acuan PD & LGD di sheet kolektif tidak dapat dibaca: " + masalahAcuan +
                                                    " Hitung ulang grup ini.");
            if (r.Acuan != null)
            {
                if (!string.Equals(r.Acuan.KodeKC, r.KodeKC, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Acuan PD & LGD di sheet milik grup " + r.Acuan.KodeKC +
                                                        ", sedangkan KC yang dicentang " + r.KodeKC + ". Hitung ulang grup ini.");
            }
            else if (wsC != null) r.Cs = BacaLgdCs(wsC);

            // Penyesuaian tersimpan (read-only). Database belum ada = kosong.
            var ovInd = new Dictionary<string, PenyesuaianIndividu>(StringComparer.OrdinalIgnoreCase);
            var ovCs  = new Dictionary<string, PenyesuaianLgdCs>(StringComparer.OrdinalIgnoreCase);
            if (Database.Ada)
            {
                ovInd = Penyesuaian.MuatIndividu();
                ovCs  = Penyesuaian.MuatLgdCs();
                using (var con = Database.Buka(false))
                {
                    object v = Database.Scalar(con,
                        "SELECT COALESCE(MAX(g.versi),0)+1 FROM run_grup g JOIN periode p ON p.id=g.periode_id " +
                        "WHERE p.tanggal=@p0 AND g.kode_kc=@p1", r.TanggalStr, r.KodeKC);
                    r.VersiBerikut = Convert.ToInt32(v);
                }
            }

            foreach (var k in ovInd.Keys) r.KontrakBerPenyesuaian.Add(k);
            BandingkanIndividu(r, ovInd);
            if (r.Acuan == null) BandingkanLgdCs(r, ovCs);   // mode acuan: LGD CS tidak dihitung → tidak ada perubahan
            return r;
        }

        // ---------------- Individu ----------------
        private static void BandingkanIndividu(Rencana r, Dictionary<string, PenyesuaianIndividu> ov)
        {
            var dasar = Penyesuaian.BacaDasar("individu.json");
            if (dasar == null)
            {
                r.Peringatan.Add("Data dasar CKPN Individu tidak ditemukan di PC ini — edit manual Individu tidak dapat dideteksi. " +
                                 "Jalankan CKPN Individu di PC ini terlebih dahulu.");
                return;
            }
            bool diterapkan = AmbilBool(dasar, "penyesuaianDiterapkan");
            if (!diterapkan)
                r.Peringatan.Add("CKPN Individu dihitung TANPA penyesuaian tersimpan — hanya edit baru yang dicatat; penyesuaian lama tidak dihapus.");

            var lewati = DiubahDiPanel("individu", dasar, r);

            var sistem = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in AmbilDaftar(dasar, "baris"))
                sistem[Convert.ToString(d["Kontrak"])] = Convert.ToDouble(d["JaminanSistem"]);

            foreach (var b in r.Indv)
            {
                double sys;
                if (!sistem.TryGetValue(b.Kontrak, out sys)) continue;   // tidak dari run terakhir → abaikan
                if (lewati.Contains(b.Kontrak)) continue;                // diubah di panel setelah dihitung

                bool beda = Math.Abs(b.Jaminan - sys) > Toleransi || Math.Abs(b.Biaya) > Toleransi;
                PenyesuaianIndividu lama;
                ov.TryGetValue(b.Kontrak, out lama);

                var baru = new PenyesuaianIndividu
                {
                    NoKontrak = b.Kontrak, Jaminan = b.Jaminan, BiayaJual = b.Biaya, JaminanSistem = sys,
                    Pengguna = Environment.UserName, Waktu = Database.Sekarang(), Periode = r.TanggalStr
                };
                string sesudah = "Agunan " + Rp(b.Jaminan) + " · Biaya " + Rp(b.Biaya);

                if (beda)
                {
                    if (lama == null)
                        r.Perubahan.Add(new Perubahan { Modul = "individu", Kunci = b.Kontrak, Aksi = "baru", Data = baru,
                            Sebelum = "Sistem " + Rp(sys), Sesudah = sesudah });
                    else if (Math.Abs(lama.Jaminan - b.Jaminan) > Toleransi || Math.Abs(lama.BiayaJual - b.Biaya) > Toleransi)
                        r.Perubahan.Add(new Perubahan { Modul = "individu", Kunci = b.Kontrak, Aksi = "ubah", Data = baru,
                            Sebelum = "Agunan " + Rp(lama.Jaminan) + " · Biaya " + Rp(lama.BiayaJual), Sesudah = sesudah });
                    else if (Math.Abs(lama.JaminanSistem - sys) > Toleransi)
                    {
                        baru.Alasan = lama.Alasan;   // hanya nilai sistem yang diperbarui
                        r.Perubahan.Add(new Perubahan { Modul = "individu", Kunci = b.Kontrak, Aksi = "sistem", Data = baru,
                            Sebelum = "Sistem " + Rp(lama.JaminanSistem), Sesudah = "Sistem " + Rp(sys) });
                    }
                }
                else if (lama != null && diterapkan)
                {
                    r.Perubahan.Add(new Perubahan { Modul = "individu", Kunci = b.Kontrak, Aksi = "kembali", Hapus = true,
                        Sebelum = "Agunan " + Rp(lama.Jaminan) + " · Biaya " + Rp(lama.BiayaJual),
                        Sesudah = "Kembali ke nilai sistem " + Rp(sys) });
                }
            }
        }

        // ---------------- LGD CS ----------------
        private static Dictionary<string, Dictionary<string, object>> DasarCs()
        {
            var dasar = Penyesuaian.BacaDasar("lgdcs.json");
            if (dasar == null) return null;
            var hasil = new Dictionary<string, Dictionary<string, object>>(StringComparer.OrdinalIgnoreCase);
            foreach (var d in AmbilDaftar(dasar, "baris")) hasil[Convert.ToString(d["Rek"])] = d;
            return hasil;
        }

        private static void BandingkanLgdCs(Rencana r, Dictionary<string, PenyesuaianLgdCs> ov)
        {
            var dasarFile = Penyesuaian.BacaDasar("lgdcs.json");
            if (dasarFile == null)
            {
                r.Peringatan.Add("Data dasar LGD CS tidak ditemukan di PC ini — edit manual LGD CS tidak dapat dideteksi. " +
                                 "Jalankan LGD Collateral Shortfall di PC ini terlebih dahulu.");
                return;
            }
            bool diterapkan = AmbilBool(dasarFile, "penyesuaianDiterapkan");
            if (!diterapkan)
                r.Peringatan.Add("LGD CS dihitung TANPA penyesuaian tersimpan — hanya edit baru yang dicatat; penyesuaian lama tidak dihapus.");
            var dasar = DasarCs();
            var lewati = DiubahDiPanel("lgdcs", dasarFile, r);

            var diSheet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var b in r.Cs)
            {
                diSheet.Add(b.Rek);
                if (lewati.Contains(b.Rek)) continue;   // diubah di panel setelah dihitung
                PenyesuaianLgdCs lama;
                ov.TryGetValue(b.Rek, out lama);
                Dictionary<string, object> d;

                if (dasar.TryGetValue(b.Rek, out d))
                {
                    // ---- baris dari sistem: cek perubahan nilai agunan (D) dan realisasi (G) ----
                    double agunanSys = Convert.ToDouble(d["Agunan"]);
                    double? agunanBaru = Math.Abs(b.Agunan - agunanSys) > Toleransi ? (double?)b.Agunan : null;
                    double? recBaru    = b.GRumus ? null : (double?)b.Recovery;

                    if (agunanBaru == null && recBaru == null)
                    {
                        if (lama != null && lama.Jenis == "ubah" && diterapkan)
                            r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = b.Rek, Aksi = "kembali", Hapus = true,
                                Sebelum = UraianUbah(lama), Sesudah = "Kembali ke nilai sistem" });
                        continue;
                    }

                    var baru = new PenyesuaianLgdCs
                    {
                        NoRek = b.Rek, Jenis = "ubah", NilaiAgunan = agunanBaru, Recovery = recBaru, Nama = b.Nama,
                        Pengguna = Environment.UserName, Waktu = Database.Sekarang(), Periode = r.TanggalStr
                    };
                    if (lama == null || lama.Jenis != "ubah")
                        r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = b.Rek, Aksi = "baru", Data = baru,
                            Sebelum = "Sistem", Sesudah = UraianUbah(baru) });
                    else if (!SamaNullable(lama.NilaiAgunan, agunanBaru) || !SamaNullable(lama.Recovery, recBaru))
                        r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = b.Rek, Aksi = "ubah", Data = baru,
                            Sebelum = UraianUbah(lama), Sesudah = UraianUbah(baru) });
                }
                else
                {
                    // ---- baris tambahan user (tidak ada di hasil sistem) ----
                    var baru = new PenyesuaianLgdCs
                    {
                        NoRek = b.Rek, Jenis = "tambah", Nama = b.Nama, PokokAwal = b.Pokok, NilaiAgunan = b.Agunan,
                        Recovery = b.GRumus ? null : (double?)b.Recovery,
                        ThnDiserahkan = b.ThnSerah, ThnEksekusi = b.ThnEks,
                        Pengguna = Environment.UserName, Waktu = Database.Sekarang(), Periode = r.TanggalStr
                    };
                    string uraian = (b.Nama ?? "") + " · pokok " + Rp(b.Pokok) + " · realisasi " + Rp(b.Recovery);
                    if (lama == null || lama.Jenis != "tambah")
                        r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = b.Rek, Aksi = "baru", Data = baru,
                            Sebelum = "Tidak ada di hasil sistem", Sesudah = "Baris manual: " + uraian });
                    else if (!SamaNullable(lama.PokokAwal, b.Pokok) || !SamaNullable(lama.NilaiAgunan, b.Agunan) ||
                             !SamaNullable(lama.Recovery, baru.Recovery) || lama.Nama != b.Nama ||
                             lama.ThnDiserahkan != b.ThnSerah || lama.ThnEksekusi != b.ThnEks)
                        r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = b.Rek, Aksi = "ubah", Data = baru,
                            Sebelum = "Baris manual (versi lama)", Sesudah = "Baris manual: " + uraian });
                }
            }

            // ---- baris sistem yang dihapus user → pengecualian baru (alasan opsional) ----
            foreach (var kv in dasar)
            {
                if (AmbilBool(kv.Value, "Dikecualikan")) continue;   // sudah dikecualikan sebelumnya
                if (diSheet.Contains(kv.Key) || lewati.Contains(kv.Key)) continue;
                PenyesuaianLgdCs lama;
                if (ov.TryGetValue(kv.Key, out lama) && lama.Jenis == "hapus") continue;

                string nama = Convert.ToString(kv.Value["Nama"] ?? "");
                r.Perubahan.Add(new Perubahan
                {
                    Modul = "lgdcs", Kunci = kv.Key, Aksi = "hapus-baris", PerluAlasan = true,
                    Data = new PenyesuaianLgdCs
                    {
                        NoRek = kv.Key, Jenis = "hapus", Nama = nama,
                        PokokAwal = Convert.ToDouble(kv.Value["Pokok"]),
                        Pengguna = Environment.UserName, Waktu = Database.Sekarang(), Periode = r.TanggalStr
                    },
                    Sebelum = nama + " · pokok " + Rp(Convert.ToDouble(kv.Value["Pokok"])),
                    Sesudah = "Dikecualikan (bukan lunas karena eksekusi agunan)"
                });
            }

            // ---- baris manual lama yang dihapus user → hapus penyesuaian 'tambah' ----
            if (diterapkan)
                foreach (var lama in ov.Values)
                    if (lama.Jenis == "tambah" && !diSheet.Contains(lama.NoRek) && !dasar.ContainsKey(lama.NoRek) &&
                        !lewati.Contains(lama.NoRek))
                        r.Perubahan.Add(new Perubahan { Modul = "lgdcs", Kunci = lama.NoRek, Aksi = "hapus-manual", Hapus = true,
                            Sebelum = "Baris manual " + (lama.Nama ?? ""), Sesudah = "Dihapus" });
        }

        /// <summary>
        /// Top-N di Master!C10 dibandingkan dengan Top-N grup (kombinasi KC yang dicentang) di susunan tahunan.
        /// Null = sesuai (atau belum ada ketetapan). Hitung grup dari tab Per grup menulis
        /// nilai tahunan ke C10 secara otomatis, sehingga ketidaksesuaian biasanya berasal
        /// dari hitung manual / tombol VBA dengan C10 yang diubah.
        /// </summary>
        private static string CekTopN(Rencana r)
        {
            string nama;
            int? ketetapan = Periode.TopNGrup(r.Tanggal.Year, r.KodeKC, out nama);
            if (!ketetapan.HasValue) return null;
            int master = ParameterMaster.BacaTopN(r.Wb);
            if (master == ketetapan.Value) return null;
            return "Top-N di Master!C10 (" + master + ") berbeda dengan ketetapan grup " + nama + " tahun " + r.Tanggal.Year +
                   " (" + ketetapan.Value + "). Hasil tidak dapat disimpan ke staging. Hitung ulang dari tab Per grup " +
                   "(Top-N otomatis disesuaikan) atau ubah Master!C10 lalu hitung ulang.";
        }

        /// <summary>
        /// Kunci yang penyesuaiannya ditambah/diubah/dihapus dari panel SETELAH sheet dihitung.
        /// Sheet belum memuat nilai baru itu; bila dibandingkan, simpan akan membatalkan
        /// perubahan dari panel. Kunci ini dilewati dan user diminta menghitung ulang grup.
        /// </summary>
        private static HashSet<string> DiubahDiPanel(string modul, Dictionary<string, object> dasar, Rencana r)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            object w;
            if (!Database.Ada || dasar == null || !dasar.TryGetValue("waktu", out w) || w == null) return set;
            string waktuHitung = Convert.ToString(w);
            try
            {
                using (var con = Database.Buka(false))
                using (var cmd = Database.Cmd(con,
                    "SELECT DISTINCT kunci FROM log_penyesuaian WHERE (modul=@p0 OR modul='semua') " +
                    "AND (aksi LIKE '%-dari-panel' OR aksi IN ('hapus-data-periode','kosongkan-database')) AND waktu>@p1",
                    modul, waktuHitung))
                using (var rd = cmd.ExecuteReader())
                    while (rd.Read()) set.Add(Convert.ToString(rd[0]));
            }
            catch (Exception ex) { CatatanLog.Tulis("Cek perubahan panel: " + ex.Message); }

            if (set.Contains("*"))
            {
                set.Remove("*");
                r.Peringatan.Add("Sebagian penyesuaian " + (modul == "individu" ? "Individu" : "LGD CS") +
                                 " dihapus lewat pengelolaan database setelah sheet dihitung. Hitung ulang grup ini sebelum menyimpan " +
                                 "agar penyesuaian yang sudah dihapus tidak tercatat lagi.");
            }
            if (set.Count > 0)
            {
                var daftar = new List<string>(set);
                daftar.Sort(StringComparer.OrdinalIgnoreCase);
                r.Peringatan.Add("Penyesuaian " + (modul == "individu" ? "Individu" : "LGD CS") + " berikut diubah dari panel setelah sheet dihitung, " +
                                 "sehingga sheet belum memuat nilainya: " + string.Join(", ", daftar.ToArray()) +
                                 ". Kunci ini tidak dibandingkan saat simpan — hitung ulang grup agar nilai terbaru dipakai.");
            }
            return set;
        }

        // ================================================================
        // Terapkan perubahan penyesuaian ke database (+ log audit)
        // ================================================================
        private static void TerapkanPerubahan(SQLiteConnection con, Rencana r, string catatan)
        {
            string now = Database.Sekarang(), user = Environment.UserName;
            foreach (var p in r.Perubahan)
            {
                if (p.Modul == "individu")
                {
                    if (p.Hapus)
                        Database.Exec(con, "DELETE FROM penyesuaian_individu WHERE no_kontrak=@p0", p.Kunci);
                    else
                    {
                        var d = (PenyesuaianIndividu)p.Data;
                        if (string.IsNullOrEmpty(d.Alasan)) d.Alasan = catatan ?? "";
                        Database.Exec(con,
                            "INSERT OR REPLACE INTO penyesuaian_individu(no_kontrak,jaminan,biaya_jual,jaminan_sistem,alasan,pengguna,waktu,periode) " +
                            "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7)",
                            d.NoKontrak, d.Jaminan, d.BiayaJual, d.JaminanSistem, d.Alasan, d.Pengguna, d.Waktu, d.Periode);
                    }
                }
                else
                {
                    if (p.Hapus)
                        Database.Exec(con, "DELETE FROM penyesuaian_lgdcs WHERE no_rek=@p0", p.Kunci);
                    else
                    {
                        var d = (PenyesuaianLgdCs)p.Data;
                        if (string.IsNullOrEmpty(d.Alasan)) d.Alasan = catatan ?? "";
                        Database.Exec(con,
                            "INSERT OR REPLACE INTO penyesuaian_lgdcs(no_rek,jenis,nilai_agunan,recovery,nama,pokok_awal,thn_diserahkan,thn_eksekusi,alasan,pengguna,waktu,periode) " +
                            "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11)",
                            d.NoRek, d.Jenis, d.NilaiAgunan, d.Recovery, d.Nama, d.PokokAwal, d.ThnDiserahkan, d.ThnEksekusi,
                            d.Alasan, d.Pengguna, d.Waktu, d.Periode);
                    }
                }

                Database.Exec(con,
                    "INSERT INTO log_penyesuaian(waktu,pengguna,periode,modul,kunci,aksi,sebelum,sesudah,alasan) " +
                    "VALUES(@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8)",
                    now, user, r.TanggalStr, p.Modul, p.Kunci, p.Aksi, p.Sebelum, p.Sesudah,
                    p.Keterangan ?? catatan ?? "");
            }
        }

        // ================================================================
        // Baca sheet
        // ================================================================
        private static List<BarisIndv> BacaIndividu(Excel.Worksheet ws)
        {
            var hasil = new List<BarisIndv>();
            int last = BarisTerakhir(ws, "F");
            if (last < 6) return hasil;

            object[,] v = Blok(ws, "B6", "K" + last);
            for (int i = 1; i <= v.GetLength(0); i++)
            {
                if (!(v[i, 1] is double)) continue;            // baris TOTAL / kosong
                string kontrak = Teks(v[i, 5]);
                if (kontrak.Length == 0) continue;
                hasil.Add(new BarisIndv
                {
                    Baris = 5 + i, Urut = (int)(double)v[i, 1],
                    Kc = Teks(v[i, 2]), Cif = Teks(v[i, 3]), Nama = Teks(v[i, 4]), Kontrak = kontrak,
                    Os = Angka(v[i, 6]), AdaPN = Teks(v[i, 7]),
                    Jaminan = Angka(v[i, 8]), Biaya = Angka(v[i, 9]), Penurunan = Angka(v[i, 10])
                });
            }
            return hasil;
        }

        private static List<BarisCs> BacaLgdCs(Excel.Worksheet ws)
        {
            var hasil = new List<BarisCs>();
            int last = BarisTerakhir(ws, "B");
            if (last < 5) return hasil;

            object[,] v = Blok(ws, "B5", "I" + last);
            object[,] fG = BlokRumus(ws, "G5", "G" + last);
            for (int i = 1; i <= v.GetLength(0); i++)
            {
                string rek = Teks(v[i, 1]);
                if (rek.Equals("Total", StringComparison.OrdinalIgnoreCase)) break;   // akhir tabel
                if (rek.Length == 0) continue;   // baris dikosongkan user (bukan dihapus) → lewati
                string rumus = fG == null ? "" : Convert.ToString(fG[i, 1] ?? "");
                hasil.Add(new BarisCs
                {
                    Baris = 4 + i, Rek = rek,
                    Pokok = Angka(v[i, 2]), Agunan = Angka(v[i, 3]),
                    ThnSerah = Teks(v[i, 4]), ThnEks = Teks(v[i, 5]),
                    Recovery = Angka(v[i, 6]), Shortfall = Angka(v[i, 7]), Nama = Teks(v[i, 8]),
                    GRumus = rumus.StartsWith("=")
                });
            }
            return hasil;
        }

        private static Dictionary<string, double> BacaRingkasan(Excel.Workbook wb)
        {
            var hasil = new Dictionary<string, double>();
            Excel.Worksheet sum = ParameterMaster.CariSheet(wb, SheetSum);
            if (sum != null)
                for (int i = 0; i < SelSummary.GetLength(0); i++)
                {
                    object v = ((Excel.Range)sum.Range[SelSummary[i, 0]]).Value2;
                    if (v is double) hasil[SelSummary[i, 1]] = (double)v;
                }

            // LGD Weighted dari blok ringkasan sheet B4 (kolom C=CS, D=ER, E=gabungan)
            Excel.Worksheet cs = ParameterMaster.CariSheet(wb, SheetCs);
            if (cs != null)
            {
                int last = BarisTerakhir(cs, "B");
                for (int r = last; r >= 5 && r > last - 30; r--)
                {
                    if (Teks(((Excel.Range)cs.Cells[r, "B"]).Value2) != "LGD Weighted") continue;
                    object c = ((Excel.Range)cs.Cells[r, "C"]).Value2;
                    object d = ((Excel.Range)cs.Cells[r, "D"]).Value2;
                    object e = ((Excel.Range)cs.Cells[r, "E"]).Value2;
                    if (c is double) hasil["lgd_cs"] = (double)c;
                    if (d is double) hasil["lgd_er"] = (double)d;
                    if (e is double) hasil["lgd_gabungan"] = (double)e;

                    // Tahap 5b: total di atasnya (Total WO, Total Recovery) — bahan koreksi LGD CS dari panel
                    for (int r2 = r - 1; r2 >= r - 3 && r2 >= 5; r2--)
                    {
                        string lab = Teks(((Excel.Range)cs.Cells[r2, "B"]).Value2);
                        string akhiran = lab == "Total WO" ? "_wo" : lab == "Total Recovery" ? "_rec" : null;
                        if (akhiran == null) continue;
                        object vc = ((Excel.Range)cs.Cells[r2, "C"]).Value2;
                        object vd = ((Excel.Range)cs.Cells[r2, "D"]).Value2;
                        if (vc is double) hasil["lgd_cs" + akhiran] = (double)vc;
                        if (vd is double) hasil["lgd_er" + akhiran] = (double)vd;
                    }
                    break;
                }
            }
            return hasil;
        }

        // ================================================================
        // Util
        // ================================================================
        private static int BarisTerakhir(Excel.Worksheet ws, string kol)
        {
            Excel.Range c = (Excel.Range)ws.Cells[ws.Rows.Count, kol];
            return ((Excel.Range)c.End[Excel.XlDirection.xlUp]).Row;
        }

        // Range.Value2 selalu dinormalisasi ke object[,] berindeks 1
        private static object[,] Blok(Excel.Worksheet ws, string a, string b)
        {
            object v = ((Excel.Range)ws.Range[a, b]).Value2;
            var arr = v as object[,];
            if (arr != null) return arr;
            var satu = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
            satu[1, 1] = v;
            return satu;
        }

        private static object[,] BlokRumus(Excel.Worksheet ws, string a, string b)
        {
            object v = ((Excel.Range)ws.Range[a, b]).Formula;
            var arr = v as object[,];
            if (arr != null) return arr;
            var satu = (object[,])Array.CreateInstance(typeof(object), new[] { 1, 1 }, new[] { 1, 1 });
            satu[1, 1] = v;
            return satu;
        }

        private static string Teks(object v)
        {
            if (v == null || v is int) return "";   // int = kode error sel
            // No. rekening yang terlanjur tersimpan sebagai angka: tulis tanpa notasi ilmiah
            if (v is double && Math.Abs((double)v % 1) < 1e-9 && Math.Abs((double)v) >= 1e6)
                return ((double)v).ToString("0", CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture).Trim();
        }

        private static double Angka(object v)
        {
            if (v is double) return (double)v;
            if (v == null || v is int) return 0;
            double d;
            return double.TryParse(Convert.ToString(v), NumberStyles.Any, CultureInfo.CurrentCulture, out d) ? d : 0;
        }

        private static bool SamaNullable(double? a, double? b)
        {
            if (!a.HasValue && !b.HasValue) return true;
            if (a.HasValue != b.HasValue) return false;
            return Math.Abs(a.Value - b.Value) <= Toleransi;
        }

        private static string UraianUbah(PenyesuaianLgdCs p)
        {
            var s = new List<string>();
            if (p.NilaiAgunan.HasValue) s.Add("agunan " + Rp(p.NilaiAgunan.Value));
            if (p.Recovery.HasValue)    s.Add("realisasi " + Rp(p.Recovery.Value));
            return s.Count == 0 ? "-" : string.Join(" · ", s.ToArray());
        }

        private static string Rp(double v) { return v.ToString("#,##0", new CultureInfo("id-ID")); }

        private static bool AmbilBool(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) && v is bool && (bool)v;
        }

        private static List<Dictionary<string, object>> AmbilDaftar(Dictionary<string, object> d, string k)
        {
            var hasil = new List<Dictionary<string, object>>();
            object v;
            if (d == null || !d.TryGetValue(k, out v) || !(v is IEnumerable)) return hasil;
            foreach (var x in (IEnumerable)v)
            {
                var item = x as Dictionary<string, object>;
                if (item != null) hasil.Add(item);
            }
            return hasil;
        }
    }
}
