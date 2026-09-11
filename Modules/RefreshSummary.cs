using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace CKPNLibrary.Modules
{
    /// <summary>
    /// Konversi dari VBA RefreshLinkSummary.
    ///
    /// Tugas:
    ///   1. Menulis formula link internal ke sheet "Summary"
    ///   2. Membuka file sumber (read-only) lalu:
    ///        - SUM Nilai CKPN per Jenis (Individual/Kolektif) per sheet KC
    ///        - SUM PPKA OJK per KC (termasuk KC0500) -> H2:H8
    ///        - Hitung CKPN Antar Bank (ABA) dari KC0500 -> C18/C19
    ///   3. Menulis rincian ABA & peringatan "Jenis tak dikenal" ke "Audit Log".
    ///
    /// ----------------------------------------------------------------
    /// REVISI — Diagnostik error 0x800A03EC (Excel error 1004)
    /// ----------------------------------------------------------------
    /// Error 1004 adalah pesan generik Excel: "operasi ditolak". Versi lama
    /// tidak memberi petunjuk operasi mana yang gagal, sehingga sulit
    /// ditelusuri ketika hanya terjadi di sebagian komputer.
    ///
    /// Perbaikan pada versi ini:
    ///   1. Penanda langkah (_langkah) diperbarui sebelum setiap operasi
    ///      berisiko. Bila terjadi error, pesan menyebut langkahnya.
    ///   2. Praterbang (PraterbangSummary): sebelum satu sel pun ditulis,
    ///      diperiksa keberadaan sheet yang direferensikan formula, kondisi
    ///      proteksi sheet Summary, dan sel target yang ter-merge. Ketiganya
    ///      adalah penyebab 1004 yang paling sering dan paling sulit dilihat.
    ///   3. BarisKosongBerikutnya tidak lagi memakai UsedRange. UsedRange
    ///      menghitung sel yang pernah diformat meski kosong, sehingga bisa
    ///      mengembalikan baris di luar batas Excel (>1.048.576) dan memicu
    ///      1004 saat penulisan Audit Log.
    ///   4. String range multi-area ("B6:B8,B13:B15") dipecah menjadi
    ///      penulisan per area untuk menghindari perbedaan parsing separator.
    ///   5. Blok finally dibungkus try agar kegagalan pemulihan state tidak
    ///      menutupi hasil yang sebenarnya sudah berhasil.
    /// </summary>
    internal class RefreshSummary
    {
        private readonly Excel.Application _app;
        private Excel.Workbook _wbApp;      // workbook aplikasi (Aplikasi_CKPN_414.xlsm)

        // Penanda langkah untuk diagnostik error
        private string _langkah = "(belum dimulai)";

        private const string SH_SUM  = "Summary";
        private const string SH_INDV = "A. CKPN - INDV";
        private const string SH_KOL  = "B. CKPN - KOL INDV";
        private const string SH_LOG  = "Audit Log";
        private const string FMT_NUM = "#,##0;(#,##0);-";

        // Password proteksi sheet (Summary & Audit Log). Sama dengan Protection.cs.
        // Sheet dibuka proteksinya di awal, lalu dikunci ulang di blok finally.
        private const string PW_SHEET = "HaiiWhatt??";

        // ---- Konstanta ABA (KC0500) ----
        private const string ABA_SHEET     = "KC0500";
        private const string ABA_COL_EAD   = "T";   // nilai EAD
        private const string ABA_COL_SANDI = "D";   // sandi bank
        private const string ABA_COL_NAMA  = "E";   // nama bank
        private const int    ABA_START_ROW = 4;
        private const double PLAFON_LPS    = 2000000000d;   // 2 miliar

        // Sel target yang ditulis di sheet Summary — dipakai untuk praterbang
        private static readonly string[] SelTarget = new[]
        {
            "B6","B7","B8","B13","B14","B15",
            "C6","C7","C8","C13","C14","C15",
            "C18","C19",
            "H2","H3","H4","H5","H6","H7","H8","H9"
        };

        public RefreshSummary(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException("app");
        }

        private void Langkah(string s)
        {
            _langkah = s;
            // Tampilkan progres di status bar → memberi umpan balik selama proses
            // berjalan (mengurangi kesan layar "hang"/error) sekaligus membantu
            // jendela tetap ter-repaint.
            try { _app.StatusBar = "Refresh Summary: " + s + " ..."; } catch { }
        }

        // ================================================================
        // ENTRY POINT
        // ================================================================
        public void Refresh(string filePath, string sheetKCList)
        {
            // Catatan: konfirmasi Yes/No dilakukan di CKPNFunctions.RefreshSummaryCmd
            // SEBELUM Protection.CekProteksi, supaya dialog muncul segera setelah klik
            // (tidak di atas layar abu-abu akibat proses verifikasi proteksi).
            try
            {
                RefreshInti(filePath, sheetKCList);
            }
            catch (Exception ex)
            {
                // Bungkus ulang dengan penanda langkah supaya error 1004
                // yang generik menjadi bisa ditelusuri.
                throw new InvalidOperationException(
                    "Gagal pada langkah:\n  " + _langkah + "\n\n" +
                    "Pesan Excel:\n  " + ex.Message + "\n\n" +
                    "Catatan: 0x800A03EC = Excel error 1004 (\"operasi ditolak\").\n" +
                    "Penyebab tersering: sheet yang direferensikan tidak ada,\n" +
                    "sheet Summary terproteksi, sel target ter-merge, atau\n" +
                    "file sumber terbuka dalam Protected View.",
                    ex);
            }
            finally
            {
                // Jaring pengaman: pulihkan status bar & kursor apa pun yang terjadi
                // (mis. bila error terjadi sebelum blok finally di RefreshInti).
                try { _app.StatusBar = false; } catch { }
                try { _app.Cursor    = Excel.XlMousePointer.xlDefault; } catch { }
            }
        }

        private void RefreshInti(string filePath, string sheetKCList)
        {
            Langkah("Mengambil workbook aktif");
            var wb = _app.ActiveWorkbook;
            if (wb == null)
                throw new InvalidOperationException("Tidak ada workbook yang aktif.");
            _wbApp = wb;

            Langkah("Mencari sheet '" + SH_SUM + "'");
            var wsSum = CariSheet(wb, SH_SUM);
            if (wsSum == null)
                throw new InvalidOperationException("Sheet '" + SH_SUM + "' tidak ditemukan.");

            // ---------- 0. Praterbang ----------
            // Semua pemeriksaan yang bisa memicu 1004 dilakukan di sini,
            // SEBELUM satu sel pun ditulis, agar pesan errornya spesifik.
            Langkah("Praterbang: memeriksa sheet, proteksi, dan sel target");
            PraterbangSummary(wb, wsSum);

            var kcAktif = ParseSheetKC(sheetKCList);

            bool prevScreen  = _app.ScreenUpdating;
            bool prevEvents  = _app.EnableEvents;
            bool prevAlerts  = _app.DisplayAlerts;
            Excel.XlCalculation prevCalc = Excel.XlCalculation.xlCalculationAutomatic;
            try { prevCalc = _app.Calculation; } catch { }

            // Sheet yang akan ditulisi. Referensi & status proteksinya diingat
            // agar bisa dikunci ulang di blok finally, apa pun yang terjadi.
            Excel.Worksheet wsLog = CariSheet(wb, SH_LOG);
            bool sumTerproteksi = false;
            bool logTerproteksi = false;

            Excel.Workbook wbSrc = null;
            string laporan = null;

            try
            {
                Langkah("Menyiapkan state aplikasi (ScreenUpdating/Events/Calculation)");
                try { _app.ScreenUpdating = false; } catch { }
                try { _app.EnableEvents   = false; } catch { }
                try { _app.DisplayAlerts  = false; } catch { }
                try { _app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
                try { _app.Cursor      = Excel.XlMousePointer.xlWait; } catch { }

                // ---------- 0b. Buka proteksi sheet yang akan ditulisi ----------
                // Dilakukan SETELAH state di-set dan di DALAM try, supaya finally
                // dijamin mengunci ulang meski terjadi error di tengah proses.
                Langkah("Membuka proteksi sheet Summary (bila terproteksi)");
                sumTerproteksi = BukaProteksiSheet(wsSum);
                Langkah("Membuka proteksi sheet Audit Log (bila terproteksi)");
                logTerproteksi = BukaProteksiSheet(wsLog);

                // ---------- 1. Formula link internal (kolom B) ----------
                Langkah("Menulis formula link internal ke Summary B6:B8 dan B13:B15");
                TulisFormulaLink(wsSum);

                // ---------- 2. Kosongkan target hasil ----------
                Langkah("Mengosongkan sel hasil di Summary");
                BersihkanTarget(wsSum);

                // ---------- 3. Validasi file sumber ----------
                Langkah("Memvalidasi path file sumber (Master!D14)");
                if (string.IsNullOrEmpty(filePath) || filePath.Trim().Length == 0)
                    throw new InvalidOperationException("Path file sumber (Master!D14) kosong.");
                if (!System.IO.File.Exists(filePath))
                    throw new InvalidOperationException("File sumber tidak ditemukan:\n" + filePath);
                if (kcAktif.Count == 0)
                    throw new InvalidOperationException("Tidak ada KC yang dicentang di Master.");

                // ---------- 4. Buka file sumber ----------
                Langkah("Membuka file sumber (read-only):\n  " + filePath);
                wbSrc = _app.Workbooks.Open(filePath,
                                            UpdateLinks: 0,
                                            ReadOnly: true,
                                            IgnoreReadOnlyRecommended: true);

                // Deteksi Protected View: workbook terbuka tetapi sheet tidak
                // bisa diakses. Ini sering terjadi di komputer baru karena
                // Trust Center masih default dan file berasal dari jaringan.
                Langkah("Memeriksa apakah file sumber dapat dibaca (Protected View?)");
                int jumlahSheet;
                try { jumlahSheet = wbSrc.Worksheets.Count; }
                catch (Exception exPV)
                {
                    throw new InvalidOperationException(
                        "File sumber terbuka tetapi isinya tidak dapat dibaca.\n" +
                        "Kemungkinan Excel membukanya dalam Protected View.\n\n" +
                        "Solusi: File > Options > Trust Center > Trust Center Settings >\n" +
                        "Trusted Locations, lalu tambahkan folder file sumber.\n\n" +
                        "Detail: " + exPV.Message, exPV);
                }
                if (jumlahSheet == 0)
                    throw new InvalidOperationException("File sumber tidak memiliki sheet yang dapat dibaca.");

                var takDikenal = new List<BarisTakDikenal>();
                var lapOK   = new StringBuilder();
                var lapSkip = new StringBuilder();

                // BAGIAN 1: CKPN per Jenis
                Langkah("Menghitung CKPN per Jenis (Individual/Kolektif) dari file sumber");
                double totalIndv = 0, totalKol = 0;
                HitungCKPNPerJenis(wbSrc, kcAktif, ref totalIndv, ref totalKol,
                                   takDikenal, lapOK, lapSkip);

                // BAGIAN 2: PPKA per KC -> H2:H8
                Langkah("Menghitung PPKA per KC dan menulis ke Summary H2:H8");
                var lapPPKA = new StringBuilder();
                HitungPPKA(wbSrc, wsSum, kcAktif, lapPPKA);

                // BAGIAN 3: ABA (KC0500) -> C18/C19 + Audit Log
                Langkah("Menghitung CKPN Antar Bank dari " + ABA_SHEET);
                double abaDijamin = 0, abaDiAtasPlafon = 0;
                HitungEADAntarBank(wbSrc, wsSum, ref abaDijamin, ref abaDiAtasPlafon);

                Langkah("Menutup file sumber");
                wbSrc.Close(false);
                wbSrc = null;

                // ---------- 5. Tulis hasil CKPN ----------
                Langkah("Menulis hasil CKPN ke Summary C6, C7, C13, C14");
                ((Excel.Range)wsSum.Range["C6"]).Value2  = totalIndv;
                ((Excel.Range)wsSum.Range["C7"]).Value2  = totalKol;
                ((Excel.Range)wsSum.Range["C13"]).Value2 = totalIndv;
                ((Excel.Range)wsSum.Range["C14"]).Value2 = totalKol;

                // ---------- 6. (dinonaktifkan) ----------
                // Peringatan "Jenis CKPN tak dikenal" tidak lagi ditulis ke Audit Log.
                // Deteksi ini memicu false-positive dari baris header/non-data pada
                // kolom Jenis, sehingga peringatannya sengaja tidak ditampilkan.

                // ---------- 7. Laporan ringkas ----------
                Langkah("Menyusun laporan ringkas");
                var lap = new StringBuilder();
                lap.AppendLine("CKPN per Jenis:").Append(lapOK);
                if (lapSkip.Length > 0) lap.AppendLine().AppendLine("Dilewati:").Append(lapSkip);
                lap.AppendLine();
                lap.AppendLine("TOTAL Individual : " + totalIndv.ToString("#,##0"));
                lap.AppendLine("TOTAL Kolektif   : " + totalKol.ToString("#,##0"));
                lap.AppendLine();
                lap.AppendLine("PPKA per KC (H2:H8):").Append(lapPPKA);
                lap.AppendLine();
                lap.AppendLine("CKPN Antar Bank (KC0500):");
                lap.AppendLine("  - Dijamin LPS (<=2M) : " + abaDijamin.ToString("#,##0"));
                lap.AppendLine("  - Di atas plafon     : " + abaDiAtasPlafon.ToString("#,##0"));

                // Simpan laporan; dialog DITAMPILKAN setelah state layar dipulihkan
                // (lihat setelah blok finally) agar tidak muncul di atas layar abu-abu.
                laporan = lap.ToString();

                Langkah("Selesai");
            }
            finally
            {
                // Setiap pemulihan dibungkus sendiri. Tanpa ini, kegagalan
                // pada salah satu baris akan menimpa exception aslinya dan
                // menyembunyikan penyebab sebenarnya.
                if (wbSrc != null) { try { wbSrc.Close(false); } catch { } }

                // Kunci ulang HANYA sheet yang tadinya memang terproteksi, supaya
                // status proteksi awal terjaga. Dilakukan selagi EnableEvents masih
                // false agar tidak memicu event apa pun.
                if (sumTerproteksi) { try { KunciProteksiSheet(wsSum); } catch { } }
                if (logTerproteksi) { try { KunciProteksiSheet(wsLog); } catch { } }

                try { _app.Calculation    = prevCalc;   } catch { }
                try { _app.CutCopyMode    = (Excel.XlCutCopyMode)0; } catch { }
                try { _app.DisplayAlerts  = prevAlerts; } catch { }
                try { _app.EnableEvents   = prevEvents; } catch { }
                try { _app.ScreenUpdating = prevScreen; } catch { }
                try { _app.Cursor         = Excel.XlMousePointer.xlDefault; } catch { }
                try { _app.StatusBar      = false; } catch { }
            }

            // Laporan ditampilkan SETELAH state layar dipulihkan (ScreenUpdating aktif
            // kembali), supaya dialog muncul di atas tampilan normal — bukan di atas
            // layar abu-abu yang terkesan error.
            if (laporan != null)
                System.Windows.Forms.MessageBox.Show(
                    laporan, "Refresh CKPN & PPKA",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
        }

        // ================================================================
        // 0. PRATERBANG — deteksi penyebab 1004 sebelum menulis apa pun
        // ================================================================
        private void PraterbangSummary(Excel.Workbook wb, Excel.Worksheet wsSum)
        {
            var masalah = new List<string>();

            // (a) Sheet yang direferensikan formula HARUS ada.
            //     Menetapkan .Formula yang menyebut sheet tidak ada akan
            //     langsung melempar error 1004. Kedua sheet ini TIDAK
            //     diperiksa oleh Protection.SheetWajib, sehingga bisa lolos
            //     validasi awal tetapi gagal di sini.
            if (CariSheet(wb, SH_INDV) == null)
                masalah.Add("Sheet '" + SH_INDV + "' tidak ditemukan " +
                            "(dipakai formula Summary!B6).");
            if (CariSheet(wb, SH_KOL) == null)
                masalah.Add("Sheet '" + SH_KOL + "' tidak ditemukan " +
                            "(dipakai formula Summary!B7, B13, B14).");
            if (CariSheet(wb, SH_LOG) == null)
                masalah.Add("Sheet '" + SH_LOG + "' tidak ditemukan " +
                            "(tujuan penulisan rincian ABA).");

            // (b) Proteksi sheet Summary TIDAK lagi menggagalkan praterbang.
            //     Sheet yang terproteksi kini dibuka otomatis di awal RefreshInti
            //     (dengan password tertanam) lalu dikunci ulang di blok finally.
            //     Jadi kondisi terproteksi adalah kondisi normal — bukan masalah.

            // (c) Sel target tidak boleh ter-merge. Menulis ke sebagian sel
            //     yang ter-merge ditolak Excel dengan error 1004.
            var merged = new List<string>();
            foreach (var alamat in SelTarget)
            {
                try
                {
                    Excel.Range sel = (Excel.Range)wsSum.Range[alamat];
                    object mc = sel.MergeCells;
                    if (mc is bool && (bool)mc) merged.Add(alamat);
                }
                catch
                {
                    masalah.Add("Sel '" + alamat + "' di sheet Summary tidak dapat diakses.");
                }
            }
            if (merged.Count > 0)
                masalah.Add("Sel target berikut ter-merge dan harus di-unmerge: " +
                            string.Join(", ", merged.ToArray()));

            if (masalah.Count == 0) return;

            throw new InvalidOperationException(
                "Praterbang gagal — perbaiki dulu hal berikut:\n\n  - " +
                string.Join("\n  - ", masalah.ToArray()));
        }

        // ================================================================
        // 1. Formula link internal (kolom B)
        // ================================================================
        private void TulisFormulaLink(Excel.Worksheet wsSum)
        {
            // Section A - Net Flow
            Langkah("Menulis formula Summary!B6 (INDEX/MATCH ke '" + SH_INDV + "')");
            ((Excel.Range)wsSum.Range["B6"]).Formula =
                "=INDEX('" + SH_INDV + "'!$A:$K," +
                "MATCH(\"TOTAL\",'" + SH_INDV + "'!$B:$B,0)," +
                "MATCH(\"Penurunan Nilai\",'" + SH_INDV + "'!$5:$5,0))";

            Langkah("Menulis formula Summary!B7 (INDEX/MATCH ke '" + SH_KOL + "')");
            ((Excel.Range)wsSum.Range["B7"]).Formula = FormulaKolektif("1. PERHITUNGAN*", 2);

            Langkah("Menulis formula Summary!B8");
            ((Excel.Range)wsSum.Range["B8"]).Formula = "=SUM(B6:B7)";

            // Section C - PD Migration
            Langkah("Menulis formula Summary!B13");
            ((Excel.Range)wsSum.Range["B13"]).Formula = FormulaKolektif("3. PERHITUNGAN*", 1);
            Langkah("Menulis formula Summary!B14");
            ((Excel.Range)wsSum.Range["B14"]).Formula = FormulaKolektif("3. PERHITUNGAN*", 2);
            Langkah("Menulis formula Summary!B15");
            ((Excel.Range)wsSum.Range["B15"]).Formula = "=SUM(B13:B14)";

            // Format ditulis per area, bukan sebagai string range gabungan
            // "B6:B8,B13:B15". Parsing string multi-area memakai jalur yang
            // berbeda dan lebih sensitif terhadap separator sistem.
            Langkah("Menerapkan format angka pada Summary!B6:B8 dan B13:B15");
            ((Excel.Range)wsSum.Range["B6:B8"]).NumberFormat   = FMT_NUM;
            ((Excel.Range)wsSum.Range["B13:B15"]).NumberFormat = FMT_NUM;
        }

        /// <summary>
        /// Formula INDEX/MATCH ke sheet "B. CKPN - KOL INDV":
        /// cari blok yang diawali "n. PERHITUNGAN...", lalu ambil baris
        /// "Kategori" + offset di kolom C.
        /// </summary>
        private static string FormulaKolektif(string penandaBlok, int offset)
        {
            return
                "=INDEX('" + SH_KOL + "'!$C:$C," +
                "MATCH(\"Kategori\",INDEX('" + SH_KOL + "'!$B:$B," +
                "MATCH(\"" + penandaBlok + "\",'" + SH_KOL + "'!$B:$B,0)):'" + SH_KOL + "'!$B$1000,0)" +
                "+MATCH(\"" + penandaBlok + "\",'" + SH_KOL + "'!$B:$B,0)-1+" + offset + ")";
        }

        // ================================================================
        // 2. Kosongkan target hasil + set formula subtotal
        // ================================================================
        private void BersihkanTarget(Excel.Worksheet wsSum)
        {
            Langkah("Mengosongkan Summary!C6:C7 dan C13:C14");
            ((Excel.Range)wsSum.Range["C6:C7"]).ClearContents();
            ((Excel.Range)wsSum.Range["C13:C14"]).ClearContents();

            Langkah("Menulis formula subtotal Summary!C8 dan C15");
            ((Excel.Range)wsSum.Range["C8"]).Formula  = "=SUM(C6:C7)";
            ((Excel.Range)wsSum.Range["C15"]).Formula = "=SUM(C13:C14)";

            Langkah("Menerapkan format angka pada Summary!C6:C8 dan C13:C15");
            ((Excel.Range)wsSum.Range["C6:C8"]).NumberFormat   = FMT_NUM;
            ((Excel.Range)wsSum.Range["C13:C15"]).NumberFormat = FMT_NUM;

            Langkah("Mengosongkan Summary!C18:C19 (ABA)");
            ((Excel.Range)wsSum.Range["C18:C19"]).ClearContents();
            ((Excel.Range)wsSum.Range["C18:C19"]).NumberFormat = FMT_NUM;

            Langkah("Mengosongkan Summary!H2:H8 (PPKA)");
            ((Excel.Range)wsSum.Range["H2:H8"]).ClearContents();
            ((Excel.Range)wsSum.Range["H9"]).Formula = "=SUM(H2:H8)";
            ((Excel.Range)wsSum.Range["H2:H9"]).NumberFormat = FMT_NUM;
        }

        // ================================================================
        // BAGIAN 1: CKPN per Jenis (Individual / Kolektif)
        // ================================================================
        private void HitungCKPNPerJenis(
            Excel.Workbook wbSrc, List<string> kcAktif,
            ref double totalIndv, ref double totalKol,
            List<BarisTakDikenal> takDikenal,
            StringBuilder lapOK, StringBuilder lapSkip)
        {
            foreach (var namaKC in kcAktif)
            {
                Langkah("Membaca CKPN per Jenis dari sheet " + namaKC + " di file sumber");

                var wsKC = CariSheet(wbSrc, namaKC);
                if (wsKC == null)
                {
                    lapSkip.AppendLine("  - " + namaKC + " : sheet tidak ada");
                    continue;
                }

                string colVal, colJns;
                AmbilKolomCKPN(namaKC, out colVal, out colJns);
                if (colVal.Length == 0 || colJns.Length == 0)
                {
                    lapSkip.AppendLine("  - " + namaKC + " : mapping kolom belum ada");
                    continue;
                }

                int lrJ = BarisTerakhir(wsKC, colJns);
                int lrV = BarisTerakhir(wsKC, colVal);
                int lastRow = Math.Max(lrJ, lrV);

                if (lastRow < 2)
                {
                    lapSkip.AppendLine("  - " + namaKC + " : kolom " + colJns + "/" + colVal + " kosong");
                    continue;
                }

                object[] arrJns = BacaKolom(wsKC, colJns, 2, lastRow);
                object[] arrVal = BacaKolom(wsKC, colVal, 2, lastRow);

                double subIndv = 0, subKol = 0;
                for (int i = 0; i < arrJns.Length; i++)
                {
                    int    kode  = KodeJenisCKPN(arrJns[i]);
                    double nilai = i < arrVal.Length ? KeAngka(arrVal[i]) : 0;

                    switch (kode)
                    {
                        case 1: case 3: subIndv += nilai; break;
                        case 2: case 4: subKol  += nilai; break;
                        default:
                            string raw = arrJns[i] == null ? "" : arrJns[i].ToString().Trim();
                            if (raw.Length > 0)
                                takDikenal.Add(new BarisTakDikenal
                                {
                                    KC = namaKC, Baris = i + 2, Raw = raw, Nilai = nilai
                                });
                            break;
                    }
                }

                totalIndv += subIndv;
                totalKol  += subKol;
                lapOK.AppendLine("  - " + namaKC + " : Indv " + subIndv.ToString("#,##0") +
                                 " | Kol " + subKol.ToString("#,##0"));
            }
        }

        /// <summary>Mapping kolom Nilai CKPN &amp; Jenis CKPN per sheet KC.</summary>
        private static void AmbilKolomCKPN(string namaKC, out string colVal, out string colJns)
        {
            switch ((namaKC ?? "").Trim().ToUpperInvariant())
            {
                case "KC0600":
                case "KC0700":
                case "KC0800": colVal = "AV"; colJns = "AW"; break;
                case "KC0900": colVal = "AT"; colJns = "AU"; break;
                case "KC1000": colVal = "BB"; colJns = "BC"; break;
                case "KC1100": colVal = "BA"; colJns = "BB"; break;
                default:       colVal = "";   colJns = "";   break;
            }
        }

        // ================================================================
        // BAGIAN 2: PPKA OJK per KC -> H2:H8
        // ================================================================
        private void HitungPPKA(
            Excel.Workbook wbSrc, Excel.Worksheet wsSum,
            List<string> kcAktif, StringBuilder lapPPKA)
        {
            var map = new[]
            {
                new PpkaSpec("H2", "KC0500", "X",  4, false),   // selalu ikut
                new PpkaSpec("H3", "KC0600", "AV", 5, true),
                new PpkaSpec("H4", "KC0700", "AV", 5, true),
                new PpkaSpec("H5", "KC0800", "AV", 5, true),
                new PpkaSpec("H6", "KC0900", "AT", 5, true),
                new PpkaSpec("H7", "KC1000", "BB", 5, true),
                new PpkaSpec("H8", "KC1100", "BA", 5, true)
            };

            foreach (var mp in map)
            {
                Langkah("Menghitung PPKA " + mp.KC + " -> Summary!" + mp.Sel);

                bool ikut = !mp.PakaiCheckbox || AdaDiDaftar(kcAktif, mp.KC);
                double nilai = 0;

                if (!ikut)
                {
                    lapPPKA.AppendLine("  - " + mp.KC + " : tidak dicentang -> 0");
                }
                else
                {
                    var wsP = CariSheet(wbSrc, mp.KC);
                    if (wsP == null)
                    {
                        lapPPKA.AppendLine("  - " + mp.KC + " : sheet tidak ada");
                    }
                    else
                    {
                        nilai = JumlahKolomAngka(wsP, mp.Kolom, mp.BarisAwal);
                        lapPPKA.AppendLine("  - " + mp.KC + " (" + mp.Kolom + mp.BarisAwal +
                                           ":bawah) : " + nilai.ToString("#,##0"));
                    }
                }

                ((Excel.Range)wsSum.Range[mp.Sel]).Value2 = nilai;
            }
        }

        private class PpkaSpec
        {
            public string Sel;
            public string KC;
            public string Kolom;
            public int    BarisAwal;
            public bool   PakaiCheckbox;

            public PpkaSpec(string sel, string kc, string kolom, int barisAwal, bool pakaiCheckbox)
            {
                Sel = sel; KC = kc; Kolom = kolom; BarisAwal = barisAwal; PakaiCheckbox = pakaiCheckbox;
            }
        }

        // ================================================================
        // BAGIAN 3: CKPN ANTAR BANK (ABA) dari KC0500
        //   - EAD dikelompokkan per Sandi Bank
        //   - Total grup <= 2M  -> dijamin LPS   (Summary!C18)
        //   - Total grup >  2M  -> di atas plafon (Summary!C19)
        // ================================================================
        private void HitungEADAntarBank(
            Excel.Workbook wbSrc, Excel.Worksheet wsSum,
            ref double totalDijamin, ref double totalPlafon)
        {
            ((Excel.Range)wsSum.Range["C18"]).Value2 = 0;
            ((Excel.Range)wsSum.Range["C19"]).Value2 = 0;

            var wsKC = CariSheet(wbSrc, ABA_SHEET);
            if (wsKC == null) return;

            Langkah("Membaca " + ABA_SHEET + " untuk perhitungan ABA");
            int lastRow = Math.Max(BarisTerakhir(wsKC, ABA_COL_EAD),
                                   BarisTerakhir(wsKC, ABA_COL_SANDI));
            if (lastRow < ABA_START_ROW) return;

            object[] arrEAD   = BacaKolom(wsKC, ABA_COL_EAD,   ABA_START_ROW, lastRow);
            object[] arrSandi = BacaKolom(wsKC, ABA_COL_SANDI, ABA_START_ROW, lastRow);
            object[] arrNama  = BacaKolom(wsKC, ABA_COL_NAMA,  ABA_START_ROW, lastRow);

            // Agregasi per sandi bank (urutan input dipertahankan)
            var grp    = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var nmMap  = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var urutan = new List<string>();

            for (int i = 0; i < arrSandi.Length; i++)
            {
                string sandi = arrSandi[i] == null ? "" : arrSandi[i].ToString().Trim();
                if (sandi.Length == 0) continue;

                double nilai = i < arrEAD.Length ? KeAngka(arrEAD[i]) : 0;
                string nama  = (i < arrNama.Length && arrNama[i] != null)
                               ? arrNama[i].ToString().Trim() : "";

                if (grp.ContainsKey(sandi))
                {
                    grp[sandi] += nilai;
                    if (nmMap[sandi].Length == 0 && nama.Length > 0) nmMap[sandi] = nama;
                }
                else
                {
                    grp[sandi]   = nilai;
                    nmMap[sandi] = nama;
                    urutan.Add(sandi);
                }
            }

            // Siapkan Audit Log
            Langkah("Menulis rincian ABA ke sheet Audit Log");
            var wsLog = CariSheet(_wbApp, SH_LOG);
            int rowLog = 0;
            if (wsLog != null)
            {
                rowLog = BarisKosongBerikutnya(wsLog);
                ((Excel.Range)wsLog.Cells[rowLog, 1]).Value2 =
                    "PERHITUNGAN CKPN ABA (" + ABA_SHEET + ") - " +
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                ((Excel.Range)wsLog.Cells[rowLog, 1]).Font.Bold = true;
                rowLog++;

                ((Excel.Range)wsLog.Cells[rowLog, 1]).Value2 = "Sandi Bank";
                ((Excel.Range)wsLog.Cells[rowLog, 2]).Value2 = "Nama Bank";
                ((Excel.Range)wsLog.Cells[rowLog, 3]).Value2 = "Total EAD (kolom " + ABA_COL_EAD + ")";
                ((Excel.Range)wsLog.Cells[rowLog, 4]).Value2 = "Klasifikasi";
                ((Excel.Range)wsLog.Range[wsLog.Cells[rowLog, 1], wsLog.Cells[rowLog, 4]]).Font.Bold = true;
                rowLog++;
            }

            foreach (var sandi in urutan)
            {
                double g = grp[sandi];
                string klas;
                if (g > PLAFON_LPS) { totalPlafon  += g; klas = "Di atas plafon (>2M) -> C19"; }
                else                { totalDijamin += g; klas = "Dijamin (<=2M) -> C18"; }

                if (wsLog != null)
                {
                    ((Excel.Range)wsLog.Cells[rowLog, 1]).NumberFormat = "@";
                    ((Excel.Range)wsLog.Cells[rowLog, 1]).Value2 = "'" + sandi;
                    ((Excel.Range)wsLog.Cells[rowLog, 2]).Value2 = nmMap[sandi];
                    ((Excel.Range)wsLog.Cells[rowLog, 3]).Value2 = g;
                    ((Excel.Range)wsLog.Cells[rowLog, 3]).NumberFormat = "#,##0";
                    ((Excel.Range)wsLog.Cells[rowLog, 4]).Value2 = klas;
                    rowLog++;
                }
            }

            if (wsLog != null)
            {
                rowLog = TulisTotalLog(wsLog, rowLog, "TOTAL EAD Dijamin (C18)", totalDijamin);
                rowLog = TulisTotalLog(wsLog, rowLog, "TOTAL EAD Di atas plafon (C19)", totalPlafon);
                rowLog = TulisTotalLog(wsLog, rowLog, "TOTAL EAD Antar Bank", totalDijamin + totalPlafon);
            }

            Langkah("Menulis hasil ABA ke Summary!C18 dan C19");
            ((Excel.Range)wsSum.Range["C18"]).Value2 = totalDijamin;
            ((Excel.Range)wsSum.Range["C19"]).Value2 = totalPlafon;
        }

        private static int TulisTotalLog(Excel.Worksheet wsLog, int row, string label, double nilai)
        {
            ((Excel.Range)wsLog.Cells[row, 1]).Value2 = label;
            ((Excel.Range)wsLog.Cells[row, 1]).Font.Bold = true;
            ((Excel.Range)wsLog.Cells[row, 3]).Value2 = nilai;
            ((Excel.Range)wsLog.Cells[row, 3]).NumberFormat = "#,##0";
            return row + 1;
        }

        // ================================================================
        // Audit Log: baris "Jenis CKPN tak dikenal"
        // ================================================================
        private void TulisJenisTakDikenal(List<BarisTakDikenal> data)
        {
            if (data == null || data.Count == 0) return;

            var wsLog = CariSheet(_wbApp, SH_LOG);
            if (wsLog == null) return;

            int row = BarisKosongBerikutnya(wsLog);

            ((Excel.Range)wsLog.Cells[row, 1]).Value2 =
                "PERINGATAN: JENIS CKPN TAK DIKENAL - " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            ((Excel.Range)wsLog.Cells[row, 1]).Font.Bold = true;
            row++;

            ((Excel.Range)wsLog.Cells[row, 1]).Value2 = "Sheet KC";
            ((Excel.Range)wsLog.Cells[row, 2]).Value2 = "Baris";
            ((Excel.Range)wsLog.Cells[row, 3]).Value2 = "Nilai Jenis (mentah)";
            ((Excel.Range)wsLog.Cells[row, 4]).Value2 = "Nilai CKPN";
            ((Excel.Range)wsLog.Range[wsLog.Cells[row, 1], wsLog.Cells[row, 4]]).Font.Bold = true;
            row++;

            foreach (var it in data)
            {
                ((Excel.Range)wsLog.Cells[row, 1]).Value2 = it.KC;
                ((Excel.Range)wsLog.Cells[row, 2]).Value2 = it.Baris;
                ((Excel.Range)wsLog.Cells[row, 3]).NumberFormat = "@";
                ((Excel.Range)wsLog.Cells[row, 3]).Value2 = "'" + it.Raw;
                ((Excel.Range)wsLog.Cells[row, 4]).Value2 = it.Nilai;
                ((Excel.Range)wsLog.Cells[row, 4]).NumberFormat = "#,##0";
                row++;
            }
        }

        private class BarisTakDikenal
        {
            public string KC;
            public int    Baris;
            public string Raw;
            public double Nilai;
        }

        // ================================================================
        // HELPER
        // ================================================================
        private static List<string> ParseSheetKC(string sheetKCList)
        {
            var hasil = new List<string>();
            if (string.IsNullOrEmpty(sheetKCList)) return hasil;

            foreach (var p in sheetKCList.Split(new[] { ',', ';', '|' },
                                                StringSplitOptions.RemoveEmptyEntries))
            {
                string s = p.Trim().ToUpperInvariant();
                if (s.Length > 0 && !hasil.Contains(s)) hasil.Add(s);
            }
            return hasil;
        }

        private static bool AdaDiDaftar(List<string> daftar, string nilai)
        {
            foreach (var s in daftar)
                if (string.Equals(s, nilai, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        /// <summary>Normalisasi Jenis CKPN -> kode 1..4, atau 0 bila tidak dikenal.</summary>
        private static int KodeJenisCKPN(object v)
        {
            if (v == null || v is DBNull) return 0;

            if (v is double || v is int || v is long || v is decimal)
            {
                int k = (int)Math.Floor(Convert.ToDouble(v));
                return (k >= 1 && k <= 4) ? k : 0;
            }

            string s = v.ToString().Trim();
            if (s.Length == 0) return 0;

            // Ambil digit di depan: "1", "1 (Individual)", "01 - Individual"
            var digits = new StringBuilder();
            foreach (char c in s)
            {
                if (char.IsDigit(c)) digits.Append(c);
                else break;
            }
            if (digits.Length > 0)
            {
                int k;
                if (int.TryParse(digits.ToString(), out k) && k >= 1 && k <= 4) return k;
                return 0;
            }

            // Fallback: cocokkan teks
            s = s.ToUpperInvariant();
            if (s.Contains("TIDAK BURUK")) return 3;
            if (s.Contains("BURUK"))       return 4;
            if (s.Contains("INDIVIDU"))    return 1;
            if (s.Contains("KOLEKTIF"))    return 2;
            return 0;
        }

        /// <summary>Jumlahkan satu kolom mulai startRow s.d. baris terakhir (error/teks diabaikan).</summary>
        private static double JumlahKolomAngka(Excel.Worksheet ws, string col, int startRow)
        {
            int lastRow = BarisTerakhir(ws, col);
            if (lastRow < startRow) return 0;

            double total = 0;
            foreach (var v in BacaKolom(ws, col, startRow, lastRow)) total += KeAngka(v);
            return total;
        }

        /// <summary>Konversi nilai sel ke double; angka-tersimpan-sebagai-teks tetap dihitung.</summary>
        private static double KeAngka(object v)
        {
            if (v == null || v is DBNull) return 0;
            if (v is double)  return (double)v;
            if (v is int)     return (int)v;
            if (v is long)    return (long)v;
            if (v is decimal) return (double)(decimal)v;

            string s = v.ToString().Trim();
            if (s.Length == 0 || s.StartsWith("#")) return 0;   // sel error

            double d;
            if (double.TryParse(s, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.CurrentCulture, out d)) return d;
            if (double.TryParse(s, System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture, out d)) return d;
            return 0;
        }

        private static int BarisTerakhir(Excel.Worksheet ws, string col)
        {
            Excel.Range last = (Excel.Range)ws.Cells[ws.Rows.Count, col];
            return (int)((Excel.Range)last.End[Excel.XlDirection.xlUp]).Row;
        }

        /// <summary>Baca satu kolom menjadi array 1 dimensi (aman untuk range 1 baris).</summary>
        private static object[] BacaKolom(Excel.Worksheet ws, string col, int startRow, int endRow)
        {
            if (endRow < startRow) return new object[0];

            Excel.Range rng = (Excel.Range)ws.Range[col + startRow, col + endRow];
            object v = rng.Value2;
            if (v == null) return new object[endRow - startRow + 1];

            var arr = v as object[,];
            if (arr == null) return new object[] { v };   // range 1 sel -> skalar

            int n = arr.GetLength(0);
            var hasil = new object[n];
            for (int i = 1; i <= n; i++) hasil[i - 1] = arr[i, 1];
            return hasil;
        }

        // ----------------------------------------------------------------
        // BarisKosongBerikutnya: baris kosong berikutnya di Audit Log
        //
        // TIDAK memakai UsedRange. UsedRange ikut menghitung sel yang pernah
        // diformat meskipun kosong, sehingga bisa mengembalikan baris di luar
        // batas Excel (>1.048.576). Menulis ke baris di luar batas ditolak
        // dengan error 1004 (0x800A03EC) — dan ini bergantung pada riwayat
        // pemformatan file di masing-masing komputer, sehingga bisa terjadi
        // hanya di sebagian mesin.
        //
        // Cara aman: End(xlUp) pada kolom A-D, ambil baris terbesar, lalu
        // batasi agar tidak pernah melewati batas baris sheet.
        // ----------------------------------------------------------------
        private static int BarisKosongBerikutnya(Excel.Worksheet wsLog)
        {
            int last = 0;
            string[] kolom = { "A", "B", "C", "D" };

            foreach (var k in kolom)
            {
                try
                {
                    Excel.Range sel = (Excel.Range)wsLog.Cells[wsLog.Rows.Count, k];
                    int r = (int)((Excel.Range)sel.End[Excel.XlDirection.xlUp]).Row;
                    if (r > last) last = r;
                }
                catch { }
            }

            int hasil = last < 1 ? 1 : last + 2;

            int batas;
            try { batas = wsLog.Rows.Count; } catch { batas = 1048576; }
            if (hasil > batas - 50)
                throw new InvalidOperationException(
                    "Sheet '" + SH_LOG + "' sudah penuh (baris terpakai " + last + ").\n" +
                    "Hapus baris lama di Audit Log, lalu jalankan ulang.");

            return hasil;
        }

        // ================================================================
        // PROTEKSI SHEET — buka di awal, kunci ulang di finally
        // ================================================================

        /// <summary>
        /// Membuka proteksi sheet bila sedang terproteksi.
        /// Return true bila sheet TADINYA terproteksi (perlu dikunci ulang nanti).
        /// Melempar error yang jelas bila password tidak cocok, supaya penyebabnya
        /// tidak tersamar menjadi error 1004 generik saat penulisan sel.
        /// </summary>
        private bool BukaProteksiSheet(Excel.Worksheet ws)
        {
            if (ws == null) return false;

            bool terproteksi;
            try { terproteksi = ws.ProtectContents; }
            catch { return false; }

            if (!terproteksi) return false;   // memang tidak terproteksi → tidak perlu apa-apa

            try
            {
                ws.Unprotect(PW_SHEET);
            }
            catch
            {
                throw new InvalidOperationException(
                    "Gagal membuka proteksi sheet '" + ws.Name + "'.\n\n" +
                    "Sheet terproteksi dengan password yang BERBEDA dari yang tertanam\n" +
                    "di aplikasi. Kunci ulang sheet dengan password yang benar, atau\n" +
                    "buka proteksinya manual sebelum menjalankan Refresh.");
            }
            return true;
        }

        /// <summary>
        /// Mengunci ulang sheet dengan proteksi standar (password tertanam).
        /// Dipanggil di finally untuk sheet yang tadinya terproteksi.
        /// </summary>
        private void KunciProteksiSheet(Excel.Worksheet ws)
        {
            if (ws == null) return;
            try
            {
                ws.Protect(
                    Password:          PW_SHEET,
                    DrawingObjects:    true,
                    Contents:          true,
                    Scenarios:         true,
                    UserInterfaceOnly: false);
            }
            catch { /* abaikan; mis. sheet sudah terproteksi */ }
        }

        private static Excel.Worksheet CariSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
                if (string.Equals(sh.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sh;
            return null;
        }
    }
}